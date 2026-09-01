/**
 * Wires the pure reconciler to the real world: Microsoft Graph for the
 * calendar, `context.notion` for the `O365 Event ID` write-back.
 *
 * Both entry points (the webhook and the sweep) go through `reconcilePage`,
 * so there is exactly one code path that changes the calendar.
 */

import { Ooo } from "./schema.js";
import {
  P,
  peopleValue,
  readCreatedBy,
  readPeople,
  readString,
  retrievePage,
  updateProps,
  type Notion,
  type NotionPage,
} from "./notion.js";
import { identityFill, needsEmailLookup, type IdentityInputs } from "./identity.js";
import { toOooRequest, type OooRequest } from "./oooRequest.js";
import { reconcile, type ReconcileDeps, type ReconcileResult } from "./reconcile.js";
import { createEvent, deleteEvent, updateEvent } from "./graphCalendar.js";
import { isNotFound } from "./errors.js";
import { disambiguateFirstNames, isDryRun, notifyChannel } from "./env.js";
import { notifyOps, postSlackMessage } from "./slack.js";

export interface LiveOptions {
  /**
   * Suppress the ops notice. The hourly sweep re-visits every blocked row, so
   * without this an approved-but-dateless row would page the channel once an
   * hour forever. The webhook still alerts once, when it happens.
   */
  quiet?: boolean;
}

export function liveDeps(notion: Notion, options: LiveOptions = {}): ReconcileDeps {
  return {
    createEvent: (request) => createEvent(request).then((event) => ({ id: event.id })),
    updateEvent: (eventId, request) => updateEvent(eventId, request).then(() => undefined),
    deleteEvent,
    storeEventId: (pageId, eventId) => updateProps(notion, pageId, { [Ooo.O365_EVENT_ID]: P.richText(eventId) }),
    clearEventId: (pageId) => updateProps(notion, pageId, { [Ooo.O365_EVENT_ID]: P.clearRichText() }),
    setTitle: (pageId, title) => updateProps(notion, pageId, { [Ooo.TITLE]: P.title(title) }),
    announce: (text) => postSlackMessage(notifyChannel(), text),
    setNotifiedStatus: (pageId, status) =>
      updateProps(notion, pageId, { [Ooo.NOTIFIED_STATUS]: P.richText(status) }),
    isNotFound,
    notify: options.quiet ? undefined : notifyOps,
    dryRun: isDryRun(),
  };
}

/**
 * Finds the Notion user whose account email matches `email`, or null.
 *
 * The workspace user list is small (tens of entries) and changes rarely, so it
 * is cached for a few minutes. `needsEmailLookup` keeps this off the hot path:
 * only an unattributed row with an address to match on gets this far.
 */
let userCache: { at: number; byEmail: Map<string, string> } | null = null;
const USER_CACHE_MS = 5 * 60 * 1000;

async function findUserIdByEmail(notion: Notion, email: string): Promise<string | null> {
  const key = email.trim().toLowerCase();
  if (!key) return null;
  if (!userCache || Date.now() - userCache.at > USER_CACHE_MS) {
    const byEmail = new Map<string, string>();
    try {
      let cursor: string | undefined;
      do {
        const res = (await notion.users.list({ start_cursor: cursor, page_size: 100 } as Parameters<
          Notion["users"]["list"]
        >[0])) as {
          results: Array<{ id?: string; type?: string; person?: { email?: string } }>;
          has_more?: boolean;
          next_cursor?: string | null;
        };
        for (const u of res.results ?? []) {
          const addr = u?.type === "person" ? u.person?.email : undefined;
          if (u?.id && addr) byEmail.set(addr.toLowerCase(), u.id);
        }
        cursor = res.has_more ? (res.next_cursor ?? undefined) : undefined;
      } while (cursor);
    } catch (err) {
      // A lookup failure must not stop the calendar work; the email fallback
      // still produces a correct name.
      console.error("[identity] couldn't list Notion users:", err);
      return null;
    }
    userCache = { at: Date.now(), byEmail };
  }
  return userCache.byEmail.get(key) ?? null;
}

/**
 * Fills in who a request belongs to, for rows submitted through a form.
 *
 * Returns the page with anything it wrote patched in locally, so the caller
 * derives the request from post-fill values without a second fetch. Never
 * overwrites a value a human set. Best-effort: a failure here is logged and
 * the reconcile continues, because a missing person link degrades the display
 * name, not the sync.
 */
export async function fillIdentity(notion: Notion, page: NotionPage): Promise<NotionPage> {
  const creator = readCreatedBy(page, Ooo.CREATED_BY);
  const inputs: IdentityInputs = {
    source: readString(page, Ooo.SOURCE),
    hasBlueLabeler: readPeople(page, Ooo.BLUELABELER).length > 0,
    email: readString(page, Ooo.EMAIL),
    createdBy: creator,
  };

  const matched = needsEmailLookup(inputs) ? await findUserIdByEmail(notion, inputs.email!) : null;
  const fill = identityFill(inputs, matched);
  if (!fill.blueLabelerId && !fill.email) return page;
  if (page.inTrash || isDryRun()) return page;

  const props: Record<string, unknown> = {};
  if (fill.email) props[Ooo.EMAIL] = { email: fill.email };
  if (fill.blueLabelerId) props[Ooo.BLUELABELER] = peopleValue([fill.blueLabelerId]);
  try {
    await updateProps(notion, page.id, props);
  } catch (err) {
    console.error(`[identity] couldn't fill ${page.id}:`, err);
    return page;
  }

  // Patch locally so the caller sees post-fill values without re-fetching.
  const properties = { ...page.properties };
  if (fill.email) properties[Ooo.EMAIL] = { type: "email", email: fill.email };
  if (fill.blueLabelerId) {
    properties[Ooo.BLUELABELER] = {
      type: "people",
      people: [{ object: "user", id: fill.blueLabelerId, name: creator?.name ?? null, person: { email: fill.email ?? null } }],
    };
  }
  console.log(`[identity] filled ${page.id}: ${Object.keys(props).join(", ")}`);
  return { ...page, properties };
}

/** Reconciles an already-loaded page. */
export function reconcileRequest(notion: Notion, request: OooRequest, options: LiveOptions = {}): Promise<ReconcileResult> {
  return reconcile(request, liveDeps(notion, options));
}

/**
 * Re-fetches the page, then reconciles.
 *
 * Always re-fetch. Notion's "Send webhook" automation payload carries only a
 * page reference with an EMPTY `properties` object, and a webhook event is a
 * signal that something changed, not a snapshot of what it changed to.
 */
export async function reconcilePage(notion: Notion, pageId: string): Promise<{ page: NotionPage; result: ReconcileResult }> {
  const fetched = await retrievePage(notion, pageId);
  // Identity first: the calendar subject is derived from it.
  const page = await fillIdentity(notion, fetched);
  const request = toOooRequest(page, disambiguateFirstNames());
  return { page, result: await reconcile(request, liveDeps(notion)) };
}
