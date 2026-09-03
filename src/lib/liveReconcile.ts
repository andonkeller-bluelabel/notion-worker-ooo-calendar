/**
 * Wires the pure reconciler to the real world: Microsoft Graph for the
 * calendar, `context.notion` for the `O365 Event ID` write-back.
 *
 * Both entry points (the webhook and the sweep) go through `reconcilePage`,
 * so there is exactly one code path that changes the calendar.
 */

import { CALENDAR_STATUSES, Ooo } from "./schema.js";
import {
  P,
  peopleValue,
  queryAll,
  readCreatedBy,
  readPeople,
  readString,
  retrievePage,
  updateProps,
  type Notion,
  type NotionPage,
} from "./notion.js";
import { filledName, identityFill, needsEmailLookup, type IdentityInputs } from "./identity.js";
import { addDays, rangesOverlap, toOooRequest, type OooRequest } from "./oooRequest.js";
import { reconcile, type ReconcileDeps, type ReconcileResult } from "./reconcile.js";
import type { OverlapRow } from "./announce.js";
import { createEvent, deleteEvent, updateEvent } from "./graphCalendar.js";
import { isNotFound } from "./errors.js";
import { disambiguateFirstNames, isDryRun, notifyChannel, oooDataSourceId, processUrl } from "./env.js";
import { dmByEmail, notifyOps, postSlackMessage } from "./slack.js";

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
    setStatus: (pageId, status) => updateProps(notion, pageId, { [Ooo.STATUS]: { status: { name: status } } }),
    dm: dmByEmail,
    setNotifiedApprover: (pageId, approverId) =>
      updateProps(notion, pageId, { [Ooo.NOTIFIED_APPROVER]: P.richText(approverId) }),
    processUrl: processUrl(),
    findOverlaps: (request) => findOverlaps(notion, request),
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
let userCache: { at: number; byEmail: Map<string, { id: string; name: string | null }> } | null = null;
const USER_CACHE_MS = 5 * 60 * 1000;

async function findUserByEmail(notion: Notion, email: string): Promise<{ id: string; name: string | null } | null> {
  const key = email.trim().toLowerCase();
  if (!key) return null;
  if (!userCache || Date.now() - userCache.at > USER_CACHE_MS) {
    const byEmail = new Map<string, { id: string; name: string | null }>();
    try {
      let cursor: string | undefined;
      do {
        const res = (await notion.users.list({ start_cursor: cursor, page_size: 100 } as Parameters<
          Notion["users"]["list"]
        >[0])) as {
          results: Array<{ id?: string; name?: string; type?: string; person?: { email?: string } }>;
          has_more?: boolean;
          next_cursor?: string | null;
        };
        for (const u of res.results ?? []) {
          const addr = u?.type === "person" ? u.person?.email : undefined;
          if (u?.id && addr) byEmail.set(addr.toLowerCase(), { id: u.id, name: u.name ?? null });
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
    hasBlueLabeler: readPeople(page, Ooo.BLUELABELER).length > 0,
    email: readString(page, Ooo.EMAIL),
    createdBy: creator,
  };

  const matched = needsEmailLookup(inputs) ? await findUserByEmail(notion, inputs.email!) : null;
  const fill = identityFill(inputs, matched?.id ?? null);
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
  //
  // The NAME matters: it becomes the calendar subject. Take it from whoever we
  // actually filled from — the matched user on the email path, the creator on
  // the Created-by path. Using the creator's name on the email path put
  // "✈️ Anonymous" on a row whose BlueLabeler was correctly resolved to a real
  // person, because an anonymous form submission's creator is a sentinel user
  // literally named "Anonymous".
  const properties = { ...page.properties };
  if (fill.email) properties[Ooo.EMAIL] = { type: "email", email: fill.email };
  if (fill.blueLabelerId) {
    properties[Ooo.BLUELABELER] = {
      type: "people",
      people: [
        {
          object: "user",
          id: fill.blueLabelerId,
          name: filledName(fill.blueLabelerId, matched, creator),
          person: { email: fill.email ?? inputs.email ?? null },
        },
      ],
    };
  }
  console.log(`[identity] filled ${page.id}: ${Object.keys(props).join(", ")}`);
  return { ...page, properties };
}

/**
 * Other rows for the same person whose dates touch this one's.
 *
 * Queried narrowly — same email, in a calendar status, starting within a
 * couple of months — then filtered exactly in code, because Notion's date
 * filters compare against a range's start and cannot express "overlaps".
 *
 * Best-effort: a failure returns nothing, so a missed warning never blocks the
 * calendar work.
 */
export async function findOverlaps(notion: Notion, request: OooRequest): Promise<OverlapRow[]> {
  if (!request.personEmail || !request.startDate || !request.endDate) return [];
  try {
    const rows = await queryAll(notion, oooDataSourceId(), {
      and: [
        { property: Ooo.EMAIL, email: { equals: request.personEmail } },
        { or: CALENDAR_STATUSES.map((name) => ({ property: Ooo.STATUS, status: { equals: name } })) },
        { property: Ooo.DATES, date: { on_or_after: addDays(request.startDate, -90) } },
      ],
    });
    return rows
      .filter((row) => row.id !== request.pageId)
      .map((row) => toOooRequest(row))
      .filter(
        (other): other is OooRequest & { startDate: string; endDate: string } =>
          Boolean(other.startDate && other.endDate) &&
          rangesOverlap(request.startDate!, request.endDate!, other.startDate!, other.endDate!),
      )
      .map((other) => ({ pageUrl: other.pageUrl, startDate: other.startDate, endDate: other.endDate }));
  } catch (err) {
    console.error(`[overlap] couldn't check ${request.pageId}:`, err);
    return [];
  }
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
