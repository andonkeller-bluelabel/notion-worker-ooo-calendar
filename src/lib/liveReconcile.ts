/**
 * Wires the pure reconciler to the real world: Microsoft Graph for the
 * calendar, `context.notion` for the `O365 Event ID` write-back.
 *
 * Both entry points (the webhook and the sweep) go through `reconcilePage`,
 * so there is exactly one code path that changes the calendar.
 */

import { Ooo } from "./schema.js";
import { P, retrievePage, updateProps, type Notion, type NotionPage } from "./notion.js";
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
  const page = await retrievePage(notion, pageId);
  const request = toOooRequest(page, disambiguateFirstNames());
  return { page, result: await reconcile(request, liveDeps(notion)) };
}
