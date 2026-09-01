/**
 * The shared-mailbox calendar operations this worker needs, and the event
 * payload it writes.
 *
 * Target is a SHARED MAILBOX, so the resource is
 * `/users/{smtp}/calendar/events`. It is deliberately not the Microsoft 365
 * group behind the Team: Graph lists Application permissions as "Not
 * supported" for `/groups/{id}/calendar/events`, so app-only auth can never
 * reach a group calendar regardless of what is granted. A shared mailbox is
 * also the only thing an Exchange Application Access Policy can scope
 * `Calendars.ReadWrite` down to. See README → Platform notes.
 */

import { graphRequest } from "./graphClient.js";
import { o365CalendarMailbox, oooTimezone } from "./env.js";
import { allDayRange, eventSubject, type OooRequest } from "./oooRequest.js";
import { slackEscape } from "./slack.js";
import { normalizeId } from "./notion.js";

/**
 * Named MAPI property carrying the Notion page id on every event we create.
 * This is the link that lets the sweep walk the calendar and ask "does the
 * Notion row behind this event still exist and still say Approved?" — which is
 * how deleted/trashed pages get their events cleaned up, since Notion's
 * "Send webhook" automations have no delete trigger to subscribe to.
 *
 * The GUID is arbitrary but must never change: events tagged with the old one
 * would stop being recognized as ours.
 */
export const NOTION_PAGE_ID_PROP = "String {c9c8e1a2-6f5b-4a3d-9e2f-7b1d0a4c8e35} Name NotionPageId";

export interface GraphEvent {
  id: string;
  subject?: string;
  isAllDay?: boolean;
  showAs?: string;
  /** "singleInstance" | "occurrence" | "exception" | "seriesMaster". */
  type?: string;
  seriesMasterId?: string;
  organizer?: { emailAddress?: { name?: string; address?: string } };
  attendees?: Array<{ emailAddress?: { name?: string; address?: string } }>;
  start?: { dateTime?: string; timeZone?: string };
  end?: { dateTime?: string; timeZone?: string };
  body?: { content?: string };
  singleValueExtendedProperties?: Array<{ id?: string; value?: string }>;
}

function calendarPath(suffix: string): string {
  return `users/${encodeURIComponent(o365CalendarMailbox())}/calendar/${suffix}`;
}

function eventsPath(): string {
  return calendarPath("events");
}

/**
 * The event body. Deliberately:
 *  - `isAllDay: true` with midnight-to-midnight bounds (`end` exclusive).
 *  - `showAs: "free"` so subscribing to the calendar never makes a viewer
 *    look busy or away themselves.
 *  - NO attendees. The event lives on the shared group calendar, which every
 *    member already sees; adding the requester as an attendee would mail them
 *    an invite and copy the event onto their personal calendar too.
 *  - `isReminderOn: false` — nobody needs a popup for a colleague's vacation.
 *  - The row's `Notes` are deliberately NOT included. The calendar grants
 *    `Default: Reviewer`, so every person in the tenant can read the body, and
 *    Notes is free text on a form — "surgery", "family emergency", whatever
 *    someone types. It stays in Notion, where approvers see it. Same reasoning
 *    that keeps a Type out of the subject line (see TIME_OFF_LABEL).
 */
export function buildEventPayload(request: OooRequest): Record<string, unknown> {
  if (!request.startDate || !request.endDate) {
    throw new Error(`Cannot build an event for ${request.pageId}: missing dates.`);
  }
  const range = allDayRange(request.startDate, request.endDate);
  const tz = oooTimezone();

  const approver = request.approverName ? `<p>Approved by ${slackEscape(request.approverName)}.</p>` : "";

  return {
    subject: eventSubject(request),
    isAllDay: true,
    start: { dateTime: range.start, timeZone: tz },
    end: { dateTime: range.end, timeZone: tz },
    // "free", not "oof": the shared calendar is a team-visibility board, and
    // marking it Out of Office would push an away status onto anyone who
    // subscribes to it.
    showAs: "free",
    isReminderOn: false,
    attendees: [],
    body: {
      contentType: "HTML",
      content:
        `<p>Time off requested in Notion.</p>${approver}` +
        `<p><a href="${request.pageUrl}">Open the request</a></p>` +
        `<p><i>Managed by notion-worker-ooo-calendar. Edits made here are overwritten from Notion.</i></p>`,
    },
    singleValueExtendedProperties: [{ id: NOTION_PAGE_ID_PROP, value: request.pageId }],
  };
}

export async function createEvent(request: OooRequest): Promise<GraphEvent> {
  return graphRequest<GraphEvent>(eventsPath(), { method: "POST", body: buildEventPayload(request) });
}

export async function updateEvent(eventId: string, request: OooRequest): Promise<GraphEvent> {
  return graphRequest<GraphEvent>(`${eventsPath()}/${encodeURIComponent(eventId)}`, {
    method: "PATCH",
    body: buildEventPayload(request),
  });
}

/** Deletes an event. A 404 propagates — callers treat it as "already gone". */
export async function deleteEvent(eventId: string): Promise<void> {
  await graphRequest(`${eventsPath()}/${encodeURIComponent(eventId)}`, { method: "DELETE" });
}

export async function getEvent(eventId: string): Promise<GraphEvent> {
  return graphRequest<GraphEvent>(`${eventsPath()}/${encodeURIComponent(eventId)}`, {
    query: { $select: "id,subject,start,end" },
  });
}

/**
 * Reads the Notion page id off an event, from the tag or (as a fallback) the
 * body link. Returned normalized (undashed, lowercase) — the tag carries a
 * dashed UUID while the body link carries the undashed form.
 */
export function notionPageIdOf(event: GraphEvent): string | null {
  const tagged = (event.singleValueExtendedProperties ?? []).find((p) => p.id === NOTION_PAGE_ID_PROP)?.value;
  if (tagged && tagged.trim()) return normalizeId(tagged.trim());
  // Fallback: the body always links back to the page, so an event whose
  // extended property got stripped (e.g. copied by a user) is still traceable.
  // Both hosts matter — Notion serves page URLs as app.notion.com in some
  // workspaces and www.notion.so in others, and matching only one silently
  // disables this fallback.
  const match = /notion\.(?:so|com)\/(?:[^/\s"]*-)?([0-9a-fA-F]{32})/.exec(event.body?.content ?? "");
  return match ? match[1]!.toLowerCase() : null;
}

/** One worker-owned calendar event paired with the row it came from. */
export interface TaggedEvent {
  /** NORMALIZED (undashed, lowercase) Notion page id. */
  pageId: string;
  event: GraphEvent;
}

/**
 * Every event this worker owns in `[startIso, endIso)`.
 *
 * Returns a LIST, not a map keyed by page id. Two events can carry the same
 * page id — that is exactly what a duplicate is — and a map would silently
 * drop all but one, making duplicates structurally invisible to the sweep that
 * exists to clean them up.
 *
 * Uses `calendarView`, which is the window-scoped read. It expands recurring
 * series into occurrences; ours are always single instances, and anything
 * untagged (an event a person added to the mailbox) is skipped, so expansion
 * only costs extra rows to filter, never a wrong deletion.
 */
export async function listWorkerEvents(startIso: string, endIso: string): Promise<TaggedEvent[]> {
  const found: TaggedEvent[] = [];
  let path: string | null = calendarPath("calendarView");
  let query: Record<string, string> | undefined = {
    startDateTime: startIso,
    endDateTime: endIso,
    $select: "id,subject,start,end,body,isAllDay,showAs",
    $expand: `singleValueExtendedProperties($filter=id eq '${NOTION_PAGE_ID_PROP}')`,
    $top: "100",
  };

  while (path) {
    const page: { value?: GraphEvent[]; "@odata.nextLink"?: string } = await graphRequest(path, {
      query,
      headers: { Prefer: `outlook.timezone="${oooTimezone()}"` },
    });
    for (const event of page.value ?? []) {
      const pageId = notionPageIdOf(event);
      // Untagged events belong to humans. Never touch them.
      if (pageId && event.id) found.push({ pageId, event });
    }
    path = page["@odata.nextLink"] ?? null;
    query = undefined; // nextLink already carries every parameter.
  }
  return found;
}

/**
 * EVERY event on the calendar in `[startIso, endIso)`, worker-owned or not.
 *
 * `listWorkerEvents` deliberately skips untagged events so the sweep can never
 * delete something a person created. This one is its opposite and exists for
 * the one-time backfill: it surfaces the pre-existing, human-created entries so
 * they can be adopted into Notion. Read-only; nothing here mutates anything.
 */
export async function listAllEvents(startIso: string, endIso: string): Promise<GraphEvent[]> {
  const out: GraphEvent[] = [];
  let path: string | null = calendarPath("calendarView");
  let query: Record<string, string> | undefined = {
    startDateTime: startIso,
    endDateTime: endIso,
    $select: "id,subject,start,end,body,isAllDay,showAs,type,seriesMasterId,organizer,attendees",
    $expand: `singleValueExtendedProperties($filter=id eq '${NOTION_PAGE_ID_PROP}')`,
    $top: "100",
    $orderby: "start/dateTime",
  };

  while (path) {
    const page: { value?: GraphEvent[]; "@odata.nextLink"?: string } = await graphRequest(path, {
      query,
      headers: { Prefer: `outlook.timezone="${oooTimezone()}"` },
    });
    for (const event of page.value ?? []) if (event?.id) out.push(event);
    path = page["@odata.nextLink"] ?? null;
    query = undefined;
  }
  return out;
}
