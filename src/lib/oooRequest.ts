/**
 * Pure translation of an OOO Entries row into the shape the reconciler
 * and the Graph payload builder work with, plus the calendar-date arithmetic.
 *
 * Everything here is pure and unit-tested (oooRequest.test.ts) — no network,
 * no Notion client — which is where the sibling workers put the logic that is
 * easy to get subtly wrong.
 */

import { Ooo, ApprovalStatus, AWAY_MARKER } from "./schema.js";
import { readDateEnd, readDateStart, readPeople, readString, pageUrl, type NotionPage } from "./notion.js";

export interface OooRequest {
  pageId: string;
  pageUrl: string;
  /** Full display name, best effort. Used in logs and ops alerts, not on the calendar. */
  personName: string;
  /** What the calendar shows: a first name, or a full name when it's ambiguous. */
  calendarName: string;
  personEmail: string | null;
  approverName: string | null;
  /** Free text from the row's Notes column, carried into the event body. */
  notes: string | null;
  /** First day off, `YYYY-MM-DD`. Null when the row hasn't been filled in. */
  startDate: string | null;
  /** Last day off, INCLUSIVE, `YYYY-MM-DD`. Equals startDate for a single day. */
  endDate: string | null;
  /** Raw `Approval Status` value; null when unset. */
  status: string | null;
  /** Current contents of the worker-owned `O365 Event ID` column, or null. */
  eventId: string | null;
  /** The row's `Title` as it reads right now. */
  currentTitle: string | null;
  /** The `Status` value the worker last announced in Slack, or null. */
  notifiedStatus: string | null;
  /**
   * Whether we resolved a real person (BlueLabeler or an email). False means
   * the name came from the row title, and writing a computed title back would
   * feed on itself.
   */
  hasIdentity: boolean;
  /** Whether the page is in Notion's trash. */
  inTrash: boolean;
}

const YMD = /^\d{4}-\d{2}-\d{2}/;

/**
 * Notion date values are either a plain `YYYY-MM-DD` or a full ISO instant
 * (when someone enabled the time component). Time of day is meaningless for an
 * all-day out-of-office block, so we keep the calendar date and drop the rest.
 */
export function toDateOnly(value: string | null | undefined): string | null {
  if (!value) return null;
  const m = YMD.exec(value.trim());
  return m ? m[0]! : null;
}

/** Adds `days` to a `YYYY-MM-DD`, in UTC so no DST shift can move the date. */
export function addDays(ymd: string, days: number): string {
  const [y, m, d] = ymd.split("-").map(Number) as [number, number, number];
  const t = Date.UTC(y, m - 1, d) + days * 86_400_000;
  return new Date(t).toISOString().slice(0, 10);
}

export interface AllDayRange {
  /** Midnight on the first day off. */
  start: string;
  /** Midnight on the day AFTER the last day off — Graph's end is exclusive. */
  end: string;
}

/**
 * Converts an inclusive Notion date range into the half-open range Graph wants
 * for `isAllDay: true`. Graph requires both ends to be midnight and treats
 * `end` as exclusive, so a one-day vacation on the 3rd is 3rd → 4th. Getting
 * this wrong is the classic off-by-one that makes every event a day short.
 */
export function allDayRange(startDate: string, endDateInclusive: string): AllDayRange {
  return { start: `${startDate}T00:00:00`, end: `${addDays(endDateInclusive, 1)}T00:00:00` };
}

/** Splits an email's local part into name-ish words: "andon.keller" → ["Andon","Keller"]. */
function nameWordsFromEmail(email: string): string[] {
  const local = email.split("@")[0] ?? email;
  return local
    .split(/[._-]+/)
    .filter(Boolean)
    .map((part) => part.charAt(0).toUpperCase() + part.slice(1));
}

/** Local part of an email, title-cased loosely — a last-resort display name. */
function nameFromEmail(email: string): string {
  return nameWordsFromEmail(email).join(" ");
}

/**
 * Uppercases the first character only. Notion display names are whatever the
 * person typed into their profile, and at least one here is all-lowercase; a
 * lowercase name on a shared calendar reads as a bug. Leaves an
 * already-capitalized name untouched.
 */
function capitalizeFirst(name: string): string {
  return name.charAt(0).toUpperCase() + name.slice(1);
}

/**
 * What goes on the shared calendar. The BlueLabeler person's Notion display
 * name when there is one, otherwise the first name from the email. First names
 * only, because that is how the team refers to each other and a wall of full
 * names reads like a directory.
 *
 * A first name listed in `disambiguate` gets the full name instead — the
 * escape hatch for when two people share one. Kept as an explicit list rather
 * than inferred from the database: inferring would mean a name silently
 * changing form the day someone new requests time off, and would still miss
 * anyone who has not requested any yet.
 */
export function calendarNameFor(
  notionName: string | null,
  email: string | null,
  fallback: string,
  disambiguate: readonly string[] = [],
): string {
  const words = email ? nameWordsFromEmail(email) : [];
  // The person picker wins over the address: it holds what someone goes by.
  const preferred = notionName ?? words[0] ?? null;
  if (!preferred) return fallback;

  const first = preferred.split(/\s+/)[0] ?? preferred;
  if (!disambiguate.some((n) => n.toLowerCase() === first.toLowerCase())) return preferred;

  // Shared first name: add a last name if we have one to add.
  if (preferred.includes(" ")) return preferred; // the Notion name already has one
  const last = words[1];
  return last ? `${preferred} ${last}` : preferred;
}

/**
 * Reads an OOO Entries page into an OooRequest. Total — never throws.
 *
 * `disambiguate` is the list of first names shared by more than one person
 * (OOO_DISAMBIGUATE_FIRST_NAMES); those get a full name on the calendar.
 */
export function toOooRequest(page: NotionPage, disambiguate: readonly string[] = []): OooRequest {
  // Identity comes from the BlueLabeler person picker when it's set, falling
  // back to the address the request form collects. The fallback matters: rows
  // for people outside the Notion workspace never resolve to a person.
  const picked = readPeople(page, Ooo.BLUELABELER)[0] ?? null;
  const notionName = picked?.name?.trim() ? capitalizeFirst(picked.name.trim()) : null;
  const email = readString(page, Ooo.EMAIL) ?? picked?.email ?? null;
  const rawTitle = readString(page, Ooo.TITLE);
  // The worker writes the title back, so strip its own marker before using the
  // title as a name fallback — otherwise each pass would append another one.
  const title = rawTitle ? rawTitle.replace(new RegExp(`\\s*${AWAY_MARKER}\\s*$`), "").trim() || null : null;

  const personName = notionName ?? (email ? nameFromEmail(email) : null) ?? title ?? "Team member";

  // One date RANGE, not two columns. Notion leaves `end` empty for a single
  // day, which is a one-day request rather than a missing value.
  const startDate = toDateOnly(readDateStart(page, Ooo.DATES));
  const endDate = toDateOnly(readDateEnd(page, Ooo.DATES)) ?? startDate;

  return {
    pageId: page.id,
    pageUrl: pageUrl(page.id, page.url),
    personName,
    calendarName: calendarNameFor(notionName, email, title ?? personName, disambiguate),
    personEmail: email,
    approverName: readPeople(page, Ooo.APPROVER)[0]?.name ?? null,
    notes: readString(page, Ooo.NOTES),
    startDate,
    endDate,
    status: readString(page, Ooo.STATUS),
    eventId: readString(page, Ooo.O365_EVENT_ID),
    notifiedStatus: readString(page, Ooo.NOTIFIED_STATUS),
    currentTitle: rawTitle,
    hasIdentity: Boolean(notionName ?? email),
    inTrash: page.inTrash,
  };
}

/** True when this row says an event should exist on the shared calendar. */
export function isApproved(request: OooRequest): boolean {
  return !request.inTrash && request.status === ApprovalStatus.APPROVED;
}

/**
 * Why an approved row still can't become an event. Null when it can.
 *
 * This exists because of a real failure mode the Read.ai worker hit: a Notion
 * automation fires on the property that changed, so someone who sets
 * `Status = Approved` *before* filling in Dates triggers a delivery against a
 * half-filled row, and no further delivery ever comes. We refuse to create a
 * bogus event, and the sweep retries the row later.
 */
export function blockedReason(request: OooRequest): string | null {
  if (!request.startDate) return "no Dates set";
  if (!request.endDate) return "no Dates set";
  if (request.endDate < request.startDate) return `the Dates range ends (${request.endDate}) before it starts (${request.startDate})`;
  return null;
}

/** Calendar subject, e.g. "Andon ✈️". */
export function eventSubject(request: OooRequest): string {
  return `${request.calendarName} ${AWAY_MARKER}`;
}
