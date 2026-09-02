/**
 * Mirroring the two holiday databases onto the shared calendar.
 *
 * Holidays are not requests: nobody submits or approves one, so they never
 * enter OOO Entries. Each holiday database carries its own worker-owned
 * `O365 Event ID`, exactly like OOO Entries, so the pattern is the same one
 * twice rather than two different ideas.
 *
 * Pure and unit-tested; the API calls live in the sync that drives this.
 */

import { US_HOLIDAY_MARKER, UsHolidays, VENDOR_HOLIDAY_MARKER, VendorHolidays } from "./schema.js";
import { toDateOnly } from "./oooRequest.js";
import { readDateEnd, readDateStart, readString, type NotionPage } from "./notion.js";

export type HolidayKind = "us" | "vendor";

export interface HolidaySource {
  kind: HolidayKind;
  titleProp: string;
  dateProp: string;
  eventIdProp: string;
  marker: string;
  /** Where to read rows from; empty disables this source. */
  dataSourceId: () => string;
}

export interface HolidayEntry {
  pageId: string;
  /** Calendar subject, e.g. "🇺🇸 Thanksgiving Day" or "🏢 Tatvasoft: Diwali". */
  subject: string;
  startDate: string;
  /** INCLUSIVE last day. Equals startDate for a single-day holiday. */
  endDate: string;
  /** Current contents of this row's `O365 Event ID`, or null. */
  eventId: string | null;
}

export function usSource(dataSourceId: () => string): HolidaySource {
  return {
    kind: "us",
    titleProp: UsHolidays.TITLE,
    dateProp: UsHolidays.DATE,
    eventIdProp: UsHolidays.O365_EVENT_ID,
    marker: US_HOLIDAY_MARKER,
    dataSourceId,
  };
}

export function vendorSource(dataSourceId: () => string): HolidaySource {
  return {
    kind: "vendor",
    titleProp: VendorHolidays.TITLE,
    dateProp: VendorHolidays.DATE,
    eventIdProp: VendorHolidays.O365_EVENT_ID,
    marker: VENDOR_HOLIDAY_MARKER,
    dataSourceId,
  };
}

/**
 * Reads a holiday row into the shape the mirror works with.
 *
 * Returns null for a row with no name or no date, which is an unfinished row
 * rather than an event. `subject` is null in that case too, so the caller can
 * still see the row's event id and clean up an event that should no longer
 * exist — a holiday whose date was cleared must come off the calendar.
 */
export function toHolidayEntry(page: NotionPage, source: HolidaySource): HolidayEntry | null {
  const title = readString(page, source.titleProp)?.trim();
  const startDate = toDateOnly(readDateStart(page, source.dateProp));
  const eventId = readString(page, source.eventIdProp);
  if (!title || !startDate) return null;
  const endDate = toDateOnly(readDateEnd(page, source.dateProp)) ?? startDate;
  if (endDate < startDate) return null;
  return { pageId: page.id, subject: `${source.marker} ${title}`, startDate, endDate, eventId };
}

/** Whether a row still has enough to be an event. Mirrors toHolidayEntry's guards. */
export function isUsable(page: NotionPage, source: HolidaySource): boolean {
  return toHolidayEntry(page, source) !== null;
}
