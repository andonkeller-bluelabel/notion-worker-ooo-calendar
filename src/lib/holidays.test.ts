/**
 * Run with: npx tsx --test src/lib/holidays.test.ts
 */

import { test } from "node:test";
import assert from "node:assert/strict";
import { toHolidayEntry, usSource, vendorSource } from "./holidays.js";
import { HOLIDAY_PAGE_ID_PROP, NOTION_PAGE_ID_PROP } from "./graphCalendar.js";
import type { NotionPage } from "./notion.js";

const US = usSource(() => "us-ds");
const VENDOR = vendorSource(() => "vendor-ds");

function page(props: Record<string, unknown>): NotionPage {
  return { id: "hol-1", url: "https://app.notion.com/hol1", inTrash: false, properties: props };
}
const title = (s: string) => ({ type: "title", title: [{ plain_text: s }] });
const date = (start: string | null, end?: string | null) => ({
  type: "date",
  date: start ? { start, end: end ?? null } : null,
});
const richText = (s: string | null) => ({ type: "rich_text", rich_text: s ? [{ plain_text: s }] : [] });

test("a US holiday takes the flag marker", () => {
  const e = toHolidayEntry(page({ Holiday: title("Thanksgiving Day"), Date: date("2026-11-26") }), US);
  assert.equal(e!.subject, "🇺🇸 Thanksgiving Day");
  assert.equal(e!.startDate, "2026-11-26");
  assert.equal(e!.endDate, "2026-11-26", "a single-day holiday ends the day it starts");
});

test("a vendor holiday takes the building marker and keeps its own naming", () => {
  // Vendor rows already self-identify, so nothing is prefixed onto the name.
  const e = toHolidayEntry(page({ Name: title("Tatvasoft: Diwali"), Date: date("2026-11-08") }), VENDOR);
  assert.equal(e!.subject, "🏢 Tatvasoft: Diwali");
});

test("the vendor source reads its own Date property", () => {
  const e = toHolidayEntry(page({ Name: title("Vstorm: Independence Day"), Date: date("2026-11-11") }), VENDOR);
  assert.equal(e!.startDate, "2026-11-11");
});

test("a multi-day holiday keeps an inclusive end", () => {
  const e = toHolidayEntry(page({ Holiday: title("Diwali week"), Date: date("2026-11-08", "2026-11-12") }), US);
  assert.equal(e!.endDate, "2026-11-12");
});

test("an unfinished row produces nothing rather than a guess", () => {
  assert.equal(toHolidayEntry(page({ Holiday: title("No date yet"), Date: date(null) }), US), null);
  assert.equal(toHolidayEntry(page({ Holiday: title("  "), Date: date("2026-11-26") }), US), null);
  assert.equal(toHolidayEntry(page({ Holiday: title("Backwards"), Date: date("2026-11-26", "2026-11-20") }), US), null);
});

test("an existing event id is carried through", () => {
  const e = toHolidayEntry(
    page({ Holiday: title("Labor Day"), Date: date("2026-09-07"), "O365 Event ID": richText("EVT-9") }),
    US,
  );
  assert.equal(e!.eventId, "EVT-9");
});

// --- the lane separation that keeps the sweeps from eating each other ---

test("the holiday tag is NOT the time-off tag", () => {
  // Sharing one tag would make the OOO sweep delete every mirrored holiday on
  // its next run, since a holiday row has no Status to be Approved.
  assert.notEqual(HOLIDAY_PAGE_ID_PROP, NOTION_PAGE_ID_PROP);
});
