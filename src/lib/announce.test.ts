/**
 * Run with: npx tsx --test src/lib/announce.test.ts
 */

import { test } from "node:test";
import assert from "node:assert/strict";
import { announcementFor, formatRange } from "./announce.js";
import type { OooRequest } from "./oooRequest.js";

function request(over: Partial<OooRequest> = {}): OooRequest {
  return {
    pageId: "page-1",
    pageUrl: "https://app.notion.com/p1",
    personName: "Daniel Deserto",
    calendarName: "Danny",
    personEmail: "daniel.deserto@bluelabellabs.com",
    approverName: null,
    notes: null,
    startDate: "2026-09-10",
    endDate: "2026-09-14",
    status: "Requested",
    eventId: null,
    notifiedStatus: null,
    currentTitle: "✈️ Danny",
    hasIdentity: true,
    inTrash: false,
    ...over,
  };
}

test("formatRange collapses a range within one month", () => {
  assert.equal(formatRange("2026-09-10", "2026-09-14"), "Sep 10–14");
});

test("formatRange keeps both months when it crosses one", () => {
  assert.equal(formatRange("2026-09-30", "2026-10-02"), "Sep 30 – Oct 2");
});

test("formatRange shows a single day once, not twice", () => {
  assert.equal(formatRange("2026-09-10", "2026-09-10"), "Sep 10");
});

test("formatRange adds years only when the range crosses one", () => {
  assert.equal(formatRange("2026-12-30", "2027-01-02"), "Dec 30, 2026 – Jan 2, 2027");
});

test("a submission names the person, the dates, and links the row", () => {
  const a = announcementFor(request(), false);
  assert.match(a!.text, /\*Danny\* requested time off, Sep 10–14/);
  assert.match(a!.text, /<https:\/\/app\.notion\.com\/p1\|the request>/);
  assert.equal(a!.status, "Requested");
});

test("no announcement when the status matches what was last announced", () => {
  assert.equal(announcementFor(request({ status: "Approved", notifiedStatus: "Approved" }), true), null);
});

test("no announcement for a row with no status at all", () => {
  assert.equal(announcementFor(request({ status: null }), false), null);
});

test("reverting to Requested reads as a reversal, not a fresh submission", () => {
  const a = announcementFor(request({ status: "Requested", notifiedStatus: "Approved" }), true);
  assert.match(a!.text, /back to Requested/);
  assert.match(a!.text, /Removed from the team calendar/);
});

test("dates not yet filled in are said plainly rather than rendered as junk", () => {
  const a = announcementFor(request({ startDate: null, endDate: null }), false);
  assert.match(a!.text, /dates not set/);
});

test("an unrecognized status is recorded but not announced", () => {
  const a = announcementFor(request({ status: "Archived" }), false);
  assert.equal(a!.text, "");
  assert.equal(a!.status, "Archived");
});
