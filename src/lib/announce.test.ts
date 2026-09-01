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
    type: "Paid Time Off",
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
  assert.match(a!.text, /<https:\/\/app\.notion\.com\/p1\|details>/);
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

// --- travel reads differently from leave ---

const TRAVEL = "Work Related Travel";

test("scheduled travel reads as added, never approved", () => {
  const a = announcementFor(request({ type: TRAVEL, status: "Scheduled", notifiedStatus: null }), false);
  assert.match(a!.text, /\*Danny\* added work related travel, Sep 10–14/);
  assert.doesNotMatch(a!.text, /approved/i);
  assert.equal(a!.status, "Scheduled");
});

test("Scheduled says nothing twice, same as any other status", () => {
  assert.equal(announcementFor(request({ type: TRAVEL, status: "Scheduled", notifiedStatus: "Scheduled" }), true), null);
});

test("approved leave still reads as approved", () => {
  const a = announcementFor(request({ status: "Approved", notifiedStatus: "Requested" }), false);
  assert.match(a!.text, /approved and on the team calendar/);
});

test("travel that is stopped reads as cancelled, not denied", () => {
  const a = announcementFor(request({ type: TRAVEL, status: "Denied", notifiedStatus: "Approved" }), true);
  assert.match(a!.text, /work related travel for Sep 10–14 was cancelled/);
  assert.doesNotMatch(a!.text, /denied/i);
});

// --- the note line ---

test("a note is appended on its own line, under the message", () => {
  const a = announcementFor(request({ notes: "Back online Monday" }), false);
  assert.match(a!.text, /requested time off, Sep 10–14\. <[^>]+>\nNote: Back online Monday$/);
});

test("no note means no extra line", () => {
  assert.doesNotMatch(announcementFor(request(), false)!.text, /\nNote:/);
  assert.doesNotMatch(announcementFor(request({ notes: "   " }), false)!.text, /\nNote:/);
});

test("the note rides along on every status, not just the submission", () => {
  for (const status of ["Approved", "Scheduled", "Denied"]) {
    const a = announcementFor(request({ status, notifiedStatus: "Pending", notes: "context" }), false);
    assert.match(a!.text, /\nNote: context$/, `status=${status}`);
  }
});

test("a note cannot inject Slack markup", () => {
  const a = announcementFor(request({ notes: "<!channel> & <https://evil|click>" }), false);
  assert.match(a!.text, /Note: &lt;!channel&gt; &amp; &lt;https:\/\/evil\|click&gt;$/);
});

test("a very long note is trimmed rather than flooding the channel", () => {
  const a = announcementFor(request({ notes: "x".repeat(500) }), false);
  const note = a!.text.split("\nNote: ")[1]!;
  assert.equal(note.length, 300);
  assert.ok(note.endsWith("…"));
});
