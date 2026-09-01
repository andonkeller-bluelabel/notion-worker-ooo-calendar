/**
 * Run with: npx tsx --test src/lib/graphCalendar.test.ts
 *
 * The event payload is the part a reviewer can't eyeball against Graph docs at
 * a glance, and the part where a wrong field silently produces the wrong thing
 * on everyone's calendar. These pin the four decisions that matter.
 */

import { test } from "node:test";
import assert from "node:assert/strict";

process.env.O365_GROUP_ID = "group-object-id";
process.env.OOO_TIMEZONE = "America/New_York";

const { buildEventPayload, notionPageIdOf, NOTION_PAGE_ID_PROP } = await import("./graphCalendar.js");
import type { OooRequest } from "./oooRequest.js";

function request(over: Partial<OooRequest> = {}): OooRequest {
  return {
    pageId: "1dfe01f9-e346-471b-9f9a-bf55a80de647",
    pageUrl: "https://www.notion.so/1dfe01f9e346471b9f9abf55a80de647",
    personName: "Andon Keller",
    calendarName: "Andon",
    personEmail: "andon.keller@bluelabellabs.com",
    approverName: null,
    notes: null,
    startDate: "2026-09-07",
    endDate: "2026-09-11",
    status: "Approved",
    eventId: null,
    inTrash: false,
    ...over,
  };
}

test("the event is an all-day free block with no attendees", () => {
  const payload = buildEventPayload(request()) as Record<string, any>;
  assert.equal(payload.isAllDay, true);
  // "free", not "oof" — subscribers must not inherit an away status.
  assert.equal(payload.showAs, "free");
  assert.equal(payload.isReminderOn, false);
  // Attendees would mail the requester an invite and copy the event onto their
  // personal calendar — the shared group calendar is the only place it belongs.
  assert.deepEqual(payload.attendees, []);
});

test("the all-day range ends the day after the last day off", () => {
  const payload = buildEventPayload(request()) as Record<string, any>;
  assert.deepEqual(payload.start, { dateTime: "2026-09-07T00:00:00", timeZone: "America/New_York" });
  assert.deepEqual(payload.end, { dateTime: "2026-09-12T00:00:00", timeZone: "America/New_York" });
});

test("the subject is a first name plus the away marker", () => {
  assert.equal((buildEventPayload(request()) as Record<string, any>).subject, "✈️ Andon");
});

test("Notes NEVER reach the event body — the calendar is readable org-wide", () => {
  const payload = buildEventPayload(request({ notes: "Minor surgery, back Monday" })) as Record<string, any>;
  assert.doesNotMatch(payload.body.content, /surgery/i);
  assert.doesNotMatch(payload.subject, /surgery/i);
});

test("the event carries the Notion page id, so the sweep can trace it back", () => {
  const payload = buildEventPayload(request()) as Record<string, any>;
  assert.deepEqual(payload.singleValueExtendedProperties, [
    { id: NOTION_PAGE_ID_PROP, value: "1dfe01f9-e346-471b-9f9a-bf55a80de647" },
  ]);
});

test("building an event without dates throws rather than sending Graph junk", () => {
  assert.throws(() => buildEventPayload(request({ startDate: null })), /missing dates/);
});

test("notionPageIdOf reads the tag, normalized", () => {
  assert.equal(
    notionPageIdOf({ id: "e1", singleValueExtendedProperties: [{ id: NOTION_PAGE_ID_PROP, value: "1DFE01F9-E346-471B-9F9A-BF55A80DE647" }] }),
    "1dfe01f9e346471b9f9abf55a80de647",
  );
});

test("notionPageIdOf falls back to the body link when the tag was stripped", () => {
  assert.equal(
    notionPageIdOf({ id: "e1", body: { content: '<p><a href="https://www.notion.so/Andon-1dfe01f9e346471b9f9abf55a80de647">Open</a></p>' } }),
    "1dfe01f9e346471b9f9abf55a80de647",
  );
});

test("the body fallback handles app.notion.com, not just notion.so", () => {
  // This workspace serves page URLs as app.notion.com. Matching only notion.so
  // silently disabled the fallback and, with it, orphan cleanup.
  assert.equal(
    notionPageIdOf({ id: "e1", body: { content: '<p><a href="https://app.notion.com/1dfe01f9e346471b9f9abf55a80de647">Open</a></p>' } }),
    "1dfe01f9e346471b9f9abf55a80de647",
  );
});

test("an untagged event returns null, so the sweep leaves human events alone", () => {
  assert.equal(notionPageIdOf({ id: "e1", subject: "Team offsite", body: { content: "<p>see you there</p>" } }), null);
});
