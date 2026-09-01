/**
 * Run with: npx tsx --test src/lib/oooRequest.test.ts
 *
 * Covers the date arithmetic (where the off-by-one lives) and the row-reading
 * that everything downstream depends on.
 */

import { test } from "node:test";
import assert from "node:assert/strict";
import { addDays, allDayRange, blockedReason, calendarNameFor, eventSubject, isApproved, toDateOnly, toOooRequest } from "./oooRequest.js";
import type { NotionPage } from "./notion.js";

function page(props: Record<string, unknown>, over: Partial<NotionPage> = {}): NotionPage {
  return { id: "1dfe01f9-e346-471b-9f9a-bf55a80de647", url: "https://notion.so/abc", inTrash: false, properties: props, ...over };
}

const person = (name: string | null) => ({
  type: "people",
  people: [{ object: "user", id: "u1", name, person: undefined }],
});
const email = (address: string | null) => ({ type: "email", email: address });
const date = (start: string | null, end?: string | null) => ({ type: "date", date: start ? { start, end: end ?? null } : null });
const status = (name: string | null) => ({ type: "status", status: name ? { name } : null });
const richText = (s: string | null) => ({ type: "rich_text", rich_text: s ? [{ plain_text: s }] : [] });
const title = (s: string) => ({ type: "title", title: [{ plain_text: s }] });

// Mirrors the live "OOO Entries" schema: the requester is an email column
// (the form collects it), and Status is a status-type property, not a select.
const bluelabeler = (name: string | null) => ({
  type: "people",
  people: name ? [{ object: "user", id: "u9", name, person: { email: "picked@bluelabellabs.com" } }] : [],
});

const APPROVED_ROW = {
  Title: title("Andon time off"),
  "Your Email": email("andon.keller@bluelabellabs.com"),
  Dates: date("2026-09-07", "2026-09-11"),
  Status: status("Approved"),
  Approver: person("Jane Approver"),
  Notes: richText(null),
  "O365 Event ID": richText(null),
};

// --- date helpers ---

test("toDateOnly keeps the calendar date and drops any time component", () => {
  assert.equal(toDateOnly("2026-09-07"), "2026-09-07");
  assert.equal(toDateOnly("2026-09-07T09:30:00.000-04:00"), "2026-09-07");
  assert.equal(toDateOnly(null), null);
  assert.equal(toDateOnly(""), null);
  assert.equal(toDateOnly("not a date"), null);
});

test("addDays crosses month, year, and leap-day boundaries", () => {
  assert.equal(addDays("2026-09-07", 1), "2026-09-08");
  assert.equal(addDays("2026-09-30", 1), "2026-10-01");
  assert.equal(addDays("2026-12-31", 1), "2027-01-01");
  assert.equal(addDays("2028-02-28", 1), "2028-02-29"); // 2028 is a leap year
  assert.equal(addDays("2026-03-01", -1), "2026-02-28");
});

test("addDays is immune to DST — the US spring-forward Sunday still advances one date", () => {
  // 2026-03-08 is the US DST transition. A local-time implementation lands on
  // the wrong day here; UTC arithmetic does not.
  assert.equal(addDays("2026-03-07", 1), "2026-03-08");
  assert.equal(addDays("2026-03-08", 1), "2026-03-09");
  assert.equal(addDays("2026-11-01", 1), "2026-11-02"); // fall-back Sunday
});

test("allDayRange makes Graph's exclusive end one day past the last day off", () => {
  assert.deepEqual(allDayRange("2026-09-07", "2026-09-11"), {
    start: "2026-09-07T00:00:00",
    end: "2026-09-12T00:00:00", // NOT 09-11: the 11th must still read as off
  });
});

test("allDayRange handles a single-day request", () => {
  assert.deepEqual(allDayRange("2026-09-07", "2026-09-07"), {
    start: "2026-09-07T00:00:00",
    end: "2026-09-08T00:00:00",
  });
});

// --- row reading ---

test("reads a complete approved row", () => {
  const request = toOooRequest(page(APPROVED_ROW));
  assert.equal(request.personName, "Andon Keller");
  assert.equal(request.personEmail, "andon.keller@bluelabellabs.com");
  assert.equal(request.approverName, "Jane Approver");
  assert.equal(request.startDate, "2026-09-07");
  assert.equal(request.endDate, "2026-09-11");
  assert.equal(request.status, "Approved");
  assert.equal(request.eventId, null);
  assert.ok(isApproved(request));
  assert.equal(eventSubject(request), "Andon ✈️");
});

test("the away marker is stripped from the title before it is used as a name fallback", () => {
  // The worker writes "<name> ✈️" back into Title. Without stripping, a row
  // with no identity would gain another marker on every pass.
  const noIdentity = { ...APPROVED_ROW, "Your Email": email(null), BlueLabeler: bluelabeler(null), Title: title("Andon ✈️") };
  const request = toOooRequest(page(noIdentity));
  assert.equal(request.calendarName, "Andon");
  assert.equal(request.hasIdentity, false);
});

test("BlueLabeler supplies the name when set, and is title-cased", () => {
  const withPicker = toOooRequest(page({ ...APPROVED_ROW, BlueLabeler: bluelabeler("Danny") }));
  assert.equal(withPicker.calendarName, "Danny");
  assert.equal(eventSubject(withPicker), "Danny ✈️");
  // A lowercase Notion profile name would otherwise read as a bug on the calendar.
  assert.equal(toOooRequest(page({ ...APPROVED_ROW, BlueLabeler: bluelabeler("parvathy") })).calendarName, "Parvathy");
});

test("falls back to the email when BlueLabeler is empty — unmatched people still sync", () => {
  const noPicker = toOooRequest(page({ ...APPROVED_ROW, BlueLabeler: bluelabeler(null) }));
  assert.equal(noPicker.calendarName, "Andon");
  assert.equal(noPicker.personEmail, "andon.keller@bluelabellabs.com");
});

test("derives the display name from the email local part", () => {
  assert.equal(toOooRequest(page({ ...APPROVED_ROW, "Your Email": email("jane.doe@bluelabellabs.com") })).personName, "Jane Doe");
  assert.equal(toOooRequest(page({ ...APPROVED_ROW, "Your Email": email("jane.doe@bluelabellabs.com") })).calendarName, "Jane");
  assert.equal(toOooRequest(page({ ...APPROVED_ROW, "Your Email": email("mmurphy@bluelabellabs.com") })).personName, "Mmurphy");
});

test("falls back to the row title when the form left the email blank", () => {
  assert.equal(toOooRequest(page({ ...APPROVED_ROW, "Your Email": email(null) })).personName, "Andon time off");
});

test("Notes are carried through for the event body", () => {
  assert.equal(toOooRequest(page({ ...APPROVED_ROW, Notes: richText("Back online Monday") })).notes, "Back online Monday");
  assert.equal(toOooRequest(page(APPROVED_ROW)).notes, null);
});

test("a range with no end is a ONE-DAY request, not a missing value", () => {
  // Notion leaves `end` empty when someone picks a single day in the picker.
  const request = toOooRequest(page({ ...APPROVED_ROW, Dates: date("2026-09-07") }));
  assert.equal(request.startDate, "2026-09-07");
  assert.equal(request.endDate, "2026-09-07");
  assert.equal(blockedReason(request), null);
});

test("reads both ends of the Dates range", () => {
  const request = toOooRequest(page({ ...APPROVED_ROW, Dates: date("2026-09-07", "2026-09-11") }));
  assert.equal(request.startDate, "2026-09-07");
  assert.equal(request.endDate, "2026-09-11");
});

test("the subject never names a reason — OOO Entries has no Type, deliberately", () => {
  // Adding "Sick" or "Personal" to a subject broadcasts it to everyone with
  // access to the shared calendar. See AWAY_MARKER in schema.ts.
  assert.equal(eventSubject(toOooRequest(page(APPROVED_ROW))), "Andon ✈️");
});

test("the calendar shows a first name only, derived from the email", () => {
  assert.equal(calendarNameFor(null, "andon.keller@bluelabellabs.com", "fallback"), "Andon");
  assert.equal(calendarNameFor(null, "jane.doe@bluelabellabs.com", "fallback"), "Jane");
  // No separator in the local part — nothing to split, so use it whole.
  assert.equal(calendarNameFor(null, "mmurphy@bluelabellabs.com", "fallback"), "Mmurphy");
});

test("the BlueLabeler name beats the email — what people go by, not what IT assigned", () => {
  assert.equal(calendarNameFor("Danny", "daniel.deserto@bluelabellabs.com", "fb"), "Danny");
  assert.equal(calendarNameFor("Lu", "ludovic.lacourte@bluelabellabs.com", "fb"), "Lu");
  // And it rescues addresses the email split can't handle.
  assert.equal(calendarNameFor("Mike", "mmurphy@bluelabellabs.com", "fb"), "Mike");
});

test("a shared first name gets the full name, case-insensitively", () => {
  assert.equal(calendarNameFor(null, "chris.boyle@bluelabellabs.com", "fb", ["chris"]), "Chris Boyle");
  assert.equal(calendarNameFor(null, "chris.ferrari@bluelabellabs.com", "fb", ["CHRIS"]), "Chris Ferrari");
  // Everyone else is unaffected by the list.
  assert.equal(calendarNameFor(null, "andon.keller@bluelabellabs.com", "fb", ["chris"]), "Andon");
});

test("a full-name Notion profile still shows only the first name", () => {
  // Some profiles store "Ralph Barile"; the calendar shows "Ralph".
  assert.equal(calendarNameFor("Ralph Barile", "ralph.barile@bluelabellabs.com", "fb"), "Ralph");
  assert.equal(calendarNameFor("Victor Guerreiro", "victor.guerreiro@bluelabellabs.com", "fb"), "Victor");
});

test("disambiguation still applies to a BlueLabeler name", () => {
  assert.equal(calendarNameFor("Chris", "chris.boyle@bluelabellabs.com", "fb", ["chris"]), "Chris Boyle");
  // A Notion name that already carries a surname is left alone.
  assert.equal(calendarNameFor("Chris Ferrari", "chris.ferrari@bluelabellabs.com", "fb", ["chris"]), "Chris Ferrari");
});

test("a listed name with no last name to add stays as-is rather than breaking", () => {
  assert.equal(calendarNameFor(null, "chris@bluelabellabs.com", "fb", ["chris"]), "Chris");
});

test("no name and no email falls back rather than producing a bare plane", () => {
  assert.equal(calendarNameFor(null, null, "Andon time off"), "Andon time off");
});

// --- approval + blocking ---

test("only 'Approved' means an event should exist — Pending included", () => {
  for (const value of ["Pending", "Requested", "Denied", null, "approved"]) {
    assert.equal(isApproved(toOooRequest(page({ ...APPROVED_ROW, Status: status(value) }))), false, `status=${value}`);
  }
  assert.equal(isApproved(toOooRequest(page(APPROVED_ROW))), true);
});

test("a trashed page is never approved, whatever its status says", () => {
  assert.equal(isApproved(toOooRequest(page(APPROVED_ROW, { inTrash: true }))), false);
});

test("blockedReason catches the approve-before-filling-in-dates case", () => {
  assert.equal(blockedReason(toOooRequest(page({ ...APPROVED_ROW, Dates: date(null) }))), "no Dates set");
});

test("blockedReason catches a reversed range", () => {
  const request = toOooRequest(page({ ...APPROVED_ROW, Dates: date("2026-09-11", "2026-09-07") }));
  assert.match(blockedReason(request) ?? "", /ends .* before it starts/);
});
