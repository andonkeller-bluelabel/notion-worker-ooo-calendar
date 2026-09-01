/**
 * Run with: npx tsx --test src/lib/identity.test.ts
 *
 * The load-bearing rule: whenever an email is present, `Created by` is ignored.
 * A Flex Request may be anonymous, and Notion does not guarantee what it
 * records as the creator of one — trusting it would file a contractor's leave
 * under whoever that turns out to be.
 */

import { test } from "node:test";
import assert from "node:assert/strict";
import { filledName, identityFill, needsEmailLookup, type IdentityInputs } from "./identity.js";

const PERSON = { id: "u-andon", name: "Andon", email: "andon.keller@bluelabellabs.com", isPerson: true };
const BOT = { id: "u-bot", name: "Notion MCP", email: null, isPerson: false };

function inputs(over: Partial<IdentityInputs> = {}): IdentityInputs {
  return { hasBlueLabeler: false, email: null, createdBy: PERSON, ...over };
}

// --- Core-shaped: no email typed, so the signed-in creator is the requester ---

test("no email means Created by supplies both the person and their verified address", () => {
  assert.deepEqual(identityFill(inputs(), null), {
    blueLabelerId: "u-andon",
    email: "andon.keller@bluelabellabs.com",
  });
});

test("no email means no user lookup — Created by already carries the address", () => {
  assert.equal(needsEmailLookup(inputs()), false);
});

// --- Flex-shaped: an email was typed, so Created by is irrelevant ---

test("an email present makes Created by IRRELEVANT, however person-like it looks", () => {
  // If Notion attributes an anonymous submission to the form's owner, this is
  // the test that stops their name landing on a contractor's calendar entry.
  const fill = identityFill(inputs({ email: "contractor@example.com", createdBy: PERSON }), null);
  assert.deepEqual(fill, {});
});

test("an email that matches a Notion account still links the person", () => {
  assert.deepEqual(identityFill(inputs({ email: "lu@bluelabellabs.com" }), "u-lu"), { blueLabelerId: "u-lu" });
});

test("a typed email is never rewritten — it is the requester's own", () => {
  assert.equal(identityFill(inputs({ email: "someone@example.com" }), "u-x").email, undefined);
});

// --- never overwrite a human ---

test("an existing BlueLabeler is never replaced, so filing on someone's behalf survives", () => {
  assert.equal(identityFill(inputs({ hasBlueLabeler: true }), "u-other").blueLabelerId, undefined);
  assert.equal(identityFill(inputs({ hasBlueLabeler: true, email: "a@b.com" }), "u-other").blueLabelerId, undefined);
});

// --- bots ---

test("a request is never attributed to an integration", () => {
  // The 155 backfilled rows were created by one.
  assert.deepEqual(identityFill(inputs({ createdBy: BOT }), null), {});
});

test("nothing to do when there is no creator and no email", () => {
  assert.deepEqual(identityFill(inputs({ createdBy: null }), null), {});
  assert.equal(needsEmailLookup(inputs({ createdBy: null })), false);
});

// --- lookup gating ---

test("the user lookup runs only when it could change something", () => {
  assert.equal(needsEmailLookup(inputs({ email: "a@b.com" })), true);
  assert.equal(needsEmailLookup(inputs({ email: "a@b.com", hasBlueLabeler: true })), false);
  assert.equal(needsEmailLookup(inputs({ email: null })), false);
});

// --- whose name ends up on the calendar ---

const ANON = { id: "notion_user-00000000-0000-0000-0000-00000000000a", name: "Anonymous", email: null, isPerson: false };

test("the name comes from the matched user, NOT the anonymous creator", () => {
  // The regression: an anonymous submission's creator is a sentinel user named
  // "Anonymous", and reading it here put "✈️ Anonymous" on the calendar for a
  // row whose BlueLabeler had resolved correctly to a real person.
  assert.equal(filledName("u-andon", { id: "u-andon", name: "Andon" }, ANON), "Andon");
});

test("the name comes from the creator on the Created-by path", () => {
  assert.equal(filledName("u-andon", null, PERSON), "Andon");
});

test("a name is never borrowed from someone we did not fill from", () => {
  assert.equal(filledName("u-andon", { id: "u-someone-else", name: "Wrong" }, null), null);
  assert.equal(filledName("u-andon", null, { ...PERSON, id: "u-different" }), null);
  assert.equal(filledName(undefined, { id: "u-andon", name: "Andon" }, PERSON), null);
});

test("a bot never supplies a name", () => {
  assert.equal(filledName("u-bot", null, BOT), null);
});
