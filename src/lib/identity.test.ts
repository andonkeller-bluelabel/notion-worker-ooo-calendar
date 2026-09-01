/**
 * Run with: npx tsx --test src/lib/identity.test.ts
 *
 * The load-bearing rule: `Created by` is trusted for Core Requests only.
 * A Flex Request may be anonymous, so trusting it there would file someone's
 * leave under whoever Notion recorded as the creator.
 */

import { test } from "node:test";
import assert from "node:assert/strict";
import { identityFill, needsEmailLookup, type IdentityInputs } from "./identity.js";

const PERSON = { id: "u-andon", name: "Andon", email: "andon.keller@bluelabellabs.com", isPerson: true };
const BOT = { id: "u-bot", name: "Notion MCP", email: null, isPerson: false };

function inputs(over: Partial<IdentityInputs> = {}): IdentityInputs {
  return { source: "Core", hasBlueLabeler: false, email: null, createdBy: PERSON, ...over };
}

// --- Core: the submitter is signed in ---

test("Core fills both the person and their verified email from Created by", () => {
  assert.deepEqual(identityFill(inputs(), null), {
    blueLabelerId: "u-andon",
    email: "andon.keller@bluelabellabs.com",
  });
});

test("Core needs no user lookup — Created by already carries the email", () => {
  assert.equal(needsEmailLookup(inputs()), false);
});

// --- Flex: the submitter may be anonymous ---

test("Flex NEVER takes identity from Created by, even when it looks like a person", () => {
  // The whole point: an anonymous submission's creator is not the requester.
  const fill = identityFill(inputs({ source: "Flex", email: "contractor@example.com" }), null);
  assert.deepEqual(fill, {}, "no attribution without an email match");
});

test("Flex sets the person when the typed email matches a Notion account", () => {
  const fill = identityFill(inputs({ source: "Flex", email: "lu@bluelabellabs.com" }), "u-lu");
  assert.deepEqual(fill, { blueLabelerId: "u-lu" });
});

test("Flex leaves the typed email alone — it is the requester's own", () => {
  const fill = identityFill(inputs({ source: "Flex", email: "someone@example.com" }), "u-x");
  assert.equal(fill.email, undefined);
});

// --- never overwrite a human ---

test("an existing BlueLabeler is never replaced, so filing on someone's behalf survives", () => {
  const fill = identityFill(inputs({ hasBlueLabeler: true }), "u-other");
  assert.equal(fill.blueLabelerId, undefined);
});

test("an existing email is never replaced", () => {
  const fill = identityFill(inputs({ email: "typed@bluelabellabs.com" }), null);
  assert.equal(fill.email, undefined);
  assert.equal(fill.blueLabelerId, "u-andon", "the person is still filled");
});

// --- bots and unknown sources ---

test("a request is never attributed to an integration", () => {
  // The 155 backfilled rows were created by an integration.
  assert.deepEqual(identityFill(inputs({ createdBy: BOT }), null), {});
});

test("a row with no Source behaves like Flex — conservative, not creative", () => {
  assert.deepEqual(identityFill(inputs({ source: null }), null), {});
  assert.deepEqual(identityFill(inputs({ source: null, email: "x@y.com" }), "u-x"), { blueLabelerId: "u-x" });
});

test("nothing to do when there is no creator and no email", () => {
  assert.deepEqual(identityFill(inputs({ createdBy: null }), null), {});
  assert.equal(needsEmailLookup(inputs({ createdBy: null })), false);
});

// --- lookup gating ---

test("the user lookup runs only when it could change something", () => {
  assert.equal(needsEmailLookup(inputs({ source: "Flex", email: "a@b.com" })), true);
  assert.equal(needsEmailLookup(inputs({ source: "Flex", email: "a@b.com", hasBlueLabeler: true })), false);
  assert.equal(needsEmailLookup(inputs({ source: "Flex", email: null })), false);
});
