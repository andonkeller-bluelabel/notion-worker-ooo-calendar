/**
 * Run with: npx tsx --test src/lib/reconcile.test.ts
 *
 * The reconciler's whole contract: the calendar ends up matching whatever the
 * row currently says, no matter which property changed or how many times the
 * same delivery arrives. These lock in the create / update / delete paths and
 * the self-healing behaviour around events that have gone missing.
 */

import { test } from "node:test";
import assert from "node:assert/strict";
import { reconcile, type ReconcileDeps } from "./reconcile.js";
import type { OooRequest } from "./oooRequest.js";

class NotFound extends Error {}

interface Harness {
  deps: ReconcileDeps;
  created: OooRequest[];
  updated: Array<{ eventId: string; request: OooRequest }>;
  deleted: string[];
  stored: Array<{ pageId: string; eventId: string }>;
  cleared: string[];
  titles: Array<{ pageId: string; title: string }>;
  announced: string[];
  notified: Array<{ pageId: string; status: string }>;
  notices: string[];
}

function harness(opts: { missingEventIds?: string[]; dryRun?: boolean; newEventId?: string } = {}): Harness {
  const missing = new Set(opts.missingEventIds ?? []);
  const h: Partial<Harness> = { created: [], updated: [], deleted: [], stored: [], cleared: [], titles: [], announced: [], notified: [], notices: [] };
  h.deps = {
    createEvent: async (request) => {
      h.created!.push(request);
      return { id: opts.newEventId ?? "EVT-NEW" };
    },
    updateEvent: async (eventId, request) => {
      if (missing.has(eventId)) throw new NotFound(eventId);
      h.updated!.push({ eventId, request });
    },
    deleteEvent: async (eventId) => {
      if (missing.has(eventId)) throw new NotFound(eventId);
      h.deleted!.push(eventId);
    },
    storeEventId: async (pageId, eventId) => {
      h.stored!.push({ pageId, eventId });
    },
    clearEventId: async (pageId) => {
      h.cleared!.push(pageId);
    },
    setTitle: async (pageId, title) => {
      h.titles!.push({ pageId, title });
    },
    announce: async (text) => {
      h.announced!.push(text);
    },
    setNotifiedStatus: async (pageId, status) => {
      h.notified!.push({ pageId, status });
    },
    isNotFound: (err) => err instanceof NotFound,
    notify: async (text) => {
      h.notices!.push(text);
    },
    dryRun: opts.dryRun,
  };
  return h as Harness;
}

function request(over: Partial<OooRequest> = {}): OooRequest {
  return {
    pageId: "page-1",
    pageUrl: "https://notion.so/page1",
    personName: "Andon Keller",
    calendarName: "Andon",
    currentTitle: "✈️ Andon",
    notifiedStatus: "Approved",
    hasIdentity: true,
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

// --- create ---

test("first approval creates the event and writes the id back", async () => {
  const h = harness();
  const result = await reconcile(request(), h.deps);

  assert.equal(result.action, "created");
  assert.equal(result.eventId, "EVT-NEW");
  assert.equal(h.created.length, 1);
  assert.deepEqual(h.stored, [{ pageId: "page-1", eventId: "EVT-NEW" }]);
  assert.equal(h.deleted.length, 0);
});

// --- update ---

test("a change after approval updates the existing event and never creates a second one", async () => {
  const h = harness();
  const result = await reconcile(request({ eventId: "EVT-1", endDate: "2026-09-18" }), h.deps);

  assert.equal(result.action, "updated");
  assert.equal(result.eventId, "EVT-1");
  assert.equal(h.created.length, 0, "must not duplicate the event");
  assert.equal(h.updated[0]?.eventId, "EVT-1");
  assert.equal(h.updated[0]?.request.endDate, "2026-09-18");
  assert.equal(h.stored.length, 0, "the id didn't change, so nothing to write back");
});

test("running the same approved row repeatedly is idempotent", async () => {
  const h = harness();
  const row = request({ eventId: "EVT-1" });
  for (let i = 0; i < 3; i++) await reconcile(row, h.deps);

  assert.equal(h.created.length, 0);
  assert.equal(h.deleted.length, 0);
  assert.equal(h.updated.length, 3, "each run re-asserts the same state");
});

test("an event deleted straight off the calendar is recreated and the new id stored", async () => {
  const h = harness({ missingEventIds: ["EVT-GONE"], newEventId: "EVT-2" });
  const result = await reconcile(request({ eventId: "EVT-GONE" }), h.deps);

  assert.equal(result.action, "recreated");
  assert.equal(result.eventId, "EVT-2");
  assert.deepEqual(h.stored, [{ pageId: "page-1", eventId: "EVT-2" }]);
});

// --- delete ---

test("un-approving deletes the event and clears the id", async () => {
  for (const status of ["Requested", "Denied", null]) {
    const h = harness();
    const result = await reconcile(request({ status, eventId: "EVT-1" }), h.deps);

    assert.equal(result.action, "deleted", `status=${status}`);
    assert.equal(result.eventId, null);
    assert.deepEqual(h.deleted, ["EVT-1"]);
    assert.deepEqual(h.cleared, ["page-1"]);
  }
});

test("a trashed page's event is deleted — the page.deleted case", async () => {
  const h = harness();
  const result = await reconcile(request({ inTrash: true, eventId: "EVT-1" }), h.deps);

  assert.equal(result.action, "deleted");
  assert.deepEqual(h.deleted, ["EVT-1"]);
  assert.deepEqual(h.cleared, [], "Notion rejects property writes to a trashed page");
});

test("a delete that 404s still clears the id, so a later re-approval can't patch a ghost", async () => {
  const h = harness({ missingEventIds: ["EVT-1"] });
  const result = await reconcile(request({ status: "Denied", eventId: "EVT-1" }), h.deps);

  assert.equal(result.action, "already-absent");
  assert.equal(result.eventId, null);
  assert.deepEqual(h.cleared, ["page-1"]);
});

// --- no-op and blocked ---

test("an unapproved row with no event id does nothing at all", async () => {
  const h = harness();
  const result = await reconcile(request({ status: "Requested" }), h.deps);

  assert.equal(result.action, "noop");
  assert.equal(h.created.length + h.updated.length + h.deleted.length, 0);
});

test("approved but dateless creates nothing, and says why", async () => {
  const h = harness();
  const result = await reconcile(request({ startDate: null, endDate: null }), h.deps);

  assert.equal(result.action, "blocked");
  assert.equal(result.reason, "no Dates set");
  assert.equal(h.created.length, 0, "a bogus all-day event is worse than none");
  assert.equal(h.notices.length, 1, "a human has to be told, since no second delivery is coming");
});

// --- dry run ---

test("dry run writes nothing on any path", async () => {
  for (const row of [request(), request({ eventId: "EVT-1" }), request({ status: "Denied", eventId: "EVT-1" })]) {
    const h = harness({ dryRun: true });
    const result = await reconcile(row, h.deps);
    assert.equal(result.action, "dry-run");
    assert.equal(h.created.length + h.updated.length + h.deleted.length + h.stored.length + h.cleared.length, 0);
  }
});

// --- known edge case, asserted so the behaviour is deliberate ---

test("KNOWN GAP: clearing O365 Event ID by hand while Approved creates a duplicate", async () => {
  const h = harness();
  const result = await reconcile(request({ eventId: null }), h.deps);
  // The original event is still on the calendar; we have no way to recognize
  // it from the row alone. Accepted: the sweep flags it for a human.
  assert.equal(result.action, "created");
});

// --- Notion title mirrors the calendar subject ---

test("a stale title is rewritten to match the calendar subject", async () => {
  const h = harness();
  await reconcile(request({ currentTitle: "New submission" }), h.deps);
  assert.deepEqual(h.titles, [{ pageId: "page-1", title: "✈️ Andon" }]);
});

test("a title that already matches is left alone — an unchanged sweep writes nothing", async () => {
  const h = harness();
  await reconcile(request({ eventId: "EVT-1" }), h.deps);
  assert.equal(h.titles.length, 0);
});

test("the title is kept in step whatever the status, not just when approved", async () => {
  for (const status of ["Requested", "Denied", "Pending"]) {
    const h = harness();
    await reconcile(request({ status, currentTitle: "whatever" }), h.deps);
    assert.deepEqual(h.titles, [{ pageId: "page-1", title: "✈️ Andon" }], `status=${status}`);
  }
});

test("no identity means no title write — it would feed on itself", async () => {
  // With no BlueLabeler and no email the name comes FROM the title, so writing
  // a computed title back would append a marker on every pass.
  const h = harness();
  await reconcile(request({ hasIdentity: false, currentTitle: "Some row" }), h.deps);
  assert.equal(h.titles.length, 0);
});

test("a trashed page is never written to", async () => {
  const h = harness();
  await reconcile(request({ inTrash: true, eventId: "EVT-1", currentTitle: "stale" }), h.deps);
  assert.equal(h.titles.length, 0);
});

test("dry run writes no title either", async () => {
  const h = harness({ dryRun: true });
  await reconcile(request({ currentTitle: "stale" }), h.deps);
  assert.equal(h.titles.length, 0);
});

// --- Slack announcements ---

test("says nothing when the status has not changed since the last announcement", async () => {
  const h = harness();
  await reconcile(request({ status: "Approved", notifiedStatus: "Approved" }), h.deps);
  assert.deepEqual(h.announced, []);
  assert.deepEqual(h.notified, []);
});

test("announces a new submission the first time it is seen", async () => {
  const h = harness();
  await reconcile(request({ status: "Requested", notifiedStatus: null }), h.deps);
  assert.equal(h.announced.length, 1);
  assert.match(h.announced[0]!, /requested time off/);
  assert.deepEqual(h.notified, [{ pageId: "page-1", status: "Requested" }]);
});

test("announces an approval, and records it so the sweep stays quiet after", async () => {
  const h = harness();
  await reconcile(request({ status: "Approved", notifiedStatus: "Requested" }), h.deps);
  assert.match(h.announced[0]!, /approved and on the team calendar/);
  assert.deepEqual(h.notified, [{ pageId: "page-1", status: "Approved" }]);
});

test("announces Requested -> Denied, which moves no event and would otherwise be invisible", async () => {
  const h = harness();
  await reconcile(request({ status: "Denied", notifiedStatus: "Requested", eventId: null }), h.deps);
  assert.match(h.announced[0]!, /was denied/);
  assert.doesNotMatch(h.announced[0]!, /Removed from the team calendar/);
});

test("a denial that DID remove an event says so", async () => {
  const h = harness();
  await reconcile(request({ status: "Denied", notifiedStatus: "Approved", eventId: "EVT-1" }), h.deps);
  assert.match(h.announced[0]!, /Removed from the team calendar/);
});

test("a trashed page is never announced — its removal is not news", async () => {
  const h = harness();
  await reconcile(request({ inTrash: true, status: "Approved", notifiedStatus: null, eventId: "EVT-1" }), h.deps);
  assert.deepEqual(h.announced, []);
});

test("dry run announces nothing", async () => {
  const h = harness({ dryRun: true });
  await reconcile(request({ status: "Approved", notifiedStatus: null }), h.deps);
  assert.deepEqual(h.announced, []);
});

test("a Slack failure never breaks the calendar work", async () => {
  const h = harness();
  h.deps.announce = async () => {
    throw new Error("slack is down");
  };
  const result = await reconcile(request({ status: "Approved", notifiedStatus: null, eventId: null }), h.deps);
  assert.equal(result.action, "created", "the event still gets created");
});
