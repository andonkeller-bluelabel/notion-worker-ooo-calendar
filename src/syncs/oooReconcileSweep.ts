/**
 * oooReconcileSweep — the scheduled safety net behind the webhook.
 *
 * It exists because a Notion "Send webhook" automation can only fire on a page
 * being added or edited. There is no delete trigger, so a trashed request
 * would otherwise leave its event on the shared calendar forever. The sweep is
 * the implementation of that edge case, not an optimization.
 *
 * Two passes:
 *   1. FORWARD  — every row that is Approved or still carries an event id is
 *      re-reconciled. Catches trashed-then-restored rows, approvals made
 *      before the dates were filled in (the automation fires once, on the
 *      status change, and never again), deliveries the platform dropped, and
 *      events a person deleted straight off the calendar.
 *   2. ORPHANS  — every worker-owned event in the window whose row no longer
 *      justifies it is removed: the page was trashed, hard-deleted, or is no
 *      longer Approved (the `page.deleted` case), OR the page is Approved but
 *      names a different event, which makes this one a stale duplicate.
 *
 * The orphan pass is skipped entirely if any row failed in pass 1: a row we
 * couldn't read is a row whose event we must not judge. Default-deny, same
 * posture as the Read.ai worker's hold gate.
 *
 * Attached to a managed anchor database because the platform only schedules
 * `worker.sync` and a sync must attach to one. The anchor stays empty; all
 * real work goes through the raw `context.notion` client and Graph.
 */

import * as Schema from "@notionhq/workers/schema";
import { worker } from "../worker.js";
import { Ooo, CALENDAR_STATUSES } from "../lib/schema.js";
import { pageUrl, queryAll, retrievePage, type Notion } from "../lib/notion.js";
import { toOooRequest, addDays, isApproved } from "../lib/oooRequest.js";
import { fillIdentity, reconcileRequest } from "../lib/liveReconcile.js";
import { deleteEvent, listWorkerEvents } from "../lib/graphCalendar.js";
import { isNotFound } from "../lib/errors.js";
import { disambiguateFirstNames, isDryRun, oooDataSourceId, sweepLookaheadDays, sweepLookbackDays } from "../lib/env.js";
import { notifyOps, slackLink } from "../lib/slack.js";

const TAG = "oooReconcileSweep";

/**
 * True when Notion says the page behind an event no longer exists at all
 * (hard-deleted after 30 days in the trash). Notion's client raises
 * `object_not_found`; the Graph helper covers the case where the failure came
 * from our side of the call instead.
 */
function isPageGone(err: unknown): boolean {
  if (isNotFound(err)) return true;
  const code = (err as { code?: unknown })?.code;
  if (code === "object_not_found") return true;
  return err instanceof Error && /could not find page|object_not_found/i.test(err.message);
}

/**
 * Scheduler anchor. A managed database's non-title columns are read-only to
 * the API, so it can never hold real data — it exists only to host this sync's
 * schedule. The actual rows live in the NATIVE "OOO Entries" database
 * addressed by OOO_DATA_SOURCE_ID.
 */
export const sweepAnchorDb = worker.database("oooSweepAnchor", {
  type: "managed",
  initialTitle: "OOO Sweep Anchor (internal — do not edit)",
  primaryKeyProperty: "Key",
  schema: {
    properties: {
      Name: Schema.title(),
      Key: Schema.richText(),
    },
  },
});

/**
 * Rows worth reconciling: anything in a calendar status (should have an event)
 * or still
 * carrying an event id (may need one removed) — AND within the calendar window
 * this sweep actually polices. A Requested row with no event id reconciles to a
 * no-op, so there is no reason to fetch it.
 *
 * The date bound matters at scale. Without it the sweep re-PATCHed every
 * approved event on every run; once a year of history was imported that was 171
 * Graph writes per run, took ~4.5 minutes, and hit the platform's execution
 * timeout mid-pass. Past time off is immutable, so policing it every ten
 * minutes buys nothing. Rows with no dates are always included: those are the
 * approved-but-unfilled ones a human still needs to fix.
 *
 * An edit to a row older than the window is still picked up — the webhook fires
 * on `Dates` and `Status` changes regardless of how old the row is. The sweep is
 * the safety net for current and upcoming time off, not an audit of the archive.
 */
function sweepFilter(windowStartDate: string) {
  return {
    and: [
      {
        or: [
          // `Status` is a status-type property, so the key is `status`, not `select`.
          ...CALENDAR_STATUSES.map((name) => ({ property: Ooo.STATUS, status: { equals: name } })),
          { property: Ooo.O365_EVENT_ID, rich_text: { is_not_empty: true } },
        ],
      },
      {
        or: [
          { property: Ooo.DATES, date: { on_or_after: windowStartDate } },
          { property: Ooo.DATES, date: { is_empty: true } },
        ],
      },
    ],
  };
}

interface SweepTally {
  reconciled: number;
  failed: number;
  orphansRemoved: number;
  duplicatesRemoved: number;
}

async function runSweep(notion: Notion): Promise<SweepTally> {
  const tally: SweepTally = { reconciled: 0, failed: 0, orphansRemoved: 0, duplicatesRemoved: 0 };

  // Window shared by both passes, so pass 1 never reconciles a row whose event
  // pass 2 cannot see.
  const today = new Date().toISOString().slice(0, 10);
  const windowStartDate = addDays(today, -sweepLookbackDays());

  // --- pass 1: forward reconcile ---
  const rows = await queryAll(notion, oooDataSourceId(), sweepFilter(windowStartDate));
  console.log(`[${TAG}] ${rows.length} row(s) to reconcile`);

  /** Event ids that a live, approved row legitimately owns after this pass. */
  const liveEventIds = new Set<string>();

  for (const row of rows) {
    // Same identity fill the webhook does, for rows whose delivery never landed.
    const request = toOooRequest(await fillIdentity(notion, row), disambiguateFirstNames());
    try {
      // quiet: the webhook already alerted on anything blocked; re-alerting
      // every hour would turn a stuck row into a stuck channel.
      const result = await reconcileRequest(notion, request, { quiet: true });
      tally.reconciled += 1;
      if (result.action !== "noop") {
        console.log(`[${TAG}] ${request.personName} (${row.id}) → ${result.action}${result.reason ? ` (${result.reason})` : ""}`);
      }
      // A blocked row keeps whatever event it had; its id is still legitimate.
      if (result.eventId) liveEventIds.add(result.eventId);
    } catch (err) {
      tally.failed += 1;
      console.error(`[${TAG}] row ${row.id} (${request.personName}) failed:`, err);
      // Keep the stored id live so a transient failure can't get its event
      // deleted as an orphan on this same run.
      if (request.eventId) liveEventIds.add(request.eventId);
    }
  }

  if (tally.failed > 0) {
    await notifyOps(
      `:warning: OOO sync: ${tally.failed} of ${rows.length} request(s) failed to reconcile. ` +
        `Skipping the calendar cleanup pass this run so nothing is removed by mistake.`,
    );
    return tally;
  }

  // --- pass 2: orphaned calendar events ---
  const windowStart = `${windowStartDate}T00:00:00`;
  const windowEnd = `${addDays(today, sweepLookaheadDays())}T00:00:00`;

  const events = await listWorkerEvents(windowStart, windowEnd);
  console.log(`[${TAG}] ${events.length} worker-owned event(s) on the calendar in window`);
  for (const { pageId: normalizedPageId, event } of events) {
    if (!event.id || liveEventIds.has(event.id)) continue;

    // Confirm against Notion before removing anything. An event whose row we
    // simply failed to see is not an orphan.
    let keep = false;
    let staleDuplicate = false;
    try {
      const page = await retrievePage(notion, normalizedPageId);
      const request = toOooRequest(page, disambiguateFirstNames());
      if (isApproved(request)) {
        if (request.eventId && request.eventId !== event.id) {
          // The row is approved and names a DIFFERENT event, so Notion — the
          // source of truth for which event is current — has already moved on
          // from this one. It is a leftover, and removing it is unambiguous.
          //
          // Two ways to get here. Concurrent deliveries: two automations fire
          // on one change, both read an empty O365 Event ID, both create an
          // event, and only the last id survives (observed 2026-09-01). Or
          // someone clears the id by hand while the row is still Approved, and
          // the next reconcile makes a second event.
          staleDuplicate = true;
        } else {
          // Approved and either pointing at this event or (transiently) at
          // nothing. Pass 1 already reconciled this row, so leave it alone.
          keep = true;
        }
      }
    } catch (err) {
      // 404 means the page was hard-deleted — a genuine orphan. Anything else
      // is an unknown, and unknowns are left alone.
      if (!isPageGone(err)) {
        console.error(`[${TAG}] couldn't check page ${normalizedPageId} behind event ${event.id}:`, err);
        continue;
      }
    }

    if (keep) continue;

    if (isDryRun()) {
      console.log(
        `[${TAG}] DRY-RUN would delete ${staleDuplicate ? "duplicate" : "orphaned"} event ${event.id} ("${event.subject ?? ""}")`,
      );
      continue;
    }
    try {
      await deleteEvent(event.id);
      if (staleDuplicate) {
        tally.duplicatesRemoved += 1;
        console.log(`[${TAG}] deleted duplicate event ${event.id} ("${event.subject ?? ""}") — its row names a different event`);
        await notifyOps(
          `:broom: OOO sync: removed a duplicate calendar event for ` +
            `${slackLink(event.subject ?? "a request", pageUrl(normalizedPageId))}. The one Notion tracks is untouched.`,
        );
      } else {
        tally.orphansRemoved += 1;
        console.log(`[${TAG}] deleted orphaned event ${event.id} ("${event.subject ?? ""}") — its request is gone`);
      }
    } catch (err) {
      if (isNotFound(err)) continue; // raced with something else; fine
      console.error(`[${TAG}] couldn't delete orphaned event ${event.id}:`, err);
    }
  }

  return tally;
}

worker.sync("oooReconcileSweep", {
  database: sweepAnchorDb,
  mode: "incremental",
  // 10m, not hourly. Notion's "Any property edited" automation trigger does
  // NOT fire on status-property changes (observed 2026-09-01: three Status
  // edits produced zero deliveries, while a Notes edit produced two). Status is
  // the property that decides whether an event should exist, so the webhook
  // cannot be relied on for the transition that matters and this sweep is the
  // primary mechanism, not a backstop. It reads only rows that are Approved or
  // still carry an event id, so a run is a couple of API calls.
  schedule: "10m",
  execute: async (_state, { notion }) => {
    const tally = await runSweep(notion);
    console.log(
      `[${TAG}] done: ${tally.reconciled} reconciled, ${tally.failed} failed, ` +
        `${tally.orphansRemoved} orphan(s) removed, ${tally.duplicatesRemoved} duplicate(s) removed`,
    );
    // The anchor exists only for scheduling — never write rows to it.
    return { changes: [], hasMore: false };
  },
});

export { runSweep };
