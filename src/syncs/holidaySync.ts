/**
 * holidaySync — mirrors the two holiday databases onto the shared calendar.
 *
 * The calendar is fed by THREE Notion databases, each in its own lane:
 * OOO Entries (a request workflow, handled by the webhook and sweep), plus
 * BlueLabel US Holidays and Vendor Partner Holidays (reference data, handled
 * here). Holidays never enter OOO Entries — a public holiday has no requester,
 * no approver, and no status anyone set.
 *
 * Each lane only ever touches events carrying its own tag, so no sync can
 * delete another's work and events a person created by hand are safe from all
 * of them.
 *
 * Holidays change a few times a year, so this runs every 6 hours rather than
 * the 10 minutes time off needs.
 */

import { worker } from "../worker.js";
import { sweepAnchorDb } from "./oooReconcileSweep.js";
import { P, pageUrl, queryAll, retrievePage, updateProps, type Notion } from "../lib/notion.js";
import { toHolidayEntry, usSource, vendorSource, type HolidaySource } from "../lib/holidays.js";
import { addDays } from "../lib/oooRequest.js";
import { createHolidayEvent, deleteEvent, listHolidayEvents, updateHolidayEvent } from "../lib/graphCalendar.js";
import { isNotFound } from "../lib/errors.js";
import { isDryRun, sweepLookaheadDays, sweepLookbackDays, usHolidaysDataSourceId, vendorHolidaysDataSourceId } from "../lib/env.js";
import { notifyOps } from "../lib/slack.js";

const TAG = "holidaySync";

interface Tally {
  created: number;
  updated: number;
  removed: number;
  failed: number;
}

/** True when Notion says the page behind an event no longer exists at all. */
function isPageGone(err: unknown): boolean {
  if (isNotFound(err)) return true;
  const code = (err as { code?: unknown })?.code;
  if (code === "object_not_found") return true;
  return err instanceof Error && /could not find page|object_not_found/i.test(err.message);
}

/** Reconciles one source's rows. Returns the event ids it legitimately owns. */
async function syncSource(notion: Notion, source: HolidaySource, windowStart: string, tally: Tally): Promise<Set<string>> {
  const dataSourceId = source.dataSourceId();
  const live = new Set<string>();
  if (!dataSourceId) {
    console.log(`[${TAG}] ${source.kind}: no data source configured — skipping`);
    return live;
  }

  // Only rows in the window the cleanup pass can also see, for the same reason
  // the OOO sweep is bounded: past holidays are immutable and re-checking a
  // decade of them on every run buys nothing.
  const rows = await queryAll(notion, dataSourceId, {
    property: source.dateProp,
    date: { on_or_after: windowStart },
  });
  console.log(`[${TAG}] ${source.kind}: ${rows.length} row(s) in window`);

  for (const row of rows) {
    const entry = toHolidayEntry(row, source);
    try {
      if (!entry) {
        // Undated or unnamed now, but it had an event: take it off the calendar.
        const stale = row.properties[source.eventIdProp];
        const staleId = typeof stale === "object" && stale ? String((stale as any).rich_text?.[0]?.plain_text ?? "") : "";
        if (staleId && !isDryRun()) {
          await deleteEvent(staleId).catch((err) => {
            if (!isNotFound(err)) throw err;
          });
          await updateProps(notion, row.id, { [source.eventIdProp]: P.clearRichText() });
          tally.removed += 1;
        }
        continue;
      }

      const payload = { ...entry, pageUrl: pageUrl(row.id, row.url) };
      if (isDryRun()) {
        console.log(`[${TAG}] DRY-RUN ${entry.eventId ? "update" : "create"} ${entry.subject}`);
        if (entry.eventId) live.add(entry.eventId);
        continue;
      }

      if (!entry.eventId) {
        const created = await createHolidayEvent(payload);
        await updateProps(notion, row.id, { [source.eventIdProp]: P.richText(created.id) });
        live.add(created.id);
        tally.created += 1;
        continue;
      }

      try {
        await updateHolidayEvent(entry.eventId, payload);
        live.add(entry.eventId);
        tally.updated += 1;
      } catch (err) {
        if (!isNotFound(err)) throw err;
        // Someone deleted the event off the calendar; the row still says it
        // should exist, so put it back.
        const created = await createHolidayEvent(payload);
        await updateProps(notion, row.id, { [source.eventIdProp]: P.richText(created.id) });
        live.add(created.id);
        tally.created += 1;
      }
    } catch (err) {
      tally.failed += 1;
      console.error(`[${TAG}] ${source.kind} row ${row.id} failed:`, err);
      // Keep a failed row's id live so a transient error can't get its event
      // deleted as an orphan on this same run.
      if (entry?.eventId) live.add(entry.eventId);
    }
  }
  return live;
}

async function run(notion: Notion): Promise<Tally> {
  const tally: Tally = { created: 0, updated: 0, removed: 0, failed: 0 };
  const today = new Date().toISOString().slice(0, 10);
  const windowStart = addDays(today, -sweepLookbackDays());

  const sources = [usSource(usHolidaysDataSourceId), vendorSource(vendorHolidaysDataSourceId)];
  const live = new Set<string>();
  for (const source of sources) {
    // One source failing must not stop the other.
    try {
      for (const id of await syncSource(notion, source, windowStart, tally)) live.add(id);
    } catch (err) {
      tally.failed += 1;
      console.error(`[${TAG}] ${source.kind} source failed:`, err);
    }
  }

  if (tally.failed > 0) {
    await notifyOps(
      `:warning: OOO sync: ${tally.failed} holiday row(s)/source(s) failed to mirror. ` +
        `Skipping the calendar cleanup pass this run so nothing is removed by mistake.`,
    );
    return tally;
  }

  // Cleanup: mirrored holiday events whose row is gone. Only ever looks at
  // events tagged into the HOLIDAY lane, so time off and human-created events
  // are untouched.
  const events = await listHolidayEvents(`${windowStart}T00:00:00`, `${addDays(today, sweepLookaheadDays())}T00:00:00`);
  console.log(`[${TAG}] ${events.length} mirrored holiday event(s) on the calendar in window`);
  for (const { pageId, event } of events) {
    if (!event.id || live.has(event.id)) continue;
    try {
      await retrievePage(notion, pageId);
      // The page exists but this event is not the one it names — a leftover
      // from an interrupted run.
    } catch (err) {
      if (!isPageGone(err)) {
        console.error(`[${TAG}] couldn't check page ${pageId} behind event ${event.id}:`, err);
        continue;
      }
    }
    if (isDryRun()) {
      console.log(`[${TAG}] DRY-RUN would delete stale holiday event ${event.id} ("${event.subject ?? ""}")`);
      continue;
    }
    try {
      await deleteEvent(event.id);
      tally.removed += 1;
      console.log(`[${TAG}] deleted stale holiday event ${event.id} ("${event.subject ?? ""}")`);
    } catch (err) {
      if (!isNotFound(err)) console.error(`[${TAG}] couldn't delete ${event.id}:`, err);
    }
  }
  return tally;
}

worker.sync("holidaySync", {
  database: sweepAnchorDb,
  mode: "incremental",
  schedule: "6h",
  execute: async (_state, { notion }) => {
    const t = await run(notion);
    console.log(`[${TAG}] done: ${t.created} created, ${t.updated} updated, ${t.removed} removed, ${t.failed} failed`);
    return { changes: [], hasMore: false };
  },
});

export { run as runHolidaySync };
