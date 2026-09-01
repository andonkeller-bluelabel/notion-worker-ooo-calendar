/**
 * Setup and smoke-test tools. Run these with `ntn workers exec` BEFORE wiring
 * the Notion automation, so the things most likely to be wrong get proven
 * first: that the Notion database is one this worker can actually write
 * `O365 Event ID` back to, that the Exchange Application Access Policy really
 * lets us reach the mailbox calendar, and that the token carries the
 * permission we think it does.
 *
 * Two notes for running these by hand with `ntn workers exec`:
 *
 *  - EVERY field is required, `.nullable()` ones included. The SDK's schema
 *    builder puts all object properties in `required` and has no `.optional()`
 *    (see schema-builder.d.ts), so omitting one is a 400
 *    InvalidToolInputError. Pass `null` to skip it:
 *    `ntn workers exec checkOooSetup -d '{"samplePageId": null}'`.
 *  - `exec` doesn't initialize the SDK pacer runtime, so a tool that paces can
 *    fail with `Pacer "..." not found`. `pace()` in graphClient.ts degrades
 *    gracefully for exactly this reason; set GRAPH_DISABLE_PACER=true only if
 *    you still hit it.
 */

import type { JSONValue } from "@notionhq/workers/types";
import { j } from "@notionhq/workers/schema-builder";
import { worker } from "../worker.js";
import { accessToken, graphRequest, safeExecute } from "../lib/graphClient.js";
import { o365CalendarMailbox, oooDataSourceId, sweepLookaheadDays, sweepLookbackDays } from "../lib/env.js";
import { Ooo } from "../lib/schema.js";
import { P, readString, retrievePage, updateProps } from "../lib/notion.js";
import { reconcilePage } from "../lib/liveReconcile.js";
import { listAllEvents, listWorkerEvents, notionPageIdOf } from "../lib/graphCalendar.js";
import { addDays, toDateOnly } from "../lib/oooRequest.js";

/**
 * The property names and Notion types this worker depends on, confirmed
 * against the live "OOO Entries" schema. `required: false` ones only enrich
 * the event body, so their absence is reported but not a failure.
 */
const EXPECTED: ReadonlyArray<{ name: string; type: string; required: boolean }> = [
  { name: Ooo.TITLE, type: "title", required: true },
  { name: Ooo.EMAIL, type: "email", required: true },
  // Preferred identity, but Your Email covers rows it can't resolve.
  { name: Ooo.BLUELABELER, type: "people", required: false },
  { name: Ooo.DATES, type: "date", required: true },
  { name: Ooo.STATUS, type: "status", required: true },
  { name: Ooo.O365_EVENT_ID, type: "rich_text", required: true },
  { name: Ooo.NOTIFIED_STATUS, type: "rich_text", required: true },
  { name: Ooo.CREATED_BY, type: "created_by", required: false },
  // Read for the event body only; absence degrades the body, not the sync.
  { name: Ooo.APPROVER, type: "people", required: false },
  { name: Ooo.NOTES, type: "rich_text", required: false },
];

worker.tool("checkOooSetup", {
  title: "Check the OOO sync setup",
  description:
    "Smoke test to run before wiring the Notion automation. Reports the OOO Entries schema against the " +
    "properties this worker needs, proves the O365 Event ID column is writable (a synced or worker-managed " +
    "database is not), and confirms Microsoft Graph lets this app reach the shared mailbox's calendar.",
  schema: j.object({
    samplePageId: j
      .string()
      .describe("A row in OOO Entries. Its O365 Event ID is read and written back unchanged, to prove the column is writable. Pass null to skip that check.")
      .nullable(),
    lookbackDays: j
      .number()
      .describe("Override how far back workerEvents scans. Defaults to OOO_SWEEP_LOOKBACK_DAYS. Use it to audit history the sweep deliberately does not scan on every run.")
      .nullable(),
    lookaheadDays: j.number().describe("Override how far ahead workerEvents scans.").nullable(),
  }),
  hints: { readOnlyHint: false },
  execute: async ({ samplePageId, lookbackDays, lookaheadDays }, { notion }) =>
    safeExecute(async (): Promise<Record<string, JSONValue>> => {
      const report: Record<string, JSONValue> = {};

      // --- Notion schema ---
      const ds = (await notion.dataSources.retrieve({ data_source_id: oooDataSourceId() })) as {
        name?: string;
        properties?: Record<string, { type?: string }>;
        parent?: unknown;
      };
      const properties = ds.properties ?? {};
      report.notionDatabase = {
        name: ds.name ?? null,
        parent: (ds.parent ?? null) as JSONValue,
        properties: EXPECTED.map((expected) => {
          const actual = properties[expected.name];
          return {
            name: expected.name,
            expectedType: expected.type,
            actualType: actual?.type ?? null,
            required: expected.required,
            ok: actual?.type === expected.type || !expected.required,
          };
        }),
        // Everything the database actually has, so a rename is obvious.
        actualProperties: Object.entries(properties)
          .map(([name, prop]) => `${name} (${prop?.type ?? "?"})`)
          .sort(),
      };

      // --- write-back proof ---
      if (samplePageId) {
        try {
          const page = await retrievePage(notion, samplePageId);
          const current = readString(page, Ooo.O365_EVENT_ID);
          await updateProps(notion, samplePageId, {
            [Ooo.O365_EVENT_ID]: current ? P.richText(current) : P.clearRichText(),
          });
          report.writeBack = { ok: true, pageId: samplePageId, valueUnchanged: current };
        } catch (err) {
          const message = err instanceof Error ? err.message : String(err);
          report.writeBack = {
            ok: false,
            pageId: samplePageId,
            message,
            hint: /read-?only/i.test(message)
              ? "This database's columns are read-only to the API — it is worker-managed or synced from elsewhere. " +
                "The worker cannot own O365 Event ID here; recreate OOO Entries as a native database."
              : null,
          };
        }
      } else {
          report.writeBack = { skipped: "pass samplePageId to prove the O365 Event ID column is writable" };
      }

      // --- Graph reachability ---
      try {
        const calendar = await graphRequest<Record<string, unknown>>(
          `users/${encodeURIComponent(o365CalendarMailbox())}/calendar`,
          { query: { $select: "id,name,owner" } },
        );
        report.graphCalendar = { ok: true, mailbox: o365CalendarMailbox(), calendar: calendar as JSONValue };
      } catch (err) {
        report.graphCalendar = {
          ok: false,
          mailbox: o365CalendarMailbox(),
          message: err instanceof Error ? err.message : String(err),
          hint:
            "ErrorAccessDenied means the Exchange Application Access Policy doesn't cover this mailbox — check " +
            "Test-ApplicationAccessPolicy, and allow up to an hour for a policy change to propagate. " +
            "Authorization_RequestDenied means the Calendars.ReadWrite application permission is missing or unconsented. " +
            "ResourceNotFound/Request_ResourceNotFound means O365_CALENDAR_MAILBOX doesn't resolve to a mailbox — note " +
            "that a Microsoft 365 group is NOT a mailbox for this endpoint and will never work app-only.",
        };
      }

      // --- what the sweep's cleanup pass can actually see ---
      //
      // This is the mechanism behind deletion handling: an event is traceable
      // back to its row only via its page-id tag (or the Notion link in its
      // body). If this comes back empty while events exist on the calendar,
      // orphan cleanup is silently doing nothing.
      try {
        const today = new Date().toISOString().slice(0, 10);
        const back = lookbackDays && lookbackDays > 0 ? Math.floor(lookbackDays) : sweepLookbackDays();
        const ahead = lookaheadDays && lookaheadDays > 0 ? Math.floor(lookaheadDays) : sweepLookaheadDays();
        const events = await listWorkerEvents(`${addDays(today, -back)}T00:00:00`, `${addDays(today, ahead)}T00:00:00`);
        // Two events sharing a page id is the signature of a duplicate.
        const seen = new Map<string, number>();
        for (const { pageId } of events) seen.set(pageId, (seen.get(pageId) ?? 0) + 1);
        const duplicated = [...seen.entries()].filter(([, n]) => n > 1);
        report.workerEvents = {
          ok: true,
          window: `${back}d back / ${ahead}d ahead`,
          traceable: events.length,
          distinctRows: seen.size,
          duplicatedRows: duplicated.length,
          events: events.map(({ pageId, event }) => ({
            subject: event.subject ?? null,
            isAllDay: event.isAllDay ?? null,
            showAs: event.showAs ?? null,
            start: event.start?.dateTime ?? null,
            end: event.end?.dateTime ?? null,
            timeZone: event.start?.timeZone ?? null,
            pageId,
            taggedVia: notionPageIdOf({ ...event, body: undefined }) ? "extendedProperty" : "bodyLink",
            eventId: event.id,
          })),
        };
      } catch (err) {
        report.workerEvents = { ok: false, message: err instanceof Error ? err.message : String(err) };
      }

      return report;
    }),
});

worker.tool("showGraphTokenRoles", {
  title: "Show the Graph token's permissions",
  description:
    "Mints an app-only token and reports the application permissions it actually carries (the `roles` claim), " +
    "plus the tenant and app it was issued to. Use this to tell a missing/wrong permission grant apart from an " +
    "Exchange Application Access Policy denial: both surface as ErrorAccessDenied on the calendar. The token " +
    "itself is never returned or logged.",
  schema: j.object({}),
  hints: { readOnlyHint: true },
  execute: () =>
    safeExecute(async (): Promise<Record<string, JSONValue>> => {
      const token = await accessToken();
      // Read the JWT payload for its claims only. No signature check: we just
      // minted this from Entra over TLS, and nothing here is a trust decision.
      const segment = token.split(".")[1];
      if (!segment) return { ok: false, message: "Token was not a JWT; cannot read its claims." };
      const claims = JSON.parse(Buffer.from(segment, "base64url").toString("utf8")) as Record<string, unknown>;
      const roles = Array.isArray(claims.roles) ? (claims.roles as string[]) : [];

      return {
        ok: true,
        appId: typeof claims.appid === "string" ? claims.appid : null,
        tenantId: typeof claims.tid === "string" ? claims.tid : null,
        audience: typeof claims.aud === "string" ? claims.aud : null,
        roles,
        note:
          roles.length === 0
            ? "No application permissions on this token. The Calendars.ReadWrite grant is missing or was never admin-consented."
            : "Calendars.ReadWrite is what /users/{smtp}/calendar/events needs. If it is present and the calendar " +
              "still returns ErrorAccessDenied, the Application Access Policy is the remaining suspect. Note that no " +
              "role makes /groups/{id}/calendar/events work: Graph lists Application permissions as unsupported there.",
      };
    }),
});

worker.tool("reconcileOooRequest", {
  title: "Reconcile one OOO request",
  description:
    "Runs the reconciler against a single OOO Entries row, exactly as the webhook would. Use it to test " +
    "create / update / delete end to end without touching the Notion automation. Honors OOO_DRYRUN.",
  schema: j.object({ pageId: j.string().describe("The OOO Entries page id.") }),
  hints: { readOnlyHint: false },
  execute: ({ pageId }, { notion }) =>
    safeExecute(async (): Promise<Record<string, JSONValue>> => {
      const { result } = await reconcilePage(notion, pageId);
      return { ok: true, action: result.action, eventId: result.eventId, reason: result.reason ?? null };
    }),
});

worker.tool("scanCalendarBacklog", {
  title: "Scan the calendar for entries to backfill",
  description:
    "READ-ONLY. Lists every event on the shared calendar starting after today, flagging which are already " +
    "worker-managed and which are pre-existing human entries that have no Notion row. Use this to review the " +
    "backfill before importing anything — it writes nothing, to Notion or the calendar.",
  schema: j.object({
    days: j.number().describe("How far ahead to scan. Defaults to OOO_SWEEP_LOOKAHEAD_DAYS.").nullable(),
  }),
  hints: { readOnlyHint: true },
  execute: ({ days }) =>
    safeExecute(async (): Promise<Record<string, JSONValue>> => {
      const today = new Date().toISOString().slice(0, 10);
      // Start at tomorrow: "starting after today" excludes anything already
      // under way, which is not a request anyone still needs to approve.
      const from = addDays(today, 1);
      const horizon = days && days > 0 ? Math.floor(days) : sweepLookaheadDays();
      const events = await listAllEvents(`${from}T00:00:00`, `${addDays(today, horizon)}T00:00:00`);

      const rows = events.map((event) => {
        const startDate = toDateOnly(event.start?.dateTime ?? null);
        const rawEnd = toDateOnly(event.end?.dateTime ?? null);
        return {
          eventId: event.id,
          subject: event.subject ?? null,
          // Graph's all-day end is exclusive, so the last day off is the day
          // before it. Timed events end on the day they say.
          startDate,
          endDateInclusive: rawEnd && event.isAllDay ? addDays(rawEnd, -1) : rawEnd,
          isAllDay: event.isAllDay ?? false,
          showAs: event.showAs ?? null,
          type: event.type ?? null,
          // Already ours: has a Notion row, needs no backfill.
          workerManaged: notionPageIdOf(event) !== null,
          organizer: event.organizer?.emailAddress?.address ?? null,
          attendees: (event.attendees ?? [])
            .map((a) => a.emailAddress?.address ?? null)
            .filter((a): a is string => Boolean(a)),
        };
      });

      const toBackfill = rows.filter((r) => !r.workerManaged);
      return {
        ok: true,
        scanned: { from: `${from}`, throughDays: horizon, total: rows.length },
        alreadyManaged: rows.length - toBackfill.length,
        needsBackfill: toBackfill.length,
        events: toBackfill as unknown as JSONValue,
        note:
          "Nothing was written. Review `events`, then supply an eventId -> email mapping to import them. " +
          "Rows are created with the existing eventId in O365 Event ID so the worker ADOPTS each event " +
          "instead of creating a duplicate alongside it. Watch for type=occurrence rows: those are instances " +
          "of a recurring series and should usually be skipped rather than imported one by one.",
      };
    }),
});
