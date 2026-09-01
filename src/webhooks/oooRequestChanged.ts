/**
 * Entry point — a Notion automation on the OOO Entries database calls
 * this endpoint whenever a row is added or edited, and the reconciler makes
 * the shared calendar match the row.
 *
 * Implemented as `worker.webhook` driven by a Notion "Send webhook" action,
 * the same mechanism the Kantata worker uses for `createSprintTasks` /
 * `updateDates` and the Read.ai worker uses for `fileApprovedMeeting`. The
 * platform `worker.automation` capability is not enabled for this workspace.
 * Notion-originated deliveries are authenticated by the platform, so there is
 * no HMAC to verify here.
 *
 * The delivery is a signal, not a snapshot: its `properties` object is empty,
 * so we take the page id and re-read the row before acting.
 *
 * Deletions do NOT arrive here — a Notion automation has no "page deleted"
 * trigger. Trashed rows are cleaned up by `oooReconcileSweep`.
 */

import { worker } from "../worker.js";
import { reconcilePage } from "../lib/liveReconcile.js";
import { isDebug } from "../lib/env.js";
import { notifyOps, slackLink } from "../lib/slack.js";
import { pageUrl } from "../lib/notion.js";

const TAG = "oooRequestChanged";

/**
 * Extracts the row's page id from the Notion automation webhook body. The
 * payload nests the page reference under `data` on some trigger types and
 * puts it at the top level on others, so accept both.
 */
function requestPageId(body: Record<string, unknown>): string | null {
  const data = ((body.data as Record<string, unknown> | undefined) ?? body) as Record<string, unknown>;
  return typeof data.id === "string" && data.id ? data.id : null;
}

worker.webhook("oooRequestChanged", {
  title: "OOO Entries changed",
  description:
    "Triggered by a Notion automation on OOO Entries (page added or edited, Send webhook). Re-reads the row " +
    "and reconciles the Microsoft 365 group calendar to match it: creates, updates, or removes the time-off event.",
  execute: async (events, context) => {
    for (const event of events) {
      if (isDebug()) {
        console.log(`[${TAG}] DEBUG delivery ${event.deliveryId} rawBody=${event.rawBody}`);
      }

      const pageId = requestPageId(event.body ?? {});
      if (!pageId) {
        console.warn(`[${TAG}] delivery ${event.deliveryId} had no page id — ignoring. Raw body: ${event.rawBody}`);
        continue;
      }

      try {
        const { page, result } = await reconcilePage(context.notion, pageId);
        console.log(`[${TAG}] page ${pageId} → ${result.action}${result.reason ? ` (${result.reason})` : ""}`);
        if (result.action === "recreated") {
          await notifyOps(
            `:calendar: OOO sync: the calendar event for ${slackLink("this request", pageUrl(page.id, page.url))} ` +
              `was missing and has been recreated.`,
          );
        }
      } catch (err) {
        console.error(`[${TAG}] page ${pageId} failed:`, err);
        await notifyOps(
          `:warning: OOO sync: couldn't reconcile ${slackLink("this request", pageUrl(pageId))} — ` +
            `${err instanceof Error ? err.message : String(err)}`,
        );
        throw err; // let the platform retry the delivery
      }
    }
  },
});
