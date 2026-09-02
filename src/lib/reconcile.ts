/**
 * The reconciler — one idempotent function, run on every relevant delivery
 * regardless of which property changed. It does not branch on transitions
 * ("was Requested, now Approved"); it makes the calendar match whatever the
 * row currently says. That is what makes replays, out-of-order deliveries,
 * and the sweep all safe to run against the same row.
 *
 * Dependencies are injected so the whole decision table is unit-testable with
 * no network — the pattern lib/router.ts uses in the Read.ai worker.
 */

import { autoApproves, isApproved, blockedReason, eventSubject, type OooRequest } from "./oooRequest.js";
import { ApprovalStatus } from "./schema.js";
import { announcementFor, approvedDm, approverRequestDm } from "./announce.js";

export interface ReconcileDeps {
  createEvent: (request: OooRequest) => Promise<{ id: string }>;
  updateEvent: (eventId: string, request: OooRequest) => Promise<void>;
  deleteEvent: (eventId: string) => Promise<void>;
  storeEventId: (pageId: string, eventId: string) => Promise<void>;
  clearEventId: (pageId: string) => Promise<void>;
  /** Writes the row's `Title` so Notion reads the same as the calendar. */
  setTitle: (pageId: string, title: string) => Promise<void>;
  /**
   * Posts a routine notification to the team channel. Returns whether it
   * actually landed — a failed post must not be recorded as announced.
   */
  announce: (text: string) => Promise<boolean>;
  /** Records the status just announced, so it is not announced twice. */
  setNotifiedStatus: (pageId: string, status: string) => Promise<void>;
  /** Moves a row's `Status`. Used only to auto-approve Work Related Travel. */
  setStatus: (pageId: string, status: string) => Promise<void>;
  /** DMs one person by account email. Returns whether it landed. */
  dm: (email: string, text: string) => Promise<boolean>;
  /** Records the approver just DM'd, so they are not asked twice. */
  setNotifiedApprover: (pageId: string, approverId: string) => Promise<void>;
  /** The team's time-off process page, linked from the approver's DM. */
  processUrl: string;
  /** True when the error from updateEvent/deleteEvent means "no such event". */
  isNotFound: (err: unknown) => boolean;
  /** Best-effort operational notice. Never throws into the reconciler. */
  notify?: (text: string) => Promise<void>;
  /** Log the decision, perform no writes. */
  dryRun?: boolean;
}

export type ReconcileAction =
  /** Approved, no event id yet → event created and the id written back. */
  | "created"
  /** Approved with an event id → event patched to match the row. */
  | "updated"
  /** Approved, but the stored event was gone from the calendar → made a new one. */
  | "recreated"
  /** Not approved (or trashed) with an event id → event removed, id cleared. */
  | "deleted"
  /** Not approved with an event id, but the event was already gone. Id cleared. */
  | "already-absent"
  /** Not approved and no event id → nothing to do. The common case. */
  | "noop"
  /** Approved but unusable (missing/invalid dates). Left for the next sweep. */
  | "blocked"
  /** dryRun was on; nothing was written. */
  | "dry-run";

export interface ReconcileResult {
  action: ReconcileAction;
  /** The event id after reconciling — null when no event should exist. */
  eventId: string | null;
  /** Set on "blocked" and "dry-run": what happened, in words. */
  reason?: string;
}

/**
 * Brings the shared calendar in line with one OOO Entries row.
 *
 * KNOWN EDGE CASE (accepted, not engineered around): if someone clears the
 * `O365 Event ID` cell by hand while the row is still Approved, this creates a
 * second event rather than adopting the existing one. The sweep will not undo
 * it either, because the orphaned event still points at a live, approved page.
 * Symptom is a duplicate on the calendar; the fix is to delete one by hand.
 */
export async function reconcile(input: OooRequest, deps: ReconcileDeps): Promise<ReconcileResult> {
  // Work Related Travel needs no approval, so move it to SCHEDULED before
  // anything reads the status — the calendar branch and the Slack notice must
  // both see the state the row is about to be in, not the one the form left.
  let request = input;
  if (autoApproves(input)) {
    if (!deps.dryRun) await deps.setStatus(input.pageId, ApprovalStatus.SCHEDULED);
    request = { ...input, status: ApprovalStatus.SCHEDULED };
  }

  const result = await reconcileCalendar(request, deps);
  await announceIfChanged(request, deps);
  await dmApproverIfChanged(request, deps);
  return result;
}

/**
 * Asks a newly assigned approver to review the row.
 *
 * Recorded only when the DM lands, for the same reason the channel notice is:
 * marking it sent regardless loses the request for good, and an approver who
 * is never asked is a request that sits forever.
 */
async function dmApproverIfChanged(request: OooRequest, deps: ReconcileDeps): Promise<void> {
  const message = approverRequestDm(request, deps.processUrl);
  if (!message || deps.dryRun || request.inTrash) return;
  try {
    if (!(await deps.dm(message.email, message.text))) {
      console.warn(`[reconcile] couldn't DM the approver for ${request.pageId}; leaving it to retry`);
      return;
    }
    await deps.setNotifiedApprover(request.pageId, request.approverId!);
  } catch (err) {
    console.error(`[reconcile] approver DM failed for ${request.pageId}:`, err);
  }
}

/**
 * Posts a Slack notice when the row's status differs from the one last
 * announced, then records the new one. Runs AFTER the calendar work, so a
 * message never claims something that failed to happen.
 *
 * `Notified Status` is recorded ONLY when the post actually landed. Recording
 * it regardless loses the message for good: the next run sees no change and
 * stays quiet forever. That is what a bad Slack token did here — the notice
 * was dropped, the row was marked announced, and nothing ever retried.
 *
 * A Slack failure still never breaks the calendar work; it is logged, the row
 * keeps its old value, and the next reconcile tries again.
 */
async function announceIfChanged(request: OooRequest, deps: ReconcileDeps): Promise<void> {
  const announcement = announcementFor(request, Boolean(request.eventId));
  if (!announcement || deps.dryRun) return;
  try {
    const posted = announcement.text ? await deps.announce(announcement.text) : true;
    if (!posted) {
      console.warn(`[reconcile] Slack rejected the notice for ${request.pageId}; leaving it to retry`);
      return;
    }
    // Tell the person their time off was approved. Best-effort and deliberately
    // NOT part of the gate: if this failed and blocked the record, the channel
    // notice would repeat on every run. A missed DM is logged instead.
    const personal = approvedDm(request);
    if (personal) {
      const sent = await deps.dm(personal.email, personal.text);
      if (!sent) console.warn(`[reconcile] couldn't DM ${personal.email} about ${request.pageId}`);
    }
    await deps.setNotifiedStatus(request.pageId, announcement.status);
  } catch (err) {
    console.error(`[reconcile] couldn't announce ${request.pageId}:`, err);
  }
}

async function reconcileCalendar(request: OooRequest, deps: ReconcileDeps): Promise<ReconcileResult> {
  const { pageId, eventId } = request;

  // Keep the Notion title in step with the calendar subject, whatever the
  // status. Skipped when we have no real identity (the name would be coming
  // from the title itself), for a trashed page (Notion rejects the write), and
  // when it already matches, so a sweep over unchanged rows writes nothing.
  const desiredTitle = eventSubject(request);
  if (request.hasIdentity && !request.inTrash && request.currentTitle !== desiredTitle && !deps.dryRun) {
    await deps.setTitle(pageId, desiredTitle);
  }

  if (isApproved(request)) {
    const blocked = blockedReason(request);
    if (blocked) {
      // Approving before the dates are filled in is a real pattern, and the
      // automation fires on the status change only — no second delivery is
      // coming. Say so loudly; the sweep retries the row on its next pass.
      await deps.notify?.(
        `OOO sync: ${request.personName}'s time off is Approved but has ${blocked}. No calendar event until that's fixed.`,
      );
      return { action: "blocked", eventId, reason: blocked };
    }

    if (!eventId) {
      if (deps.dryRun) return { action: "dry-run", eventId: null, reason: "would create an event" };
      const created = await deps.createEvent(request);
      await deps.storeEventId(pageId, created.id);
      return { action: "created", eventId: created.id };
    }

    if (deps.dryRun) return { action: "dry-run", eventId, reason: `would update event ${eventId}` };
    try {
      await deps.updateEvent(eventId, request);
      return { action: "updated", eventId };
    } catch (err) {
      if (!deps.isNotFound(err)) throw err;
      // Someone deleted the event straight off the calendar. The row is still
      // approved, so the calendar is what's wrong — put it back.
      const created = await deps.createEvent(request);
      await deps.storeEventId(pageId, created.id);
      return { action: "recreated", eventId: created.id };
    }
  }

  // Not approved: Requested, Denied, blank, or the page is in the trash.
  if (!eventId) return { action: "noop", eventId: null };

  if (deps.dryRun) return { action: "dry-run", eventId, reason: `would delete event ${eventId}` };

  let alreadyGone = false;
  try {
    await deps.deleteEvent(eventId);
  } catch (err) {
    if (!deps.isNotFound(err)) throw err;
    alreadyGone = true;
  }
  // Clear the id even when the delete 404'd — a stale id is worse than none,
  // since a later re-approval would try to patch an event that doesn't exist.
  // Skipped for a trashed page: Notion rejects property writes to the trash,
  // and the row is gone anyway.
  if (!request.inTrash) await deps.clearEventId(pageId);
  return { action: alreadyGone ? "already-absent" : "deleted", eventId: null };
}
