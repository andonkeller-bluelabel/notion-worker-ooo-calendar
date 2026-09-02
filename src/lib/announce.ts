/**
 * Decides what to say in Slack about a request, and when to stay quiet.
 *
 * Pure and unit-tested: no Slack client, no Notion client. The reconciler
 * holds no memory of previous runs, so "did this change?" is answered by
 * comparing the row's current `Status` against the `Notified Status` the
 * worker last wrote. That comparison — not the calendar outcome — is what
 * makes `Requested → Denied` announceable, since it moves no event.
 */

import { ApprovalStatus, CALENDAR_STATUSES, RequestType } from "./schema.js";
import { blockedReason, type OooRequest } from "./oooRequest.js";

const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

/**
 * Human date range: "Sep 10" · "Sep 10–14" · "Sep 30 – Oct 2".
 * Years appear only when the range crosses one, since almost none do.
 */
export function formatRange(start: string, end: string): string {
  const [sy, sm, sd] = start.split("-").map(Number) as [number, number, number];
  const [ey, em, ed] = end.split("-").map(Number) as [number, number, number];
  const s = `${MONTHS[sm - 1]} ${sd}`;
  if (sy !== ey) return `${s}, ${sy} – ${MONTHS[em - 1]} ${ed}, ${ey}`;
  if (sm !== em) return `${s} – ${MONTHS[em - 1]} ${ed}`;
  if (sd !== ed) return `${s}–${ed}`;
  return s;
}

/** Escapes the three characters Slack treats as markup. */
function esc(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

/** Slack `<url|label>`, with the label escaped. */
function link(label: string, url: string): string {
  return `<${url}|${esc(label).replace(/\|/g, "-")}>`;
}

/** Longest note we will repeat into Slack before trimming. */
const MAX_NOTE = 300;

/**
 * The nudge for a row nobody can act on yet, or "".
 *
 * The Flex form is public, so it cannot pre-assign an approver — every request
 * through it arrives unassigned. Without this the channel notice would report a
 * request and give no one a reason to pick it up, and the approver DM cannot
 * fire because there is nobody to send it to.
 *
 * Only while the row is still undecided. An approved or denied row needs no
 * approver assigned after the fact.
 */
function actionLine(request: OooRequest): string {
  const decided = request.status === ApprovalStatus.DENIED || CALENDAR_STATUSES.includes(request.status ?? "");
  // A decided row that cannot become an event is the more urgent ask: it looks
  // handled and is not.
  if (decided && blockedReason(request)) {
    return `\n\n*ACTION:* ${link("Add the dates.", request.pageUrl)}`;
  }
  if (decided || request.approverId) return "";
  // The link rides on the action rather than sitting separately, so the message
  // carries exactly one link to the row and it is on the thing to click.
  return `\n\n*ACTION:* ${link("Assign an approver.", request.pageUrl)}`;
}

/**
 * The row's Notes, as a second line, or "" when there are none.
 *
 * Deliberately NOT on the calendar event — see graphCalendar.ts. Slack is a
 * narrower audience than a company-wide calendar, so repeating it here is a
 * different call from putting it in an event body, but it is still a broadcast:
 * whatever someone types into Notes reaches everyone in the channel.
 */
function noteLine(notes: string | null): string {
  const text = (notes ?? "").trim();
  if (!text) return "";
  const trimmed = text.length > MAX_NOTE ? `${text.slice(0, MAX_NOTE - 1)}…` : text;
  return `\nNote: ${esc(trimmed)}`;
}

export interface Announcement {
  /** Message text for Slack. */
  text: string;
  /** The status value to record as announced. */
  status: string;
}

/**
 * What to post about `request`, or null to stay quiet.
 *
 * Quiet when: the status has not changed since the last announcement, the row
 * has no status at all, or the page is in the trash (its removal is the
 * calendar's business, not something to narrate).
 *
 * `hadEvent` says whether a calendar event existed before this reconcile, so
 * an un-approval can mention that the entry came off the calendar.
 */
export function announcementFor(request: OooRequest, hadEvent: boolean): Announcement | null {
  const action = actionLine(request);
  // One link to the row per message. When there is an action, the link belongs
  // on it; otherwise the message carries its own "details".
  const announcement = buildAnnouncement(request, hadEvent, action ? "" : link("details", request.pageUrl));
  if (!announcement || !announcement.text) return announcement;
  return { ...announcement, text: announcement.text.trimEnd() + noteLine(request.notes) + action };
}

function buildAnnouncement(request: OooRequest, hadEvent: boolean, open: string): Announcement | null {
  const status = request.status;
  if (!status || request.inTrash) return null;
  if (status === request.notifiedStatus) return null;

  const who = `*${request.calendarName}*`;

  // Approved or scheduled, but unusable. Say so plainly rather than claiming a
  // calendar entry that the reconciler refused to create.
  //
  // The recorded status carries a ":blocked" suffix so this announces once, and
  // filling the dates in — which leaves `Status` unchanged — still produces the
  // real approval notice rather than silence.
  const blocked = CALENDAR_STATUSES.includes(status) ? blockedReason(request) : null;
  if (blocked) {
    // Dedupe against the blocked key, not the bare status: the guard above
    // compares the raw value and would let this repeat every run.
    const key = `${status}:blocked`;
    if (key === request.notifiedStatus) return null;
    return {
      status: key,
      text:
        `:warning: ${who}'s ${request.type ?? "time off"} is ${status.toLowerCase()} but has ` +
        `${blocked} — nothing goes on the team calendar until that's fixed.`,
    };
  }

  const when = request.startDate && request.endDate ? formatRange(request.startDate, request.endDate) : "dates not set";
  const firstTime = !request.notifiedStatus;
  const isTravel = request.type === RequestType.TRAVEL;

  switch (status) {
    case ApprovalStatus.SCHEDULED:
      return { status, text: `:airplane: ${who} added work related travel, ${when}. ${open}` };
    case ApprovalStatus.APPROVED:
      // Travel is announced rather than granted, so it never reads as an
      // approval — nobody approved it, and saying so would be misleading.
      return isTravel
        ? { status, text: `:airplane: ${who} added work related travel, ${when}. ${open}` }
        : { status, text: `:white_check_mark: ${who} is off ${when} — approved and on the team calendar. ${open}` };
    case ApprovalStatus.DENIED:
      return {
        status,
        text:
          `:x: ${who}'s ${isTravel ? "work related travel" : "time off"} for ${when} was ` +
          `${isTravel ? "cancelled" : "denied"}.${hadEvent ? " Removed from the team calendar." : ""} ${open}`,
      };
    case ApprovalStatus.REQUESTED:
    case ApprovalStatus.PENDING:
      return firstTime
        ? { status, text: `:palm_tree: ${who} requested time off, ${when}. ${open}` }
        : {
            status,
            text:
              `:arrows_counterclockwise: ${who}'s ${isTravel ? "work related travel" : "time off"} for ` +
              `${when} is back to ${status}.${hadEvent ? " Removed from the team calendar." : ""} ${open}`,
          };
    default:
      // An unrecognized status option. Record it so we don't loop, say nothing.
      return { status, text: "" };
  }
}


/** A direct message to one person. */
export interface DirectMessage {
  /** Account email to resolve to a Slack user. */
  email: string;
  text: string;
}

/**
 * The "please review this" DM to a newly assigned approver, or null.
 *
 * Fires when the approver CHANGES, which no status transition accompanies —
 * hence its own marker rather than riding on `Notified Status`.
 *
 * Silent once a row is decided: assigning an approver to something already
 * approved, scheduled or denied is bookkeeping, not a request for action, and
 * asking someone to review a settled row wastes their time.
 */
export function approverRequestDm(request: OooRequest, processUrl: string): DirectMessage | null {
  const { approverId, approverEmail, status, notifiedApprover } = request;
  if (!approverId || !approverEmail) return null;
  if (approverId === notifiedApprover) return null;
  if (status === ApprovalStatus.DENIED || CALENDAR_STATUSES.includes(status ?? "")) return null;

  const when =
    request.startDate && request.endDate ? formatRange(request.startDate, request.endDate) : "dates not set";
  return {
    email: approverEmail,
    text:
      `*${request.calendarName}* has requested time off, ${when}. ${link("details", request.pageUrl)}` +
      noteLine(request.notes) +
      `\n\n*ACTION:*\nFollow our time-off process: ${link("Requesting time off", processUrl)}`,
  };
}

/**
 * The "you're approved" DM to the person taking the time off, or null.
 *
 * Only on APPROVED. Scheduled travel is not approved by anyone, so telling
 * someone their travel was approved would be a small lie.
 */
export function approvedDm(request: OooRequest): DirectMessage | null {
  if (request.status !== ApprovalStatus.APPROVED) return null;
  if (request.status === request.notifiedStatus) return null;
  if (!request.personEmail || request.inTrash) return null;

  const when =
    request.startDate && request.endDate ? formatRange(request.startDate, request.endDate) : "your time off";
  const kind = request.type ?? "time off";
  return {
    email: request.personEmail,
    text:
      `Your ${when} ${kind} is approved! :penguin_dance:\n` +
      ":calendar: Added to the `Out of Office` calendar in Outlook.",
  };
}
