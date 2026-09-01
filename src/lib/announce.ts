/**
 * Decides what to say in Slack about a request, and when to stay quiet.
 *
 * Pure and unit-tested: no Slack client, no Notion client. The reconciler
 * holds no memory of previous runs, so "did this change?" is answered by
 * comparing the row's current `Status` against the `Notified Status` the
 * worker last wrote. That comparison — not the calendar outcome — is what
 * makes `Requested → Denied` announceable, since it moves no event.
 */

import { ApprovalStatus, RequestType } from "./schema.js";
import type { OooRequest } from "./oooRequest.js";

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

/** Slack `<url|label>`, with the characters Slack treats as markup escaped. */
function link(label: string, url: string): string {
  return `<${url}|${label.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\|/g, "-")}>`;
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
  const status = request.status;
  if (!status || request.inTrash) return null;
  if (status === request.notifiedStatus) return null;

  const who = `*${request.calendarName}*`;
  const when = request.startDate && request.endDate ? formatRange(request.startDate, request.endDate) : "dates not set";
  const open = link("details", request.pageUrl);
  const firstTime = !request.notifiedStatus;
  const isTravel = request.type === RequestType.TRAVEL;

  switch (status) {
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
