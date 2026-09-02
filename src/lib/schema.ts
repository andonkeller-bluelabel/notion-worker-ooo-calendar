/**
 * Property-name constants for the "OOO Entries" database (Team Ops → More).
 * Centralized here so a rename in Notion is a one-line change, matching the
 * Read.ai worker's lib/schema.ts.
 *
 * These are the database's REAL property names, confirmed against the live
 * schema with `checkOooSetup`. Requests arrive through the "Request Time Off"
 * form view, which is why the requester is identified by an email column
 * rather than a `people` property — changing that would break the form.
 *
 * This is a NATIVE, user-owned database, deliberately not a `worker.database`
 * managed one. A managed database's non-title columns are read-only to the API
 * (Notion Workers beta: writes come back 400 "read-only"), and this worker must
 * write `O365 Event ID` back onto the row. See README → Platform notes.
 */
export const Ooo = {
  /** title — whatever the requester called it. Display-name fallback only. */
  TITLE: "Title",
  /**
   * people — the person off. Preferred source of the display name, because it
   * carries what someone actually goes by ("Danny", not "Daniel"), can't be
   * typo'd the way a typed address can, and works for addresses with no
   * separator to split on.
   */
  BLUELABELER: "BlueLabeler",
  /**
   * email — what the request form collects. Fallback identity for rows where
   * BlueLabeler is empty: someone outside the Notion workspace, or a row
   * created before that column existed.
   */
  EMAIL: "Your Email",
  /**
   * date — a RANGE. `start` is the first day off; `end` is the INCLUSIVE last
   * day off, and is empty for a single day. Replaced the separate
   * "Start Date" / "End Date" columns.
   */
  DATES: "Dates",
  /** status (NOT select) — Pending | Requested | Approved | Denied. */
  STATUS: "Status",
  APPROVER: "Approver", // people, optional — named in the event body
  NOTES: "Notes", // rich_text, optional
  /** select — Paid Time Off | Work Related Travel. */
  TYPE: "Type",
  /** rich_text, hidden. Worker-owned: the Graph event id, or empty. */
  O365_EVENT_ID: "O365 Event ID",
  /**
   * rich_text, hidden. Worker-owned: the `Status` value we last announced in
   * Slack. The reconciler is otherwise stateless, so this is what lets it tell
   * a real transition from a re-run — including `Requested → Denied`, which
   * changes nothing on the calendar and would leave no other trace.
   */
  NOTIFIED_STATUS: "Notified Status",
  /**
   * rich_text, hidden. Worker-owned: the Notion user id of the approver we
   * last DM'd. `Notified Status` only tracks the status, so without this a
   * change of approver — which moves no status — would leave no trace and
   * either never notify or notify on every sweep.
   */
  NOTIFIED_APPROVER: "Notified Approver",
  /** created_by — the Notion user who submitted. Carries a verified email. */
  CREATED_BY: "Created by",
} as const;

/**
 * `Status` option names. Only APPROVED puts an event on the calendar; every
 * other value (including an empty one) means "no event should exist", which is
 * what makes the reconciler a single idempotent function rather than a set of
 * per-transition branches.
 *
 * Notion cannot create status-type options through the API, so these must
 * already exist on the database. We only ever read them.
 */
export const ApprovalStatus = {
  /** The form's default: nobody has looked at it yet. No calendar entry. */
  PENDING: "Pending",
  /** A human granted leave. On the calendar. */
  APPROVED: "Approved",
  /**
   * Happening, but nobody approved it — Work Related Travel is announced, not
   * requested. On the calendar exactly like APPROVED; the two differ only in
   * what they claim about how the row got there.
   */
  SCHEDULED: "Scheduled",
  /** Refused or cancelled. No calendar entry. */
  DENIED: "Denied",
  /**
   * No longer an option on the database, kept so a row still carrying it (or a
   * later decision to bring it back) behaves sensibly rather than silently.
   */
  REQUESTED: "Requested",
} as const;

/**
 * The statuses that mean "this is happening, put it on the calendar".
 *
 * Two positives rather than one. `Status` is on neither form, so only someone
 * with database access can set either — SCHEDULED is not a way for a requester
 * to skip approval, it is an approver saying a trip is going ahead.
 */
export const CALENDAR_STATUSES: readonly string[] = [ApprovalStatus.APPROVED, ApprovalStatus.SCHEDULED];

/**
 * Prefixed to the first name on every calendar event, e.g. "✈️ Andon".
 * The plane is the house shorthand for "away"; leading it makes a column of
 * entries scan as a list of icons rather than a ragged right edge.
 */
export const AWAY_MARKER = "✈️";

/**
 * `Type` option names.
 *
 * TRAVEL is auto-approved: a client onsite is something a person announces,
 * not something they ask permission for. See `autoApproves` in oooRequest.ts.
 */
export const RequestType = {
  PTO: "Paid Time Off",
  UNPAID: "Unpaid Time Off",
  TRAVEL: "Work Related Travel",
} as const;

/**
 * `Source` option names, one per form.
 *
 * CORE: the responder is signed in, so `Created by` identifies them.
 * FLEX: the responder may be anonymous, so the typed email is the identity and
 * `Created by` must NOT be trusted — it would attribute the request to whoever
 * or whatever Notion records as the creator of an anonymous submission.
 */
export const SourceForm = {
  CORE: "Core",
  FLEX: "Flex",
} as const;

/**
 * The two holiday databases the worker mirrors onto the shared calendar.
 *
 * These rows are not requests — nobody submits or approves a public holiday —
 * so they never enter OOO Entries. Each gets the same worker-owned
 * `O365 Event ID` column that OOO Entries has, so the pattern is identical and
 * the id is visible when something needs debugging.
 */
export const UsHolidays = {
  TITLE: "Holiday",
  DATE: "Date",
  O365_EVENT_ID: "O365 Event ID",
} as const;

export const VendorHolidays = {
  TITLE: "Name",
  DATE: "Date",
  /**
   * select — which partner. Currently empty on every row; the vendor is in the
   * title instead ("Tatvasoft: Diwali"), which is why nothing reads this yet.
   */
  PARTNER: "Partner",
  O365_EVENT_ID: "O365 Event ID",
} as const;

/** Prefixes that make the three kinds of entry readable at a glance. */
export const US_HOLIDAY_MARKER = "🇺🇸";
export const VENDOR_HOLIDAY_MARKER = "🏢";
