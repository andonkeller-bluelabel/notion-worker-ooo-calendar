/**
 * Centralized env access. Every value is read through a lazy getter (a
 * function called at use-time, not module load) because the Notion Workers
 * runtime injects env per capability invocation — reading `process.env` at
 * import time can see stale/empty values.
 *
 * Secrets are set with `ntn workers env set` (see README + .env.example) and
 * never live in the repo. `NOTION_API_TOKEN` is needed only by the sweep:
 * scheduled `worker.sync` is NOT auto-injected with Notion access, whereas
 * `worker.webhook` gets an authenticated client as `context.notion`.
 */

/** Reads a required env var; throws a clear error naming the missing key. */
export function requireEnv(name: string): string {
  const value = process.env[name];
  if (!value || !value.trim()) {
    throw new Error(`Missing required env var ${name}. Set it with \`ntn workers env set ${name}=...\`.`);
  }
  return value.trim();
}

// --- Microsoft Graph (app-only / client credentials) ---
export const graphTenantId = () => requireEnv("GRAPH_TENANT_ID");
export const graphClientId = () => requireEnv("GRAPH_CLIENT_ID");
export const graphClientSecret = () => requireEnv("GRAPH_CLIENT_SECRET");
export const graphBaseUrl = () => (process.env.GRAPH_BASE_URL ?? "https://graph.microsoft.com/v1.0").replace(/\/+$/, "");
export const graphLoginBaseUrl = () =>
  (process.env.GRAPH_LOGIN_BASE_URL ?? "https://login.microsoftonline.com").replace(/\/+$/, "");

/**
 * SMTP address of the mailbox whose calendar we write, e.g.
 * `teamooo@bluelabellabs.com`. Must be a USER or SHARED mailbox.
 *
 * Deliberately not a Microsoft 365 group. Graph does not support app-only
 * (application permission) access to `/groups/{id}/calendar/events` at all —
 * the reference lists Application as "Not supported", and no permission grant
 * changes that. A shared mailbox is what an Exchange Application Access Policy
 * can actually scope `Calendars.ReadWrite` down to. See README → Platform notes.
 */
export const o365CalendarMailbox = () => requireEnv("O365_CALENDAR_MAILBOX");

// --- Notion ---
/** Data source id of the (native, user-owned) "OOO Entries" database. */
export const oooDataSourceId = () => requireEnv("OOO_DATA_SOURCE_ID");
/** Data source id of "BlueLabel US Holidays" (read-only mirror source). */
export const usHolidaysDataSourceId = () => process.env.US_HOLIDAYS_DATA_SOURCE_ID ?? "";
/** Data source id of "Vendor Partner Holidays" (read-only mirror source). */
export const vendorHolidaysDataSourceId = () => process.env.VENDOR_HOLIDAYS_DATA_SOURCE_ID ?? "";

/** Only the sweep needs this — webhooks get `context.notion` from the platform. */
export const notionApiToken = () => requireEnv("NOTION_API_TOKEN");

// --- Calendar behaviour ---
/**
 * IANA zone stamped on the all-day events. Graph requires a timeZone even for
 * all-day events; using the team's zone keeps the day boundaries where people
 * expect them rather than sliding by the UTC offset.
 */
export const oooTimezone = () => process.env.OOO_TIMEZONE ?? "America/New_York";

/**
 * First names that need a last name on the calendar because more than one
 * person shares them, e.g. "chris,sarah". Comma-separated, case-insensitive.
 *
 * Events read "<first name> ✈️" by default. Add a name here the day a second
 * Chris joins, and both of them become "Chris Boyle ✈️" / "Chris Ferrari ✈️"
 * on the next sweep. Empty (the default) means first names only.
 */
export const disambiguateFirstNames = () =>
  (process.env.OOO_DISAMBIGUATE_FIRST_NAMES ?? "")
    .split(",")
    .map((s) => s.trim().toLowerCase())
    .filter(Boolean);

// --- Sweep window ---
/** How far back the sweep scans the calendar for events to reconcile. */
export const sweepLookbackDays = () => {
  const n = Number(process.env.OOO_SWEEP_LOOKBACK_DAYS);
  return Number.isFinite(n) && n > 0 ? n : 30;
};
/** How far forward the sweep scans. Time off is rarely booked more than a year out. */
export const sweepLookaheadDays = () => {
  const n = Number(process.env.OOO_SWEEP_LOOKAHEAD_DAYS);
  return Number.isFinite(n) && n > 0 ? n : 400;
};

// --- Operational ---
/** Log what would happen, write nothing — to Notion or to the calendar. */
export const isDryRun = () => process.env.OOO_DRYRUN === "true";
/** Verbose delivery logging. */
export const isDebug = () => process.env.OOO_DEBUG === "true";
/** Force the rate-limit pacer off (local smoke testing only). */
export const pacerDisabled = () => process.env.GRAPH_DISABLE_PACER === "true";

// --- Slack (optional, best-effort) ---
export const slackBotToken = () => process.env.SLACK_BOT_TOKEN ?? "";
export const opsChannel = () => process.env.OOO_OPS_CHANNEL ?? "";
/**
 * Channel for routine request notifications (submitted / approved / denied).
 * Defaults to the ops channel, so one setting covers both until there is a
 * reason to split them.
 */
export const notifyChannel = () => process.env.OOO_NOTIFY_CHANNEL || opsChannel();
/** Link to the OOO Entries database, so ops alerts are actionable. */
export const oooDatabaseUrl = () => process.env.OOO_DATABASE_URL ?? "";

/**
 * The team's time-off process page, linked from the approver's "please review"
 * DM so the action and the instructions arrive together.
 */
export const processUrl = () =>
  process.env.OOO_PROCESS_URL ??
  "https://app.notion.com/p/bluelabellabs/Managing-your-time-off-OOO-9184eecdf9f746649c410e2d07d726fc";
