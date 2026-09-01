# notion-worker-ooo-calendar

A [Notion Workers](https://developers.notion.com/) worker that keeps the shared Microsoft 365
calendar for BlueLabel team time off in step with the **OOO Entries** Notion database. Approve a time-off request in Notion and the all-day out-of-office block appears
on the shared calendar; un-approve it, deny it, or trash the row and the block goes away.

Separate repo and deployment from the Kantata and Read.ai workers. It shares their conventions
(`@notionhq/workers`, the `ntn` CLI, ESM/NodeNext, `lib/errors` + `lib/retry` + lazy `lib/env`,
`node:test`) and differs from them in exactly two places, both documented below: how it
authenticates, and the fact that it needs a scheduled sweep to see deletions.

## How it works

```
Notion automation on OOO Entries         ── page added / edited → "Send webhook"
  → worker.webhook oooRequestChanged
      • the payload's `properties` is EMPTY — take the page id, re-read the row
      • reconcile(row)

  → worker.sync oooReconcileSweep  ── every 10m
      pass 1  every row that is Approved OR still carries an event id → reconcile(row)
      pass 2  every worker-owned event in the window whose row is gone → delete it
              (skipped entirely if any row in pass 1 failed)

reconcile(row):
  Approved?  ┬─ no event id  → create the event, write its id to O365 Event ID
             ├─ has event id → patch the event   (404 → recreate, store the new id)
             └─ no dates yet → do nothing, alert a human, retry next sweep
  otherwise  ┬─ no event id  → nothing to do
             └─ has event id → delete the event (404 is fine), clear O365 Event ID
```

One idempotent function, run on every delivery regardless of which property changed. It does not
branch on transitions; it makes the calendar match whatever the row currently says. That is what
makes replays, out-of-order deliveries, and the sweep all safe against the same row.

Events are titled `✈️ <first name>` (the plane is the house shorthand for "away"), and are
`isAllDay: true`, `showAs: "free"`, **no attendees** (the shared calendar is already the audience;
an attendee would get mailed an invite and a copy on their personal calendar), and
`isReminderOn: false`.

`showAs` is `free`, not `oof`, so that subscribing to this calendar never makes a viewer look busy
or away themselves.

The name comes from `BlueLabeler` when it is set, since a person picker holds what someone actually
goes by ("Danny", not "Daniel"), cannot be typo'd, and works for addresses with no separator to
split on. It falls back to the email's local part, which keeps rows syncing for people who are not
Notion workspace members. Notion profile names are title-cased on the way out, because at least one
is stored all-lowercase and that reads as a bug on a shared calendar.

First names only, because that is how the team refers to each other. When two people share one, add
it to `OOO_DISAMBIGUATE_FIRST_NAMES` and both get full names from the next sweep on. That list is
explicit rather than inferred from the database, because inferring would make a name silently
change form the day someone new requested time off, and would still miss anyone who had not
requested any yet.

## Two deliberate departures from the sibling workers

**1. App-only auth, not `worker.oauth`.** Kantata, Read.ai, and DPPT all act as a *user* and use
`worker.oauth` (authorization code + a one-time browser consent via `ntn workers oauth start`).
This worker acts as an *application*: client-credentials grant, no authorization endpoint, no
redirect URI, no consenting user. `worker.oauth` requires all three, so tokens are minted and
cached directly in `src/lib/graphClient.ts` against `/{tenant}/oauth2/v2.0/token`. Everything
downstream of that — pacer, retry, error normalization, `safeExecute` — is unchanged.

**2. Deletions need the sweep, not a webhook.** A Notion automation can only fire on a page being
*added* or *edited*. There is no delete trigger, so a trashed request would otherwise leave its
event on the calendar forever. `oooReconcileSweep` is the implementation of that requirement, not
an optimization. It also catches the failure mode the Read.ai worker hit in production: someone
sets `Approval Status = Approved` *before* filling in the dates, the automation fires once on the
status change against a half-filled row, and no second delivery ever comes.

To let the sweep trace an event back to its row, every event this worker creates carries the Notion
page id in a named MAPI property (`NOTION_PAGE_ID_PROP` in `lib/graphCalendar.ts`), with the body's
Notion link as a fallback. **Untagged events are never touched**, so anything a person put on the
calendar by hand is safe.

## Prerequisites

- An **Entra ID app registration** with the Microsoft Graph **application** permission
  `Calendars.ReadWrite`, admin-consented, and narrowed by an **Exchange Online Application Access
  Policy** to the target mailbox only. Verify with `Test-ApplicationAccessPolicy` before deploying;
  a policy that doesn't cover the mailbox shows up here as `ErrorAccessDenied`.
- A **shared mailbox** for the calendar, added to the policy's scope group. It must be a user or
  shared mailbox, **not** a Microsoft 365 group (see Platform notes for why that isn't a choice).
- **OOO Entries as a native Notion database** — see Platform notes.
- The `ntn` CLI and **Node 22** (`nvm use`).
- An internal Notion integration with access to that database (its token → `NOTION_API_TOKEN`).

## Notion database

**OOO Entries** (Team Ops → More) is user-owned and native. Requests arrive through its
"Request Time Off" form view, which is why the requester is an email column rather than a `people`
property. The worker reads every column and writes exactly one.

| Property | Type | Worker access |
|---|---|---|
| Name | title | read (display-name fallback) |
| BlueLabeler | people | read (**preferred** identity: the display name shown on the calendar) |
| Your Email | email | read (identity fallback when BlueLabeler is empty) |
| Dates | date **range** | read (`start` = first day off, `end` = **inclusive** last day off) |
| Status | **status** — Pending, Requested, Approved, Denied | read |
| Approver | people, optional | read (named in the event body) |
| Notes | rich_text, optional | read (stays in Notion — never written to the calendar) |
| **O365 Event ID** | rich_text, hidden | **read + write — worker-owned** |
| **Notified Status** | rich_text, hidden | **read + write — worker-owned** |

The worker rewrites `Title` to match the calendar subject (`✈️ <name>`) on every reconcile, whatever
the status, so the two views never disagree. It skips the write when the title already matches, when
the page is trashed, and when no identity resolved — in that last case the name is derived FROM the
title, and writing a computed one back would append a marker on every pass. The away marker is also
stripped when the title is read as a fallback, as a second guard against that loop.

**Keep `Title`, `O365 Event ID`, and `Notified Status` out of the automation triggers.** The worker
writes all three, so a trigger on any of them would feed deliveries back to itself.

## Slack notifications

Every reconcile posts to `OOO_NOTIFY_CHANNEL` when a row's `Status` differs from the `Notified
Status` the worker last recorded: a new submission, an approval, a denial, or a reversion. Errors go
to `OOO_OPS_CHANNEL` separately. Both need `SLACK_BOT_TOKEN` and the bot invited to the channel;
without them the worker runs normally and the messages go to the run logs instead.

`Notified Status` is why the notifier is correct rather than approximate. The reconciler holds no
memory between runs, so without it `Requested → Denied` — which moves no calendar event — would
leave no trace to notify on, and the sweep would either re-announce every approved row every ten
minutes or miss transitions entirely.

Announcing happens AFTER the calendar work, so a message never claims something that failed. A Slack
outage cannot break the sync: the failure is logged and the row keeps its old `Notified Status`, so
the next run tries again rather than going quiet.

**When adding this to a database that already has rows, backfill `Notified Status` to each row's
current status before deploying.** Otherwise the first run treats every existing row as new and
announces all of them at once.

`Status` is a status-type property, not a select, so data-source filters use the `status` key.
Notion cannot create status options through the API; we only read them.

`Dates` is a single date **range**, replacing the earlier separate `Start Date` / `End Date`
columns. Notion leaves `end` empty when someone picks one day in the picker, so an empty `end` means
a one-day request, not a missing value.

The calendar grants `Default: Reviewer`, so **every person in the tenant can read an event body**.
That is why `Notes` is read but never written to the event: it is free text on a form, and the
reasons people put in it belong in Notion, where approvers see them, not on a company-wide calendar.

There is deliberately **no Type column**. Putting "Sick" or "Personal" in a subject line broadcasts
it to everyone with calendar access, so every event reads `✈️ <first name>` and nothing more. To
reintroduce the distinction, add a `Type` select, read it in `toOooRequest`, and use it in
`eventSubject` (see `AWAY_MARKER` in `src/lib/schema.ts`).

Property names live in `src/lib/schema.ts`; a rename in Notion is a one-line change there.
`checkOooSetup` reports any drift and lists what the database actually has.

## Setup

```bash
nvm use
npm install
ntn workers deploy        # first deploy registers the worker + the sweep anchor DB
```

```bash
ntn workers env set GRAPH_TENANT_ID=... GRAPH_CLIENT_ID=... GRAPH_CLIENT_SECRET=...
ntn workers env set NOTION_API_TOKEN=... OOO_DATA_SOURCE_ID=...
ntn workers env set OOO_TIMEZONE=America/New_York
ntn workers deploy
```

**Point it at the mailbox.** No id resolution needed; a mailbox is addressed by its SMTP address.

```bash
ntn workers env set O365_CALENDAR_MAILBOX=teamooo@bluelabellabs.com && ntn workers deploy
```

**Smoke-test before wiring anything up.** This is the step that catches the two failures worth
catching early — a database whose columns the API can't write, and a Graph permission that doesn't
actually reach the calendar:

```bash
ntn workers exec checkOooSetup -d '{"samplePageId":"<any row in OOO Entries>"}'
```

Every field is required, including the nullable ones (the SDK's schema builder puts all object
properties in `required` by design and has no `.optional()`). Pass `null` explicitly to skip one:

```bash
ntn workers exec checkOooSetup -d '{"samplePageId": null}'
```

It reports the database schema against the properties above, round-trips `O365 Event ID`
(writing the current value back unchanged) to prove the column is writable, and fetches the
mailbox calendar. Read all three sections before continuing.

If the `graphCalendar` section fails, `showGraphTokenRoles` separates a permission problem from a
policy problem — both surface as `ErrorAccessDenied`:

```bash
ntn workers exec showGraphTokenRoles -d '{}'
```

**Then test the reconciler end to end**, still without any automation wired up:

```bash
ntn workers env set OOO_DRYRUN=true && ntn workers deploy
ntn workers exec reconcileOooRequest -d '{"pageId":"<an approved row>"}'   # says what it would do
ntn workers env set OOO_DRYRUN=false && ntn workers deploy
ntn workers exec reconcileOooRequest -d '{"pageId":"<an approved row>"}'   # creates the event
```

**Wire the trigger last.** On OOO Entries, add a Notion automation: *When a page is added or
edited → Send webhook* to the `oooRequestChanged` URL (`ntn workers webhooks list`). Scope it to
the whole row rather than a single property — the reconciler is idempotent, so extra deliveries
cost nothing, while a missed date edit after approval leaves a wrong event on the calendar.

## Testing

```bash
npm test        # node:test via tsx — pure logic, no network
npm run typecheck
```

41 tests across three files, all pure-function with injected dependencies (the pattern
`lib/router.test.ts` uses in the Read.ai worker):

- `lib/reconcile.test.ts` — the create / update / delete decision table, idempotency under repeated
  runs, recreate-on-404, clear-the-id-even-when-the-delete-404s, dry run, and the known duplicate
  gap asserted so the behaviour stays deliberate.
- `lib/oooRequest.test.ts` — date arithmetic (month/year/leap-day/DST boundaries), the inclusive →
  exclusive end conversion, row reading, and the approved-but-dateless guard.
- `lib/graphCalendar.test.ts` — the event payload's four load-bearing decisions and the page-id tag.

## Known edge cases

- **Duplicate events.** Two routes produce one: concurrent webhook deliveries (two Notion
  automations firing on a single change, both reading an empty `O365 Event ID`, both creating an
  event — observed 2026-09-01), or someone clearing that column by hand while the row is still
  Approved. The reconciler can't recognize an existing event from the row alone, so it creates a
  second. The sweep resolves it: an owned event whose Approved row names a *different* event is
  stale by definition, since Notion is the source of truth for which event is current, so it is
  deleted and a Slack notice records it. Worst case is a duplicate visible for up to 10 minutes.
  The create-anyway behaviour is asserted in `reconcile.test.ts` so it can't change silently.
- **Avoid stacking automations.** Use ONE Notion automation with several triggers (`Page added`,
  plus property triggers on `Status`, `Start Date`, `End Date`) rather than several automations.
  Separate automations fire concurrently on the same change and race each other.
- **Someone edits the event in Outlook.** The next reconcile overwrites it. The event body says so.
- **A row approved with no dates.** No event is created — a wrong all-day block is worse than a
  missing one. The webhook alerts once; the sweep retries hourly and stays quiet.
- **Hard deletion after 30 days in the trash.** The sweep treats a Notion `object_not_found` as a
  genuine orphan and removes the event. Anything else it can't confirm, it leaves alone.

## Platform notes / gotchas

- **OOO Entries must be a NATIVE database.** A `worker.database` (managed) database's
  non-title properties are read-only to the API — writes come back 400 "read-only" — so the worker
  could never own `O365 Event ID` on one. This is the beta restriction the Kantata and Read.ai
  workers both hit; the Read.ai worker's pending queue is native for the same reason. The sweep's
  managed anchor DB exists **only** to host a schedule (the platform schedules work through
  `worker.sync`, and a sync must attach to a managed database) and stays empty.
- **`automation` capability is disabled for this workspace**, so the trigger is a `worker.webhook`
  called by a Notion "Send webhook" automation, not `worker.automation`.
- **"Send webhook" payloads have empty `properties`.** The body carries only the page reference, so
  the handler `pages.retrieve`s the row. A webhook event is a signal, not a snapshot.
- **Notion automations have no delete trigger.** Hence the sweep.
- **Notion's "Any property edited" trigger does not fire on STATUS-property changes.** Observed
  2026-09-01 on this database: three separate `Status` edits (Approved, Denied, Requested) produced
  zero webhook deliveries over ~35 minutes, while a `Notes` (rich_text) edit produced two within
  seconds. `Page added` fires normally. Since `Status` is the property that decides whether an event
  should exist, the webhook cannot be trusted for the one transition that matters — which is why
  `oooReconcileSweep` runs every 10 minutes and is the PRIMARY mechanism here, not a backstop. The
  webhook stays wired because it makes the common case fast when it does fire, and a redundant
  delivery costs one no-op log line.
- **Scheduled syncs need `NOTION_API_TOKEN`.** Webhooks get `context.notion` injected; syncs throw
  without the token.
- **Tool inputs have no optional fields.** The SDK's schema builder puts every object property in
  `required` and offers only `.nullable()`, never `.optional()`. Omitting a nullable field from an
  `ntn workers exec -d` payload is a 400 `InvalidToolInputError`; pass `null` instead. Harmless for
  agent-called tools, which always fill every field, but it catches you out running `exec` by hand.
- **`ntn workers exec` doesn't initialize the pacer runtime** (`Pacer "..." not found`, reproduces
  on a clean scaffold). `pace()` in `graphClient.ts` degrades gracefully; set
  `GRAPH_DISABLE_PACER=true` only if you still hit it, and never on prod.
- **Graph's all-day `end` is exclusive.** A vacation through Friday the 11th ends
  `2026-09-12T00:00:00`. Getting this wrong makes every event a day short; `allDayRange()` owns it
  and is tested.
- **App-only auth cannot reach a Microsoft 365 group calendar. At all.** Graph's reference for
  `POST /groups/{id}/calendar/events` lists Application permissions as **"Not supported"**;
  `Group.ReadWrite.All` is delegated-only for that endpoint, and there is no calendar
  resource-specific-consent permission to scope a grant to one team. The failure surfaces as
  `ErrorAccessDenied` from the Exchange backend, which reads exactly like a policy denial and sends
  you auditing a correct configuration. Confirmed the hard way here: the group object id was right,
  `Test-ApplicationAccessPolicy` returned Granted, and the token carried `Calendars.ReadWrite`.
  Hence the shared mailbox. It is also the only thing an Application Access Policy can scope, since
  those policies constrain mailbox permissions, not directory ones.
- **A group's Entra "Object ID" *is* its Graph group id**, despite the folklore. If you ever do need
  it, that field is the value; no separate lookup required.
- **A group's Graph `mail` property holds its PRIMARY SMTP**, which may be the
  `*.onmicrosoft.com` address while the branded address is a proxy alias. Filtering
  `?$filter=mail eq '<branded address>'` then returns nothing for a group that plainly exists.
- **`@notionhq/client` must be v5+** for the `dataSources` API used throughout.

## Error handling & reliability

- Every Graph call goes through `graphClient.ts`'s `graphRequest`, which retries network errors,
  429s, and 5xxs with exponential backoff + jitter (`lib/retry.ts`) — up to 5 attempts, honoring
  `Retry-After`. 401/403/404/400 are **not** retried. A 401 drops the cached token first, so a
  rotated client secret recovers without a redeploy.
- Errors are normalized into `GraphApiError` with Graph's machine code attached, which is what
  separates an Application Access Policy denial (`ErrorAccessDenied`) from a missing permission
  grant (`Authorization_RequestDenied`) from a mailbox that doesn't resolve (`ResourceNotFound`).
- Tools wrap their calls in `safeExecute()`, so an agent gets a structured
  `{ ok: false, error: { kind, message, status, code, retryable } }` instead of a stack trace.
- The sweep's cleanup pass is **default-deny**: if any row failed to reconcile, it removes nothing
  that run and says so in Slack.
- Slack notices are best-effort and never throw into the caller. Without `SLACK_BOT_TOKEN` /
  `OOO_OPS_CHANNEL` everything still works; the alerts just go to the run logs
  (`ntn workers runs logs`).

## This deployment (BlueLabel)

- Entra app `156735ca-2a91-4026-86b8-b3a612278048` in tenant
  `ac6518c5-7e0f-4868-8d67-c6d5aa98c290`, holding application `Calendars.ReadWrite`, scoped by an
  Exchange Application Access Policy to the mail-enabled security group **OOOSyncWorkerScope**.
- Calendar: a **shared mailbox** in that scope group. The events are surfaced to the team by
  granting `Reviewer` on its calendar folder, so people read it and only the worker writes.
- Superseded: **Holidays & Vacation Calendar (BlueLabel Team)**
  (`vacationcalendar@klink.onmicrosoft.com`, Entra object id
  `1dfe01f9-e346-471b-9f9a-bf55a80de647`) is the Microsoft 365 group this originally targeted. It
  is unreachable app-only, per Platform notes. Nothing was wrong with its configuration.
- Webhook URLs via `ntn workers webhooks list`; sync status via `ntn workers sync status`.
