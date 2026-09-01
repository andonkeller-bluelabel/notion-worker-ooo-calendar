/**
 * Shared Worker instance and Microsoft Graph pacer.
 *
 * A Worker project has exactly one Worker instance; every capability
 * (webhook / sync / tool) imports `worker` from here and registers itself as a
 * side-effect. `src/index.ts` imports them all, then `export default worker`.
 *
 * NOTE — no `worker.oauth` here, unlike the Kantata / Read.ai / DPPT workers.
 * Those three authenticate as a *user* (OAuth 2.0 authorization code + one-time
 * browser consent via `ntn workers oauth start`). This worker authenticates as
 * an *application*: the Entra app uses the client-credentials grant, which has
 * no authorization endpoint, no redirect URI, and no consenting user — and
 * `worker.oauth` requires all three (see UserManagedOAuthConfiguration in
 * @notionhq/workers). Tokens are therefore fetched and cached directly in
 * lib/graphClient.ts.
 */

import { Worker } from "@notionhq/workers";

export const worker = new Worker();

/**
 * Microsoft Graph throttles per-app per-tenant on a sliding window (outlook
 * resources are documented at 10,000 requests per 10 minutes per app per
 * mailbox, plus a concurrency cap of 4). This worker's traffic is tiny — a
 * handful of calls per approval — so 4 requests/second is a conservative
 * shared budget across the webhook, the sweep, and the tools.
 */
export const graphPacer = worker.pacer("microsoftGraph", {
  allowedRequests: 4,
  intervalMs: 1000,
});
