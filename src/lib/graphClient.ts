/**
 * Everything that talks to Microsoft Graph funnels through here, so auth,
 * rate-limit pacing, timeouts, retry, and error shape stay consistent —
 * the same role kantataClient.ts / readClient.ts play in the sibling workers.
 *
 * AUTH DIFFERS from those two. They use `worker.oauth` (authorization code +
 * a one-time human consent) because they act as a user. This worker acts as
 * an *application*: the Entra app holds the Graph application permission
 * `Calendars.ReadWrite`, narrowed to one mailbox by an Exchange Online
 * Application Access Policy. That is the client-credentials grant — no
 * authorization endpoint, no redirect URI, no consenting user — which
 * `worker.oauth` cannot express. So we mint tokens ourselves against
 * `/{tenant}/oauth2/v2.0/token` and cache them in module scope.
 */

import type { JSONValue } from "@notionhq/workers/types";
import { graphPacer } from "../worker.js";
import { GraphApiError, InputValidationError, type GraphErrorKind } from "./errors.js";
import { graphBaseUrl, graphClientId, graphClientSecret, graphLoginBaseUrl, graphTenantId, pacerDisabled } from "./env.js";
import { withRetries } from "./retry.js";

const DEFAULT_TIMEOUT_MS = 20_000;
/** Refresh this long before the token actually expires, to absorb clock skew. */
const TOKEN_SKEW_MS = 60_000;

/**
 * The SDK pacer runtime state isn't initialized in every capability context
 * (`ntn workers exec` and webhook handlers both throw `Pacer "..." not found`
 * — reproduces on a clean scaffold, so it's a platform quirk, not ours).
 * Treat pacing as best-effort: pace where the runtime provides it, skip
 * otherwise — the retry/backoff in `withRetries` still absorbs any 429s.
 * GRAPH_DISABLE_PACER forces it off everywhere (local smoke testing).
 */
async function pace(): Promise<void> {
  if (pacerDisabled()) return;
  try {
    await graphPacer.wait();
  } catch (err) {
    if (!(err instanceof Error && err.message.includes("Pacer"))) throw err;
  }
}

// --- token ---

let cachedToken: { value: string; expiresAtMs: number } | null = null;

/**
 * Acquires an app-only access token via the client-credentials grant, cached
 * in module scope. The cache is best-effort: worker invocations may run in a
 * cold isolate, in which case this simply mints a fresh token (one extra
 * round trip, never a correctness problem).
 */
export async function accessToken(): Promise<string> {
  if (cachedToken && cachedToken.expiresAtMs > Date.now()) return cachedToken.value;

  const url = `${graphLoginBaseUrl()}/${encodeURIComponent(graphTenantId())}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: graphClientId(),
    client_secret: graphClientSecret(),
    // `.default` asks for every application permission already consented on
    // the app registration — here, just Calendars.ReadWrite.
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials",
  });

  return withRetries(async () => {
    const controller = new AbortController();
    const timeoutHandle = setTimeout(() => controller.abort(), DEFAULT_TIMEOUT_MS);

    let response: Response;
    try {
      response = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded", Accept: "application/json" },
        body,
        signal: controller.signal,
      });
    } catch (err) {
      throw new GraphApiError({
        kind: "network",
        message:
          err instanceof Error && err.name === "AbortError"
            ? `Token request to Entra timed out after ${DEFAULT_TIMEOUT_MS}ms.`
            : err instanceof Error
              ? err.message
              : "Network request to Entra failed.",
        endpoint: url,
        method: "POST",
      });
    } finally {
      clearTimeout(timeoutHandle);
    }

    const text = await response.text();
    let data: Record<string, unknown> = {};
    try {
      data = text ? (JSON.parse(text) as Record<string, unknown>) : {};
    } catch {
      /* fall through to the status-based message */
    }

    if (!response.ok) {
      // Entra token errors are `{ error, error_description }`. AADSTS7000215
      // (invalid client secret) and AADSTS700016 (unknown app) are the two
      // you'll actually hit — both are permanent, so classify as auth.
      const description = typeof data.error_description === "string" ? data.error_description : null;
      const code = typeof data.error === "string" ? data.error : null;
      throw new GraphApiError({
        kind: response.status >= 500 ? "server" : "auth",
        message: description ?? `Entra token request failed: HTTP ${response.status}`,
        status: response.status,
        code,
        endpoint: url,
        method: "POST",
      });
    }

    const token = typeof data.access_token === "string" ? data.access_token : null;
    if (!token) {
      throw new GraphApiError({
        kind: "unexpected",
        message: "Entra token response contained no access_token.",
        status: response.status,
        endpoint: url,
        method: "POST",
      });
    }
    const expiresInSec = typeof data.expires_in === "number" ? data.expires_in : 3600;
    cachedToken = { value: token, expiresAtMs: Date.now() + expiresInSec * 1000 - TOKEN_SKEW_MS };
    return token;
  });
}

/** Drops the cached token. Used after a 401, so the retry mints a fresh one. */
export function invalidateToken(): void {
  cachedToken = null;
}

// --- requests ---

type QueryValue = string | number | boolean | undefined | null;
export type QueryParams = Record<string, QueryValue>;

function buildUrl(path: string, query?: QueryParams): string {
  // Graph paginates with an absolute `@odata.nextLink`; pass those through
  // untouched rather than re-resolving them against the base URL.
  const url = /^https?:\/\//i.test(path) ? new URL(path) : new URL(path.replace(/^\/+/, ""), `${graphBaseUrl()}/`);
  if (query) {
    for (const [key, value] of Object.entries(query)) {
      if (value === undefined || value === null) continue;
      url.searchParams.set(key, String(value));
    }
  }
  return url.toString();
}

function classifyStatus(status: number): GraphErrorKind {
  if (status === 401) return "auth";
  if (status === 403) return "forbidden";
  if (status === 404) return "not_found";
  if (status === 400 || status === 422) return "validation";
  if (status === 429) return "rate_limited";
  if (status >= 500) return "server";
  return "unexpected";
}

/**
 * Graph errors are `{ error: { code, message, innerError } }`. Returns the
 * human message plus the machine code, which is what distinguishes an
 * Application Access Policy denial (`ErrorAccessDenied`) from a missing
 * permission grant (`Authorization_RequestDenied`).
 */
function parseErrorBody(text: string, status: number, statusText: string): { message: string; code: string | null } {
  try {
    const data = JSON.parse(text) as { error?: { code?: unknown; message?: unknown } };
    const err = data?.error;
    if (err) {
      const code = typeof err.code === "string" ? err.code : null;
      const message = typeof err.message === "string" ? err.message : statusText;
      return { message: code ? `${code}: ${message}` : message, code };
    }
  } catch {
    /* fall through */
  }
  return { message: statusText || `HTTP ${status}`, code: null };
}

export interface GraphRequestOptions {
  method?: "GET" | "POST" | "PATCH" | "DELETE";
  query?: QueryParams;
  body?: unknown;
  timeoutMs?: number;
  /** Extra headers, e.g. Prefer: outlook.timezone. */
  headers?: Record<string, string>;
}

/**
 * Low-level request: auth injection, pacing, timeout, retry on
 * network/429/5xx, and normalization of every failure into a GraphApiError.
 * A 401 drops the cached token before rethrowing, so the next call re-mints
 * (a secret rotation shouldn't need a redeploy to recover from).
 */
export async function graphRequest<T = unknown>(path: string, options: GraphRequestOptions = {}): Promise<T> {
  const { method = "GET", query, body, timeoutMs = DEFAULT_TIMEOUT_MS, headers = {} } = options;
  const url = buildUrl(path, query);

  return withRetries(async () => {
    await pace();
    const token = await accessToken();

    const controller = new AbortController();
    const timeoutHandle = setTimeout(() => controller.abort(), timeoutMs);

    let response: Response;
    try {
      response = await fetch(url, {
        method,
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
          Accept: "application/json",
          ...headers,
        },
        body: body !== undefined ? JSON.stringify(body) : undefined,
        signal: controller.signal,
      });
    } catch (err) {
      throw new GraphApiError({
        kind: "network",
        message:
          err instanceof Error && err.name === "AbortError"
            ? `Request to Microsoft Graph timed out after ${timeoutMs}ms.`
            : err instanceof Error
              ? err.message
              : "Network request to Microsoft Graph failed.",
        endpoint: url,
        method,
      });
    } finally {
      clearTimeout(timeoutHandle);
    }

    if (response.ok) {
      if (response.status === 204) return null as T;
      const text = await response.text();
      return (text ? JSON.parse(text) : null) as T;
    }

    if (response.status === 401) invalidateToken();

    let retryAfterMs: number | null = null;
    const retryAfterHeader = response.headers.get("Retry-After");
    if (retryAfterHeader) {
      const seconds = Number(retryAfterHeader);
      if (!Number.isNaN(seconds)) retryAfterMs = seconds * 1000;
    }

    const { message, code } = parseErrorBody(await response.text(), response.status, response.statusText);
    throw new GraphApiError({
      kind: classifyStatus(response.status),
      message,
      status: response.status,
      code,
      retryAfterMs,
      endpoint: url,
      method,
    });
  });
}

/**
 * Wraps a tool's calls so failures come back as a structured
 * `{ ok: false, error: {...} }` result instead of an uncaught exception —
 * the agent always gets JSON it can reason about, never a stack trace.
 * Same contract as the Kantata worker's `safeExecute`.
 */
export async function safeExecute<T extends JSONValue>(
  fn: () => Promise<T>,
): Promise<T | ReturnType<GraphApiError["toResult"]> | ReturnType<InputValidationError["toResult"]>> {
  try {
    return await fn();
  } catch (err) {
    if (err instanceof GraphApiError || err instanceof InputValidationError) return err.toResult();
    console.error("[graph] unexpected error in tool execution", err);
    return new GraphApiError({
      kind: "unexpected",
      message: err instanceof Error ? err.message : "An unexpected error occurred.",
      endpoint: "unknown",
      method: "unknown",
    }).toResult();
  }
}
