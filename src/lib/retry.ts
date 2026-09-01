import { GraphApiError } from "./errors.js";

export interface RetryOptions {
  maxAttempts?: number;
  baseDelayMs?: number;
  maxDelayMs?: number;
}

const DEFAULTS: Required<RetryOptions> = {
  maxAttempts: 5,
  baseDelayMs: 500,
  maxDelayMs: 15_000,
};

function sleep(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/**
 * Runs `fn` with exponential backoff + full jitter. Only retries errors
 * marked `.retryable` on a GraphApiError (network errors, 429s, 5xxs).
 * Everything else — auth failures, 403 policy denials, 404s, 400s —
 * propagates on the first attempt, since retrying those would just waste
 * time and hide a real problem from the caller.
 *
 * Honors `retryAfterMs` from a 429/503's Retry-After header as a floor for
 * the next wait, so we don't hammer Graph faster than it asked us to.
 */
export async function withRetries<T>(fn: (attempt: number) => Promise<T>, options: RetryOptions = {}): Promise<T> {
  const { maxAttempts, baseDelayMs, maxDelayMs } = { ...DEFAULTS, ...options };
  let lastError: unknown;

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      return await fn(attempt);
    } catch (err) {
      lastError = err;
      const retryable = err instanceof GraphApiError ? err.retryable : false;

      if (!retryable || attempt === maxAttempts) {
        throw err;
      }

      const exponential = Math.min(maxDelayMs, baseDelayMs * 2 ** (attempt - 1));
      const jittered = Math.random() * exponential;
      const retryAfter = err instanceof GraphApiError ? (err.retryAfterMs ?? 0) : 0;
      const delay = Math.max(jittered, retryAfter);

      console.warn(
        `[graph] retrying after ${err instanceof GraphApiError ? err.kind : "unknown"} error ` +
          `(attempt ${attempt}/${maxAttempts}, waiting ${Math.round(delay)}ms)`,
      );

      await sleep(delay);
    }
  }

  // Unreachable — the loop always returns or throws — but keeps TS happy.
  throw lastError;
}
