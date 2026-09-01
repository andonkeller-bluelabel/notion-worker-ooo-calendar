/**
 * Normalized error shape for anything that can go wrong calling Microsoft
 * Graph. Nothing in this worker should throw a bare Error or leak a raw fetch
 * Response — every failure gets wrapped here so callers can decide: retry,
 * self-heal, drop-and-log, or surface. Modeled on the Kantata worker's
 * errors.ts.
 */

export type GraphErrorKind =
  | "auth" // 401 - missing/expired/invalid token, or the client secret rotated
  | "forbidden" // 403 - token valid, but the Application Access Policy denies this mailbox
  | "not_found" // 404 - group/event id unknown (or the event was already deleted)
  | "validation" // 400 - Graph rejected the payload
  | "rate_limited" // 429
  | "server" // 5xx from Graph
  | "network" // fetch threw: DNS, timeout, connection reset, etc.
  | "unexpected"; // anything that doesn't fit the above

export class GraphApiError extends Error {
  readonly kind: GraphErrorKind;
  readonly status: number | null;
  readonly retryable: boolean;
  /** Graph's own machine-readable code, e.g. "ErrorAccessDenied", "ErrorItemNotFound". */
  readonly code: string | null;
  /** Present on 429/503 responses when Graph sends a Retry-After header. */
  readonly retryAfterMs: number | null;
  readonly endpoint: string;
  readonly method: string;

  constructor(params: {
    kind: GraphErrorKind;
    message: string;
    status?: number | null;
    code?: string | null;
    retryAfterMs?: number | null;
    endpoint: string;
    method: string;
  }) {
    super(params.message);
    this.name = "GraphApiError";
    this.kind = params.kind;
    this.status = params.status ?? null;
    this.code = params.code ?? null;
    this.retryAfterMs = params.retryAfterMs ?? null;
    this.endpoint = params.endpoint;
    this.method = params.method;
    this.retryable = params.kind === "rate_limited" || params.kind === "server" || params.kind === "network";
  }

  /** Compact, JSON-safe summary to return from a tool instead of throwing. */
  toResult() {
    return {
      ok: false as const,
      error: {
        kind: this.kind,
        message: this.message,
        status: this.status,
        code: this.code,
        retryable: this.retryable,
        endpoint: this.endpoint,
        method: this.method,
      },
    };
  }
}

/** Thrown for problems caught before ever calling the network. */
export class InputValidationError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "InputValidationError";
  }

  toResult() {
    return {
      ok: false as const,
      error: {
        kind: "validation" as const,
        message: this.message,
        status: null,
        code: null,
        retryable: false,
        endpoint: null,
        method: null,
      },
    };
  }
}

export function isGraphApiError(err: unknown): err is GraphApiError {
  return err instanceof GraphApiError;
}

/** True when Graph says the thing we addressed no longer exists. */
export function isNotFound(err: unknown): boolean {
  return err instanceof GraphApiError && err.kind === "not_found";
}

export function isKnownError(err: unknown): err is GraphApiError | InputValidationError {
  return err instanceof GraphApiError || err instanceof InputValidationError;
}
