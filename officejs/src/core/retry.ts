// Transient-failure retry wrapper for a FetchLike. LLM endpoints rate-limit (429)
// and have transient 5xx/network blips; a spreadsheet recalc that fans out dozens
// of calls hits these constantly. Retrying idempotent chat/embedding POSTs with
// exponential backoff + jitter turns most of those into eventual successes instead
// of #ERROR cells. Pure and injectable (sleep + rng) so it is deterministically
// unit-testable with no timers or randomness.

import { FetchLike } from "./llm";

/* global setTimeout */

export interface RetryOptions {
  /** Additional attempts after the first (default 2 → up to 3 tries total). */
  retries?: number;
  baseDelayMs?: number;
  maxDelayMs?: number;
  /** Injectable for tests; defaults to a real timer. */
  sleep?: (ms: number) => Promise<void>;
  /** Injectable jitter source in [0,1); defaults to Math.random. */
  rng?: () => number;
  /** Which HTTP statuses are worth retrying (default 408/429/5xx). */
  isRetryable?: (status: number) => boolean;
}

const defaultSleep = (ms: number): Promise<void> => new Promise((r) => setTimeout(r, ms));
const defaultRetryable = (status: number): boolean => status === 408 || status === 429 || status >= 500;

export function withRetry(fetch: FetchLike, opts: RetryOptions = {}): FetchLike {
  const retries = opts.retries ?? 2;
  const base = opts.baseDelayMs ?? 500;
  const max = opts.maxDelayMs ?? 8000;
  const sleep = opts.sleep ?? defaultSleep;
  const rng = opts.rng ?? Math.random;
  const retryable = opts.isRetryable ?? defaultRetryable;

  return async (url, init) => {
    let lastErr: unknown;
    for (let attempt = 0; attempt <= retries; attempt++) {
      const isLast = attempt === retries;
      try {
        const resp = await fetch(url, init);
        // Retry transient HTTP failures; return anything else (incl. the final
        // transient response) so the caller surfaces the real error message.
        if (!isLast && !resp.ok && retryable(resp.status)) {
          await sleep(backoff(attempt, base, max, rng));
          continue;
        }
        return resp;
      } catch (e) {
        // Network-level failure (timeout, connection refused): retry then rethrow.
        lastErr = e;
        if (isLast) throw e;
        await sleep(backoff(attempt, base, max, rng));
      }
    }
    throw lastErr; // unreachable; satisfies the type checker
  };
}

/** Exponential backoff with full jitter, capped at maxDelayMs. */
function backoff(attempt: number, base: number, max: number, rng: () => number): number {
  const exp = Math.min(max, base * 2 ** attempt);
  return exp / 2 + rng() * (exp / 2);
}
