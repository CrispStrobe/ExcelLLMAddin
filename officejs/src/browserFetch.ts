// Browser fetch adapter with a timeout, shared by the task pane and the custom-
// functions runtime. Without a timeout, a request to an unreachable provider
// (e.g. Ollama not running on localhost) hangs the UI for the OS default (~1 min)
// instead of failing fast.

import { FetchLike } from "./core/llm";
import { withRetry } from "./core/retry";

/* global fetch, AbortController, setTimeout, clearTimeout, RequestInit */

const TIMEOUT_MS = 45000;

export const browserFetch: FetchLike = async (url, init) => {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), TIMEOUT_MS);
  try {
    const r = await fetch(url, { ...(init as RequestInit), signal: controller.signal });
    return { ok: r.ok, status: r.status, text: () => r.text() };
  } catch (e) {
    if (e instanceof Error && e.name === "AbortError") {
      throw new Error(`Request timed out after ${TIMEOUT_MS / 1000}s (is the provider reachable?)`);
    }
    throw e;
  } finally {
    clearTimeout(timer);
  }
};

// What callers should use: the timeout-bounded fetch above, wrapped so transient
// rate-limits (429) and 5xx/network blips retry with backoff instead of failing a
// cell. A per-request timeout still bounds each individual attempt.
export const resilientFetch: FetchLike = withRetry(browserFetch);
