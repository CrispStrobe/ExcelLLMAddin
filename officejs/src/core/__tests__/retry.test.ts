// Tests for the transient-failure retry wrapper. sleep + rng are injected so the
// backoff is instant and deterministic — no timers, no randomness.

import { withRetry } from "../retry";
import { FetchLike } from "../llm";

const instant = { sleep: async () => {}, rng: () => 0 };

/** A fetch double that yields queued outcomes: a number = HTTP status, an Error = throw. */
function scripted(outcomes: Array<number | Error>): { fetch: FetchLike; count: () => number } {
  let i = 0;
  const fetch: FetchLike = async () => {
    const o = outcomes[Math.min(i++, outcomes.length - 1)];
    if (o instanceof Error) throw o;
    return { ok: o >= 200 && o < 300, status: o, text: async () => `status ${o}` };
  };
  return { fetch, count: () => i };
}

describe("withRetry", () => {
  test("passes a first success straight through (one call)", async () => {
    const s = scripted([200]);
    const r = await withRetry(s.fetch, instant)("u", {});
    expect(r.status).toBe(200);
    expect(s.count()).toBe(1);
  });

  test("retries a 429 then succeeds", async () => {
    const s = scripted([429, 200]);
    const r = await withRetry(s.fetch, instant)("u", {});
    expect(r.ok).toBe(true);
    expect(s.count()).toBe(2);
  });

  test("retries a 503 then succeeds", async () => {
    const s = scripted([503, 200]);
    const r = await withRetry(s.fetch, instant)("u", {});
    expect(r.status).toBe(200);
    expect(s.count()).toBe(2);
  });

  test("gives up after `retries` and returns the last transient response", async () => {
    const s = scripted([429, 429, 429, 429]);
    const r = await withRetry(s.fetch, { ...instant, retries: 2 })("u", {});
    expect(r.status).toBe(429); // caller surfaces the real error
    expect(s.count()).toBe(3); // 1 + 2 retries
  });

  test("does NOT retry a non-retryable 400", async () => {
    const s = scripted([400, 200]);
    const r = await withRetry(s.fetch, instant)("u", {});
    expect(r.status).toBe(400);
    expect(s.count()).toBe(1);
  });

  test("retries a thrown network error then succeeds", async () => {
    const s = scripted([new Error("ECONNREFUSED"), 200]);
    const r = await withRetry(s.fetch, instant)("u", {});
    expect(r.status).toBe(200);
    expect(s.count()).toBe(2);
  });

  test("rethrows the network error after exhausting retries", async () => {
    const s = scripted([new Error("boom"), new Error("boom"), new Error("boom")]);
    await expect(withRetry(s.fetch, { ...instant, retries: 2 })("u", {})).rejects.toThrow("boom");
    expect(s.count()).toBe(3);
  });

  test("honors a custom isRetryable predicate", async () => {
    const s = scripted([418, 200]); // teapot: retryable only if we say so
    const r = await withRetry(s.fetch, { ...instant, isRetryable: (st) => st === 418 })("u", {});
    expect(r.status).toBe(200);
    expect(s.count()).toBe(2);
  });

  test("backoff delays grow and stay within the cap", async () => {
    const delays: number[] = [];
    const s = scripted([429, 429, 200]);
    await withRetry(s.fetch, {
      sleep: async (ms) => {
        delays.push(ms);
      },
      rng: () => 0, // no jitter → delay == exp/2
      retries: 2,
      baseDelayMs: 500,
      maxDelayMs: 8000,
    })("u", {});
    // attempt 0 → 500*2^0/2 = 250, attempt 1 → 500*2^1/2 = 500
    expect(delays).toEqual([250, 500]);
  });
});
