// Token-usage accounting. Providers report token counts in their responses
// (OpenAI: usage.{prompt,completion,total}_tokens; Ollama: prompt_eval_count +
// eval_count). parseUsage normalizes them; a shared UsageTracker accumulates a
// running session total that the task pane displays. Pure + unit-tested; the
// tracker is a module singleton so the (shared-runtime) task pane and custom
// functions report into the same counter.

import { ProviderStyle } from "./providers";

export interface TokenUsage {
  promptTokens: number;
  completionTokens: number;
  totalTokens: number;
}

export interface UsageTotals extends TokenUsage {
  /** Number of calls that reported usage. */
  calls: number;
}

function num(v: unknown): number {
  return typeof v === "number" && isFinite(v) ? v : 0;
}

/** Extract token usage from a raw chat response body; null if none is present. */
export function parseUsage(text: string, style: ProviderStyle): TokenUsage | null {
  let d: any;
  try {
    d = JSON.parse(text);
  } catch {
    return null;
  }
  if (!d || typeof d !== "object") return null;

  if (style === "ollama") {
    const p = num(d.prompt_eval_count);
    const c = num(d.eval_count);
    if (p === 0 && c === 0) return null;
    return { promptTokens: p, completionTokens: c, totalTokens: p + c };
  }

  const u = d.usage;
  if (!u || typeof u !== "object") return null;
  const p = num(u.prompt_tokens);
  const c = num(u.completion_tokens);
  const t = num(u.total_tokens) || p + c;
  if (p === 0 && c === 0 && t === 0) return null;
  return { promptTokens: p, completionTokens: c, totalTokens: t };
}

class UsageTracker {
  private t: UsageTotals = { calls: 0, promptTokens: 0, completionTokens: 0, totalTokens: 0 };

  add(u: TokenUsage | null): void {
    if (!u) return;
    this.t.calls += 1;
    this.t.promptTokens += u.promptTokens;
    this.t.completionTokens += u.completionTokens;
    this.t.totalTokens += u.totalTokens;
  }

  total(): UsageTotals {
    return { ...this.t };
  }

  reset(): void {
    this.t = { calls: 0, promptTokens: 0, completionTokens: 0, totalTokens: 0 };
  }
}

/** Shared session-wide token counter. */
export const usageTracker = new UsageTracker();
