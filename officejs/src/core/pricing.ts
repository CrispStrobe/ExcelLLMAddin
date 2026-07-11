// Rough USD list-prices per 1M tokens [input, output] for common models, matched
// by substring (longest match wins). Open models are priced differently by each
// host, so this is a best-effort estimate; unknown models return null and the
// meter simply shows no dollar figure. Pure + unit-tested.

import { TokenUsage } from "./usage";

// [modelSubstring, inputPerMillion, outputPerMillion]
const PRICES: Array<[string, number, number]> = [
  ["gpt-4o-mini", 0.15, 0.6],
  ["gpt-4o", 2.5, 10],
  ["gpt-4.1-mini", 0.4, 1.6],
  ["gpt-4.1-nano", 0.1, 0.4],
  ["gpt-4.1", 2.0, 8.0],
  ["o3-mini", 1.1, 4.4],
  ["o4-mini", 1.1, 4.4],
  ["claude-3-5-haiku", 0.8, 4.0],
  ["claude-3-5-sonnet", 3.0, 15],
  ["claude-3-7-sonnet", 3.0, 15],
  ["claude-haiku", 0.8, 4.0],
  ["claude-sonnet", 3.0, 15],
  ["gemini-2.0-flash", 0.1, 0.4],
  ["gemini-1.5-flash", 0.075, 0.3],
  ["gemini-1.5-pro", 1.25, 5.0],
  ["mistral-small", 0.2, 0.6],
  ["mistral-large", 2.0, 6.0],
  ["llama-3.1-8b", 0.05, 0.08],
  ["llama-3.1-70b", 0.59, 0.79],
  ["llama-3.3-70b", 0.59, 0.79],
];

/** Estimate the USD cost of one usage record for a model, or null if unpriced. */
export function estimateCost(usage: TokenUsage, model: string): number | null {
  const m = (model || "").toLowerCase();
  let best: [string, number, number] | null = null;
  for (const row of PRICES) {
    if (m.includes(row[0]) && (!best || row[0].length > best[0].length)) best = row;
  }
  if (!best) return null;
  return (usage.promptTokens * best[1] + usage.completionTokens * best[2]) / 1_000_000;
}

/** Format a USD amount for the meter (more precision for tiny amounts). */
export function formatUsd(v: number): string {
  if (v === 0) return "$0.00";
  if (v < 0.01) return "<$0.01";
  return "$" + v.toFixed(v < 1 ? 3 : 2);
}
