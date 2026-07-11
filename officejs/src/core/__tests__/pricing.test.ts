import { estimateCost, formatUsd } from "../pricing";

const u = (p: number, c: number) => ({ promptTokens: p, completionTokens: c, totalTokens: p + c });

describe("estimateCost", () => {
  test("prices a known model by input/output rates", () => {
    // gpt-4o-mini: 0.15 in / 0.60 out per 1M.
    const cost = estimateCost(u(1_000_000, 1_000_000), "gpt-4o-mini");
    expect(cost).toBeCloseTo(0.75);
  });

  test("matches the longest model substring (gpt-4o-mini over gpt-4o)", () => {
    const mini = estimateCost(u(1_000_000, 0), "openai/gpt-4o-mini");
    const full = estimateCost(u(1_000_000, 0), "openai/gpt-4o");
    expect(mini).toBeCloseTo(0.15);
    expect(full).toBeCloseTo(2.5);
  });

  test("is case-insensitive and tolerates provider prefixes", () => {
    expect(estimateCost(u(1_000_000, 0), "Meta-Llama/Llama-3.3-70B-Instruct")).toBeCloseTo(0.59);
  });

  test("returns null for an unpriced model", () => {
    expect(estimateCost(u(100, 100), "some-obscure-model")).toBeNull();
    expect(estimateCost(u(100, 100), "")).toBeNull();
  });
});

describe("formatUsd", () => {
  test("formats ranges with sensible precision", () => {
    expect(formatUsd(0)).toBe("$0.00");
    expect(formatUsd(0.004)).toBe("<$0.01");
    expect(formatUsd(0.123)).toBe("$0.123");
    expect(formatUsd(12.5)).toBe("$12.50");
  });
});
