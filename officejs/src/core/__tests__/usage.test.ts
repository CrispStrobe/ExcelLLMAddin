import { parseUsage, usageTracker } from "../usage";

describe("parseUsage", () => {
  test("reads OpenAI usage", () => {
    const u = parseUsage(`{"usage":{"prompt_tokens":10,"completion_tokens":5,"total_tokens":15}}`, "openai");
    expect(u).toEqual({ promptTokens: 10, completionTokens: 5, totalTokens: 15 });
  });

  test("derives total when OpenAI omits it", () => {
    const u = parseUsage(`{"usage":{"prompt_tokens":7,"completion_tokens":3}}`, "openai");
    expect(u?.totalTokens).toBe(10);
  });

  test("reads Ollama eval counts", () => {
    const u = parseUsage(`{"message":{"content":"hi"},"prompt_eval_count":8,"eval_count":4}`, "ollama");
    expect(u).toEqual({ promptTokens: 8, completionTokens: 4, totalTokens: 12 });
  });

  test("returns null when no usage is present", () => {
    expect(parseUsage(`{"choices":[{"message":{"content":"hi"}}]}`, "openai")).toBeNull();
    expect(parseUsage(`{"message":{"content":"hi"}}`, "ollama")).toBeNull();
    expect(parseUsage("not json", "openai")).toBeNull();
  });
});

describe("usageTracker", () => {
  beforeEach(() => usageTracker.reset());

  test("accumulates across calls and counts them", () => {
    usageTracker.add({ promptTokens: 10, completionTokens: 5, totalTokens: 15 });
    usageTracker.add({ promptTokens: 2, completionTokens: 3, totalTokens: 5 });
    expect(usageTracker.total()).toEqual({ calls: 2, promptTokens: 12, completionTokens: 8, totalTokens: 20 });
  });

  test("ignores null and reset clears", () => {
    usageTracker.add(null);
    expect(usageTracker.total().calls).toBe(0);
    usageTracker.add({ promptTokens: 1, completionTokens: 1, totalTokens: 2 });
    usageTracker.reset();
    expect(usageTracker.total()).toEqual({ calls: 0, promptTokens: 0, completionTokens: 0, totalTokens: 0 });
  });
});
