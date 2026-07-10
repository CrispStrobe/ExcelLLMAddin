import { classify, extract, translate, summarize, mapRange, matchCategory } from "../tasks";
import { LlmSettings, Deps, FetchLike } from "../llm";

/** Mock fetch that wraps `reply` as an Ollama chat response and records calls. */
function mockFetch(reply: string): { deps: Deps; calls: Array<{ url: string; init: any }> } {
  const calls: Array<{ url: string; init: any }> = [];
  const fetch: FetchLike = async (url, init) => {
    calls.push({ url, init });
    return { ok: true, status: 200, text: async () => JSON.stringify({ message: { content: reply } }) };
  };
  return { deps: { fetch }, calls };
}

const settings: LlmSettings = { provider: "ollama", model: "m" };

describe("classify", () => {
  test("returns the matched category and passes labels in the system prompt", async () => {
    const { deps, calls } = mockFetch("Positive");
    expect(await classify("great product", ["Positive", "Negative"], settings, deps)).toBe("Positive");
    const body = JSON.parse(calls[0].init.body);
    expect(body.messages[0].content).toContain("Positive");
    expect(body.messages[0].content).toContain("Negative");
  });

  test("normalizes casing/extra words to the provided label", async () => {
    const { deps } = mockFetch("The answer is positive.");
    expect(await classify("x", ["Positive", "Negative"], settings, deps)).toBe("Positive");
  });

  test("guards empty category list", async () => {
    const { deps } = mockFetch("x");
    expect(await classify("x", [], settings, deps)).toMatch(/no categories/);
  });
});

describe("matchCategory (unit)", () => {
  test("exact, contained, fallback", () => {
    expect(matchCategory("neg", ["pos", "neg"])).toBe("neg");
    expect(matchCategory("I'd say POS here", ["pos", "neg"])).toBe("pos");
    expect(matchCategory("maybe", ["pos", "neg"])).toBe("maybe");
  });
});

describe("extract / translate / summarize", () => {
  test("extract trims the reply", async () => {
    const { deps } = mockFetch("  bob@example.com \n");
    expect(await extract("mail bob@example.com", "the email", settings, deps)).toBe("bob@example.com");
  });

  test("translate sends the target language", async () => {
    const { deps, calls } = mockFetch("Hallo");
    expect(await translate("Hello", "German", settings, deps)).toBe("Hallo");
    expect(JSON.parse(calls[0].init.body).messages[1].content).toContain("German");
  });

  test("summarize includes a word cap when given", async () => {
    const { deps, calls } = mockFetch("short");
    await summarize("long text", 10, settings, deps);
    expect(JSON.parse(calls[0].init.body).messages[1].content).toContain("10 words");
  });
});

describe("mapRange", () => {
  test("preserves shape, skips empty cells, fills the rest", async () => {
    const { deps, calls } = mockFetch("X");
    const out = await mapRange([["a", "b"], ["", "c"]], "uppercase", settings, deps, 2);
    expect(out).toEqual([["X", "X"], ["", "X"]]);
    expect(calls.length).toBe(3); // the empty cell made no call
  });

  test("coerces non-string cells and passes the instruction", async () => {
    const { deps, calls } = mockFetch("R");
    const out = await mapRange([[42, true]], "describe", settings, deps, 4);
    expect(out).toEqual([["R", "R"]]);
    expect(JSON.parse(calls[0].init.body).messages[1].content).toContain("describe");
  });
});
