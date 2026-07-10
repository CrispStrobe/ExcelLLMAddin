import {
  classify,
  extract,
  translate,
  summarize,
  mapRange,
  matchCategory,
  sentiment,
  listItems,
} from "../tasks";
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

describe("sentiment / listItems", () => {
  test("sentiment returns one of the three labels", async () => {
    const { deps } = mockFetch("Negative");
    expect(await sentiment("this is terrible", settings, deps)).toBe("Negative");
  });

  test("listItems parses a JSON array", async () => {
    const { deps } = mockFetch(`["red","green","blue"]`);
    expect(await listItems("primary colors", undefined, settings, deps)).toEqual(["red", "green", "blue"]);
  });

  test("listItems respects count", async () => {
    const { deps } = mockFetch(`["a","b","c","d"]`);
    expect(await listItems("letters", 2, settings, deps)).toEqual(["a", "b"]);
  });

  test("listItems falls back to line-splitting with bullets/numbers", async () => {
    const { deps } = mockFetch("1. alpha\n2. beta\n- gamma");
    expect(await listItems("greek", undefined, settings, deps)).toEqual(["alpha", "beta", "gamma"]);
  });
});

describe("mapRange", () => {
  test("batches multiple cells into a single call", async () => {
    const { deps, calls } = mockFetch(`["A","B","C"]`);
    const out = await mapRange([["a", "b", "c"]], "uppercase", settings, deps);
    expect(out).toEqual([["A", "B", "C"]]);
    expect(calls.length).toBe(1); // one batched call, not three
  });

  test("preserves shape and skips empty cells", async () => {
    const { deps } = mockFetch(`["A","C"]`);
    const out = await mapRange([["a", ""], ["", "c"]], "up", settings, deps);
    expect(out).toEqual([["A", ""], ["", "C"]]);
  });

  test("falls back to per-cell when the batch reply isn't a JSON array", async () => {
    const { deps, calls } = mockFetch("X");
    const out = await mapRange([["a", "b"]], "up", settings, deps);
    expect(out).toEqual([["X", "X"]]);
    expect(calls.length).toBe(3); // 1 batch attempt + 2 per-cell fallback
  });

  test("batchSize 1 does per-cell calls and coerces non-strings", async () => {
    const { deps, calls } = mockFetch("R");
    const out = await mapRange([[42, true]], "describe", settings, deps, { batchSize: 1 });
    expect(out).toEqual([["R", "R"]]);
    expect(calls.length).toBe(2);
    expect(JSON.parse(calls[0].init.body).messages[1].content).toContain("describe");
  });
});
