import {
  classify,
  extract,
  translate,
  summarize,
  mapRange,
  matchCategory,
  sentiment,
  listItems,
  extractFields,
  ask,
  similarity,
  cosine,
  tagText,
  editText,
  generateTable,
  fillByExample,
  writeFormula,
  explainFormula,
  analyzeImage,
  recall,
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

describe("similarity / cosine", () => {
  function embedMock(vectors: number[][]): { deps: Deps } {
    let i = 0;
    const fetch: FetchLike = async () => ({
      ok: true,
      status: 200,
      text: async () => JSON.stringify({ data: [{ embedding: vectors[Math.min(i++, vectors.length - 1)] }] }),
    });
    return { deps: { fetch } };
  }

  test("cosine: identical = 1, orthogonal = 0", () => {
    expect(cosine([1, 2, 3], [1, 2, 3])).toBeCloseTo(1);
    expect(cosine([1, 0], [0, 1])).toBeCloseTo(0);
  });

  test("similarity embeds both texts and returns cosine", async () => {
    const { deps } = embedMock([[1, 0, 0], [1, 0, 0]]);
    expect(await similarity("cat", "cat", "m", settings, deps)).toBeCloseTo(1);
  });

  test("similarity requires an embedding model", async () => {
    const { deps } = embedMock([[1]]);
    await expect(similarity("a", "b", "", settings, deps)).rejects.toThrow(/embedding model/i);
  });

  test("recall ranks candidates by cosine and returns top-k", async () => {
    // Embed order: query, then each candidate.
    const { deps } = embedMock([[1, 0, 0], [1, 0, 0], [0, 1, 0], [0.9, 0.1, 0]]);
    const out = await recall("q", ["c1", "c2", "c3"], 2, "m", settings, deps);
    expect(out).toHaveLength(2);
    expect(out[0][0]).toBe("c1");
    expect(out[0][1]).toBeCloseTo(1);
    expect(out[1][0]).toBe("c3"); // 0.99 beats c2's 0
  });

  test("recall drops blank candidates and returns [] on none", async () => {
    const { deps } = embedMock([[1]]);
    expect(await recall("q", ["", "   "], 3, "m", settings, deps)).toEqual([]);
  });

  test("recall propagates the missing-embedding-model error", async () => {
    const { deps } = embedMock([[1]]);
    await expect(recall("q", ["a"], 1, "", settings, deps)).rejects.toThrow(/embedding model/i);
  });

  test("recall tolerates a failed candidate embed (it sinks to the bottom)", async () => {
    // Query + good candidate succeed; the bad candidate's embed returns an error.
    let n = 0;
    const fetch: FetchLike = async () => {
      n++;
      if (n === 3) return { ok: false, status: 500, text: async () => `{"error":"boom"}` }; // 2nd candidate
      return {
        ok: true,
        status: 200,
        text: async () => JSON.stringify({ data: [{ embedding: [1, 0, 0] }] }),
      };
    };
    const out = await recall("q", ["good", "bad"], 2, "m", settings, { fetch });
    expect(out[0][0]).toBe("good");
    expect(out[1][0]).toBe("bad");
    expect(out[1][1]).toBeLessThan(0); // failed embed scored -1
  });
});

describe("extractFields / ask", () => {
  test("extractFields returns a value per field from a JSON array", async () => {
    const { deps } = mockFetch(`["Bob","bob@x.com","30"]`);
    expect(await extractFields("Bob bob@x.com 30", ["name", "email", "age"], settings, deps)).toEqual([
      "Bob",
      "bob@x.com",
      "30",
    ]);
  });

  test("extractFields falls back to per-field when the array doesn't match", async () => {
    const { deps, calls } = mockFetch("X");
    const out = await extractFields("t", ["a", "b"], settings, deps);
    expect(out).toEqual(["X", "X"]);
    expect(calls.length).toBe(3); // 1 batch attempt + 2 per-field
  });

  test("ask includes context + question and trims the answer", async () => {
    const { deps, calls } = mockFetch("  42 \n");
    expect(await ask("how many apples?", "there are 42 apples", settings, deps)).toBe("42");
    const body = JSON.parse(calls[0].init.body).messages[1].content;
    expect(body).toContain("there are 42 apples");
    expect(body).toContain("how many apples?");
  });
});

describe("tagText", () => {
  test("returns only the recognized labels, order-preserving", async () => {
    const { deps } = mockFetch("This is about Billing and Feature requests");
    const out = await tagText("...", ["Bug", "Billing", "Feature"], settings, deps);
    expect(out).toBe("Billing, Feature");
  });

  test("returns empty when the model matches nothing valid", async () => {
    const { deps } = mockFetch("none of the above");
    expect(await tagText("x", ["Bug", "Billing"], settings, deps)).toBe("");
  });

  test("matches whole words only (debugging does not match Bug)", async () => {
    const { deps } = mockFetch("This is about debugging the feature");
    expect(await tagText("x", ["Bug", "Feature"], settings, deps)).toBe("Feature");
  });

  test("errors with no categories", async () => {
    const { deps } = mockFetch("x");
    expect(await tagText("x", [], settings, deps)).toMatch(/no categories/);
  });
});

describe("editText", () => {
  test("returns the revised text trimmed", async () => {
    const { deps, calls } = mockFetch("  Their house is nice.  ");
    expect(await editText("they're house is nice", undefined, settings, deps)).toBe("Their house is nice.");
    expect(JSON.parse(calls[0].init.body).messages[1].content).toMatch(/Fix spelling and grammar/);
  });

  test("passes a custom instruction into the system prompt", async () => {
    const { deps, calls } = mockFetch("HELLO");
    await editText("hello", "make it uppercase", settings, deps);
    expect(JSON.parse(calls[0].init.body).messages[1].content).toMatch(/make it uppercase/);
  });
});

describe("generateTable", () => {
  test("parses a JSON array-of-arrays into a grid", async () => {
    const { deps } = mockFetch(`[["Country","Pop"],["France","68"],["Spain","48"]]`);
    const grid = await generateTable("EU countries", settings, deps);
    expect(grid).toEqual([["Country", "Pop"], ["France", "68"], ["Spain", "48"]]);
  });

  test("treats a flat array as a single column", async () => {
    const { deps } = mockFetch(`["a","b","c"]`);
    expect(await generateTable("x", settings, deps)).toEqual([["a"], ["b"], ["c"]]);
  });

  test("reports a parse error instead of throwing", async () => {
    const { deps } = mockFetch("sorry, no table here");
    expect((await generateTable("x", settings, deps))[0][0]).toMatch(/could not parse/);
  });
});

describe("fillByExample", () => {
  test("batches examples + inputs and returns aligned outputs", async () => {
    const { deps, calls } = mockFetch(`["FR","ES"]`);
    const out = await fillByExample(
      [{ input: "Germany", output: "DE" }],
      ["France", "Spain"],
      settings,
      deps
    );
    expect(out).toEqual(["FR", "ES"]);
    const body = JSON.parse(calls[0].init.body).messages[1].content;
    expect(body).toMatch(/Germany.*=>.*DE/);
  });

  test("errors on every input when no example pairs are given", async () => {
    const { deps, calls } = mockFetch(`["x"]`);
    const out = await fillByExample([], ["a", "b"], settings, deps);
    expect(out).toEqual(["Error: need at least one example (input, output) pair", "Error: need at least one example (input, output) pair"]);
    expect(calls.length).toBe(0);
  });

  test("falls back to per-input when the batch reply is malformed", async () => {
    const { deps, calls } = mockFetch("R"); // not a JSON array
    const out = await fillByExample([{ input: "a", output: "A" }], ["x", "y"], settings, deps);
    expect(out).toEqual(["R", "R"]);
    expect(calls.length).toBe(3); // 1 batch + 2 per-input
  });
});

describe("writeFormula", () => {
  test("returns the formula as-is when clean", async () => {
    const { deps } = mockFetch('=SUMIF(A:A,">100",B:B)');
    expect(await writeFormula("sum B where A>100", settings, deps)).toBe('=SUMIF(A:A,">100",B:B)');
  });

  test("strips code fences and surrounding prose", async () => {
    const { deps } = mockFetch("Here you go:\n```excel\n=XLOOKUP(A2,D:D,E:E)\n```");
    expect(await writeFormula("lookup", settings, deps)).toBe("=XLOOKUP(A2,D:D,E:E)");
  });

  test("prepends = when the model omits it", async () => {
    const { deps } = mockFetch("SUM(B2:B10)");
    expect(await writeFormula("sum", settings, deps)).toBe("=SUM(B2:B10)");
  });

  test("does not cut at an inner comparison = when the leading = is missing", async () => {
    const { deps } = mockFetch("IF(A1=B1,1,0)");
    expect(await writeFormula("compare", settings, deps)).toBe("=IF(A1=B1,1,0)");
  });

  test("strips same-line prose before the formula", async () => {
    const { deps } = mockFetch("The formula is =XLOOKUP(A2,D:D,E:E)");
    expect(await writeFormula("lookup", settings, deps)).toBe("=XLOOKUP(A2,D:D,E:E)");
  });
});

describe("explainFormula", () => {
  test("returns the explanation trimmed", async () => {
    const { deps, calls } = mockFetch("  It sums B2:B10.  ");
    expect(await explainFormula("=SUM(B2:B10)", settings, deps)).toBe("It sums B2:B10.");
    expect(JSON.parse(calls[0].init.body).messages[1].content).toContain("=SUM(B2:B10)");
  });
});

describe("analyzeImage (VISION)", () => {
  const visionSettings: LlmSettings = { provider: "openai", model: "gpt-4o", apiKey: "sk-test" };

  test("builds a multimodal content array and returns the answer", async () => {
    const { deps, calls } = mockFetch("A cat.");
    const out = await analyzeImage("https://x/cat.png", "what is this?", visionSettings, deps);
    expect(out).toBe("A cat.");
    const msg = JSON.parse(calls[0].init.body).messages[1];
    expect(msg.content[0]).toEqual({ type: "text", text: "what is this?" });
    expect(msg.content[1]).toEqual({ type: "image_url", image_url: { url: "https://x/cat.png" } });
  });

  test("defaults the question when none is given", async () => {
    const { deps, calls } = mockFetch("ok");
    await analyzeImage("data:image/png;base64,AAAA", "", visionSettings, deps);
    expect(JSON.parse(calls[0].init.body).messages[1].content[0].text).toMatch(/Describe this image/);
  });

  test("errors on an empty image without fetching", async () => {
    const { deps, calls } = mockFetch(`{}`);
    expect(await analyzeImage("", "q", visionSettings, deps)).toMatch(/no image/i);
    expect(calls.length).toBe(0);
  });

  test("rejects the Ollama style (different image format)", async () => {
    const { deps } = mockFetch(`{}`);
    await expect(analyzeImage("https://x.png", "q", { provider: "ollama", model: "llava" }, deps)).rejects.toThrow(
      /Ollama/
    );
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

  test("tolerates a trailing comma in the batch reply (no fallback)", async () => {
    const { deps, calls } = mockFetch(`["A","B",]`);
    const out = await mapRange([["a", "b"]], "up", settings, deps);
    expect(out).toEqual([["A", "B"]]);
    expect(calls.length).toBe(1); // repaired in-place, not per-cell
  });

  test("tolerates single-quoted strings in the batch reply", async () => {
    const { deps, calls } = mockFetch(`['A', 'B']`);
    const out = await mapRange([["a", "b"]], "up", settings, deps);
    expect(out).toEqual([["A", "B"]]);
    expect(calls.length).toBe(1);
  });

  test("reports a per-cell error instead of throwing when the call fails", async () => {
    const fetch: FetchLike = async () => ({
      ok: false,
      status: 500,
      text: async () => JSON.stringify({ error: { message: "upstream down" } }),
    });
    // A single non-empty cell takes the single-input chunk path (its own try/catch).
    const out = await mapRange([["only"]], "up", settings, { fetch });
    expect(out[0][0]).toMatch(/Error: upstream down/);
  });
});
