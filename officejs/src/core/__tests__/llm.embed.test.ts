// Coverage for embed() and the remaining llm.ts branches not hit by llm.test.ts:
// embedding transport (direct openai/ollama + proxy), its guards and cache, plus
// listModels error paths, chat/model extraction fallbacks, and header/error edges.

import {
  embed,
  embedBatch,
  listModels,
  directHeaders,
  extractChatContent,
  extractModelList,
  runPrompt,
  LlmError,
  LlmSettings,
  Deps,
  FetchLike,
} from "../llm";
import { PROVIDERS } from "../providers";
import { createLruCache } from "../cache";

function mockFetch(
  body: string,
  opts: { ok?: boolean; status?: number } = {}
): { deps: Deps; calls: Array<{ url: string; init: any }> } {
  const calls: Array<{ url: string; init: any }> = [];
  const fetch: FetchLike = async (url, init) => {
    calls.push({ url, init });
    return { ok: opts.ok ?? true, status: opts.status ?? 200, text: async () => body };
  };
  return { deps: { fetch }, calls };
}

const ollama: LlmSettings = { provider: "ollama", model: "test-model" };
const openai: LlmSettings = { provider: "openai", model: "gpt-4o-mini", apiKey: "sk-test" };

describe("embed", () => {
  test("openai direct returns the vector and sends {input}", async () => {
    const { deps, calls } = mockFetch(`{"data":[{"embedding":[1,2,3]}]}`);
    expect(await embed("hi", "text-embedding-3-small", openai, deps)).toEqual([1, 2, 3]);
    const sent = JSON.parse(calls[0].init.body);
    expect(sent).toMatchObject({ model: "text-embedding-3-small", input: "hi" });
    expect(calls[0].url).toContain("/embeddings");
  });

  test("ollama direct reads {embedding} and sends {prompt}", async () => {
    const { deps, calls } = mockFetch(`{"embedding":[4,5]}`);
    expect(await embed("hi", "nomic", ollama, deps)).toEqual([4, 5]);
    expect(JSON.parse(calls[0].init.body)).toMatchObject({ model: "nomic", prompt: "hi" });
  });

  test("reads Ollama batch shape {embeddings:[[...]]}", async () => {
    const { deps } = mockFetch(`{"embeddings":[[6,7]]}`);
    expect(await embed("hi", "nomic", ollama, deps)).toEqual([6, 7]);
  });

  test("throws when no embedding model is provided", async () => {
    const { deps, calls } = mockFetch(`{}`);
    await expect(embed("hi", "", openai, deps)).rejects.toThrow(/embedding model/i);
    expect(calls.length).toBe(0);
  });

  test("throws on missing api key before fetching", async () => {
    const { deps, calls } = mockFetch(`{}`);
    await expect(embed("hi", "m", { provider: "openai", model: "x" }, deps)).rejects.toThrow(/No API key/);
    expect(calls.length).toBe(0);
  });

  test("surfaces a provider error", async () => {
    const { deps } = mockFetch(`{"error":{"message":"model gone"}}`);
    await expect(embed("hi", "m", openai, deps)).rejects.toThrow("model gone");
  });

  test("throws when the response has no embedding", async () => {
    const { deps } = mockFetch(`{"foo":1}`);
    await expect(embed("hi", "m", openai, deps)).rejects.toThrow(/No embedding found/);
  });

  test("identical text hits the cache (one fetch)", async () => {
    const { deps, calls } = mockFetch(`{"embedding":[1,1]}`);
    const cached = { ...deps, cache: createLruCache() };
    expect(await embed("t", "m", ollama, cached)).toEqual([1, 1]);
    expect(await embed("t", "m", ollama, cached)).toEqual([1, 1]);
    expect(calls.length).toBe(1);
  });

  describe("via proxy", () => {
    const proxied: LlmSettings = { provider: "openai", model: "m", proxyUrl: "https://proxy.example/api" };

    test("returns the proxy embedding, no key needed", async () => {
      const { deps, calls } = mockFetch(`{"embedding":[7,8]}`);
      expect(await embed("hi", "m", proxied, deps)).toEqual([7, 8]);
      expect(JSON.parse(calls[0].init.body)).toMatchObject({ op: "embed", provider: "openai" });
    });

    test("throws when the proxy returns no embedding", async () => {
      const { deps } = mockFetch(`{"content":"oops"}`);
      await expect(embed("hi", "m", proxied, deps)).rejects.toThrow(/no embedding/i);
    });
  });
});

describe("embedBatch", () => {
  test("openai sends one {input:[...]} request and returns vectors in order", async () => {
    const { deps, calls } = mockFetch(
      `{"data":[{"embedding":[3,3],"index":1},{"embedding":[1,1],"index":0}]}`
    );
    const vecs = await embedBatch(["a", "b"], "m", openai, deps);
    // Rows arrive out of order but carry index → re-sorted to input order.
    expect(vecs).toEqual([[1, 1], [3, 3]]);
    expect(calls.length).toBe(1);
    expect(JSON.parse(calls[0].init.body)).toMatchObject({ model: "m", input: ["a", "b"] });
  });

  test("returns [] for no texts without fetching", async () => {
    const { deps, calls } = mockFetch(`{}`);
    expect(await embedBatch([], "m", openai, deps)).toEqual([]);
    expect(calls.length).toBe(0);
  });

  test("throws when the provider returns the wrong number of rows", async () => {
    const { deps } = mockFetch(`{"data":[{"embedding":[1]}]}`); // asked for 2
    await expect(embedBatch(["a", "b"], "m", openai, deps)).rejects.toThrow(/mismatch/i);
  });

  test("ollama falls back to one request per text", async () => {
    const { deps, calls } = mockFetch(`{"embedding":[9]}`);
    const vecs = await embedBatch(["a", "b", "c"], "nomic", ollama, deps);
    expect(vecs).toEqual([[9], [9], [9]]);
    expect(calls.length).toBe(3);
  });

  test("serves cached texts and only requests the misses", async () => {
    const { deps, calls } = mockFetch(`{"data":[{"embedding":[2,2]}]}`);
    const cached = { ...deps, cache: createLruCache() };
    await embedBatch(["x"], "m", openai, cached); // warms cache for "x" (1 request)
    const vecs = await embedBatch(["x", "y"], "m", openai, cached); // only "y" is fetched
    expect(vecs).toEqual([[2, 2], [2, 2]]);
    const bodies = calls.map((c) => JSON.parse(c.init.body).input);
    expect(bodies).toEqual([["x"], ["y"]]); // second call omits the cache hit
  });

  test("proxy sends {op:'embed', inputs} and reads {embeddings}", async () => {
    const { deps, calls } = mockFetch(`{"embeddings":[[1],[2]]}`);
    const proxied: LlmSettings = { provider: "openai", model: "m", proxyUrl: "https://proxy.example/api" };
    expect(await embedBatch(["a", "b"], "m", proxied, deps)).toEqual([[1], [2]]);
    expect(JSON.parse(calls[0].init.body)).toMatchObject({ op: "embed", inputs: ["a", "b"] });
  });
});

describe("listModels error paths", () => {
  test("throws on a missing api key", async () => {
    const { deps, calls } = mockFetch(`{}`);
    await expect(listModels({ provider: "openai", model: "x" }, deps)).rejects.toThrow(/No API key/);
    expect(calls.length).toBe(0);
  });

  test("surfaces an HTTP error body", async () => {
    const { deps } = mockFetch(`{"error":{"message":"forbidden"}}`, { ok: false, status: 403 });
    await expect(listModels(openai, deps)).rejects.toThrow("forbidden");
  });
});

describe("extraction fallbacks", () => {
  test("extractChatContent falls back to choice.text", () => {
    expect(extractChatContent(`{"choices":[{"text":"legacy completion"}]}`)).toBe("legacy completion");
  });

  test("extractChatContent throws when there is no content", () => {
    expect(() => extractChatContent(`{"unexpected":1}`)).toThrow(/No content/);
  });

  test("extractModelList throws on an error body", () => {
    expect(() => extractModelList(PROVIDERS.openai, `{"error":{"message":"nope"}}`)).toThrow("nope");
  });

  test("extractModelList returns [] on an unrecognized shape", () => {
    expect(extractModelList(PROVIDERS.openai, `{}`)).toEqual([]);
    expect(extractModelList(PROVIDERS.ollama, `{}`)).toEqual([]);
  });

  test("extractModelList accepts a bare array of models (Together AI shape)", () => {
    // Together returns [{id, ...}] instead of {data:[{id}]} — live-observed.
    expect(extractModelList(PROVIDERS.together, `[{"id":"a"},{"id":"b"}]`)).toEqual(["a", "b"]);
  });

  test("extractModelList drops entries with no id/name", () => {
    expect(extractModelList(PROVIDERS.openai, `{"data":[{"id":"a"},{"object":"model"}]}`)).toEqual(["a"]);
  });
});

describe("directHeaders", () => {
  test("adds OpenRouter attribution headers", () => {
    const h = directHeaders(PROVIDERS.openrouter, "key");
    expect(h["Authorization"]).toBe("Bearer key");
    expect(h["HTTP-Referer"]).toBeDefined();
    expect(h["X-Title"]).toBeDefined();
  });

  test("omits Authorization when there is no key", () => {
    expect(directHeaders(PROVIDERS.ollama)["Authorization"]).toBeUndefined();
  });
});

describe("gateway error metadata", () => {
  test("appends provider_name when no raw detail is present", async () => {
    const { deps } = mockFetch(
      `{"error":{"message":"Provider returned error","metadata":{"provider_name":"Together"}}}`
    );
    await expect(runPrompt("x", openai, deps)).rejects.toThrow(/via Together/);
  });
});
