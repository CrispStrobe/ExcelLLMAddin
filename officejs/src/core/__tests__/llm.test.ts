import {
  runPrompt,
  listModels,
  buildChatBody,
  extractChatContent,
  LlmError,
  LlmSettings,
  Deps,
  FetchLike,
} from "../llm";
import { PROVIDERS } from "../providers";

/** A fetch double that returns a canned body and records requests. */
function mockFetch(
  body: string,
  opts: { ok?: boolean; status?: number } = {}
): { deps: Deps; calls: Array<{ url: string; init: any }> } {
  const calls: Array<{ url: string; init: any }> = [];
  const fetch: FetchLike = async (url, init) => {
    calls.push({ url, init });
    return {
      ok: opts.ok ?? true,
      status: opts.status ?? 200,
      text: async () => body,
    };
  };
  return { deps: { fetch }, calls };
}

const ollama: LlmSettings = { provider: "ollama", model: "test-model" };
const openai: LlmSettings = { provider: "openai", model: "gpt-4o-mini", apiKey: "sk-test" };

describe("runPrompt", () => {
  test("ollama basic", async () => {
    const { deps } = mockFetch(`{"message":{"content":"Hello Excel"},"done":true}`);
    expect(await runPrompt("hi", ollama, deps)).toBe("Hello Excel");
  });

  test("openai basic", async () => {
    const { deps } = mockFetch(`{"choices":[{"message":{"role":"assistant","content":"Hi there"}}]}`);
    expect(await runPrompt("hi", openai, deps)).toBe("Hi there");
  });

  test("unicode content survives (native JSON)", async () => {
    const { deps } = mockFetch(`{"message":{"content":"Café naïve 😀"}}`);
    expect(await runPrompt("x", ollama, deps)).toBe("Café naïve 😀");
  });

  test("embedded quotes and newlines", async () => {
    const { deps } = mockFetch(`{"message":{"content":"Line1\\nLine2 \\"q\\""}}`);
    expect(await runPrompt("x", ollama, deps)).toBe('Line1\nLine2 "q"');
  });

  test("provider error object -> LlmError with message", async () => {
    const { deps } = mockFetch(`{"error":{"message":"invalid api key","type":"auth"}}`);
    await expect(runPrompt("x", openai, deps)).rejects.toThrow("invalid api key");
  });

  test("provider error string", async () => {
    const { deps } = mockFetch(`{"error":"model not found"}`);
    await expect(runPrompt("x", ollama, deps)).rejects.toThrow("model not found");
  });

  test("surfaces gateway metadata (OpenRouter) in the error", async () => {
    const { deps } = mockFetch(
      `{"error":{"message":"Provider returned error","code":429,"metadata":{"raw":"rate-limited by upstream","provider_name":"Together"}}}`
    );
    await expect(runPrompt("x", openai, deps)).rejects.toThrow(/rate-limited by upstream/);
  });

  test("http !ok surfaces the body's error message", async () => {
    const { deps } = mockFetch(`{"error":{"message":"rate limited"}}`, { ok: false, status: 429 });
    await expect(runPrompt("x", openai, deps)).rejects.toThrow("rate limited");
  });

  test("missing api key throws before any fetch", async () => {
    const { deps, calls } = mockFetch(`{}`);
    await expect(runPrompt("x", { provider: "openai", model: "gpt-4o-mini" }, deps)).rejects.toThrow(
      /No API key/
    );
    expect(calls.length).toBe(0);
  });

  test("unknown provider throws", async () => {
    const { deps } = mockFetch(`{}`);
    await expect(runPrompt("x", { provider: "nope", model: "m" }, deps)).rejects.toThrow(LlmError);
  });

  test("hits the correct endpoints", async () => {
    const a = mockFetch(`{"message":{"content":"ok"}}`);
    await runPrompt("x", ollama, a.deps);
    expect(a.calls[0].url).toContain("/api/chat");

    const b = mockFetch(`{"choices":[{"message":{"content":"ok"}}]}`);
    await runPrompt("x", openai, b.deps);
    expect(b.calls[0].url).toContain("/chat/completions");
    expect(b.calls[0].init.headers["Authorization"]).toBe("Bearer sk-test");
  });

  test("request body round-trips tricky characters", async () => {
    const { deps, calls } = mockFetch(`{"message":{"content":"ok"}}`);
    const userText = 'He said "hi"\nGrüße 😀 C:\\path';
    await runPrompt(userText, ollama, deps);
    const sent = JSON.parse(calls[0].init.body);
    expect(sent.model).toBe("test-model");
    expect(sent.stream).toBe(false);
    expect(sent.messages[1].content).toBe(userText);
  });
});

describe("listModels", () => {
  test("ollama", async () => {
    const { deps } = mockFetch(`{"models":[{"name":"llama3.2"},{"name":"mistral"}]}`);
    expect(await listModels(ollama, deps)).toEqual(["llama3.2", "mistral"]);
  });

  test("openai", async () => {
    const { deps } = mockFetch(`{"object":"list","data":[{"id":"gpt-4o"},{"id":"gpt-4o-mini"}]}`);
    expect(await listModels(openai, deps)).toEqual(["gpt-4o", "gpt-4o-mini"]);
  });
});

describe("proxy transport", () => {
  const proxied: LlmSettings = { provider: "openai", model: "gpt-4o-mini", proxyUrl: "https://proxy.example/api" };

  test("chat posts an envelope and returns content, no key needed", async () => {
    const { deps, calls } = mockFetch(`{"content":"from proxy"}`);
    expect(await runPrompt("hi", proxied, deps)).toBe("from proxy");
    expect(calls[0].url).toBe("https://proxy.example/api");
    const env = JSON.parse(calls[0].init.body);
    expect(env).toMatchObject({ op: "chat", provider: "openai", model: "gpt-4o-mini", prompt: "hi" });
  });

  test("models returns the proxy's list", async () => {
    const { deps } = mockFetch(`{"models":["a","b"]}`);
    expect(await listModels(proxied, deps)).toEqual(["a", "b"]);
  });

  test("proxy error surfaces", async () => {
    const { deps } = mockFetch(`{"error":"upstream down"}`, { ok: false, status: 502 });
    await expect(runPrompt("x", proxied, deps)).rejects.toThrow("upstream down");
  });
});

describe("buildChatBody / extractChatContent units", () => {
  test("ollama body carries stream:false, others don't", () => {
    expect(buildChatBody(PROVIDERS.ollama, "m", "p").stream).toBe(false);
    expect(buildChatBody(PROVIDERS.openai, "m", "p").stream).toBeUndefined();
  });

  test("extractChatContent prefers choices, falls back to message", () => {
    expect(extractChatContent(`{"choices":[{"message":{"content":"A"}}]}`)).toBe("A");
    expect(extractChatContent(`{"message":{"content":"B"}}`)).toBe("B");
  });
});
