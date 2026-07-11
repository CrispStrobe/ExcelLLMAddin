// Tests for the streaming chat driver. We install a fake `fetch` returning a
// ReadableStream-like body so the SSE/NDJSON reader path runs without a network,
// plus the non-streaming fallback (body without getReader) and error handling.

import { streamChat } from "../stream";
import { LlmSettings } from "../core/llm";

interface FetchOpts {
  ok?: boolean;
  status?: number;
  text?: string;
  chunks?: string[];
  noBody?: boolean;
}

function installFetch(opts: FetchOpts) {
  (global as any).fetch = async (_url: string, _init: any) => {
    const ok = opts.ok ?? true;
    if (!ok) return { ok, status: opts.status ?? 500, text: async () => opts.text ?? "", body: null };
    if (opts.noBody) return { ok, status: 200, text: async () => opts.text ?? "", body: null };
    const chunks = opts.chunks ?? [];
    let i = 0;
    const enc = new TextEncoder();
    const reader = {
      read: async () =>
        i < chunks.length ? { done: false, value: enc.encode(chunks[i++]) } : { done: true, value: undefined },
    };
    return { ok, status: 200, text: async () => opts.text ?? "", body: { getReader: () => reader } };
  };
}

const openai: LlmSettings = { provider: "openai", model: "gpt-4o-mini", apiKey: "sk-test" };
const ollama: LlmSettings = { provider: "ollama", model: "llama3.2" };

describe("streamChat", () => {
  test("accumulates OpenAI SSE deltas and reports each partial", async () => {
    installFetch({
      chunks: [
        'data: {"choices":[{"delta":{"content":"Hel"}}]}\n\n',
        'data: {"choices":[{"delta":{"content":"lo"}}]}\n\n',
        "data: [DONE]\n\n",
      ],
    });
    const partials: string[] = [];
    const full = await streamChat("hi", openai, (p) => partials.push(p));
    expect(full).toBe("Hello");
    expect(partials).toEqual(["Hel", "Hello"]);
  });

  test("parses Ollama NDJSON chunks", async () => {
    installFetch({
      chunks: ['{"message":{"content":"a"}}\n', '{"message":{"content":"b"}}\n'],
    });
    const partials: string[] = [];
    const full = await streamChat("hi", ollama, (p) => partials.push(p));
    expect(full).toBe("ab");
    expect(partials).toEqual(["a", "ab"]);
  });

  test("falls back to a single read when the body cannot stream", async () => {
    installFetch({ noBody: true, text: 'data: {"choices":[{"delta":{"content":"Hi"}}]}' });
    const partials: string[] = [];
    const full = await streamChat("hi", openai, (p) => partials.push(p));
    expect(full).toBe("Hi");
    expect(partials).toEqual(["Hi"]);
  });

  test("surfaces a provider error message on a non-ok response", async () => {
    installFetch({ ok: false, status: 429, text: '{"error":{"message":"rate limited"}}' });
    await expect(streamChat("hi", openai, () => {})).rejects.toThrow("rate limited");
  });

  test("throws when a key is required but missing", async () => {
    installFetch({ chunks: [] });
    await expect(streamChat("hi", { ...openai, apiKey: "" }, () => {})).rejects.toThrow(/API key/i);
  });

  test("throws on an unknown provider", async () => {
    installFetch({ chunks: [] });
    await expect(streamChat("hi", { ...openai, provider: "bogus" }, () => {})).rejects.toThrow(/Unknown provider/i);
  });
});
