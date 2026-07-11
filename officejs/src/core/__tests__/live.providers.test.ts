// Live integration tests: drive the REAL runPrompt/listModels/embed against live
// provider endpoints using Node's global fetch. Skipped by default (no network in
// CI); enable with LIVE_PROVIDERS=1 and the relevant *_API_KEY env vars, e.g.
//
//   set -a; source ../.env; set +a
//   LIVE_PROVIDERS=1 npx jest live.providers
//
// Each provider's test self-skips when its key is absent, so you only exercise
// what you have credentials for. Unlike the curl probes, this runs the exact code
// path the add-in uses.

import { runPrompt, listModels, embed, FetchLike, LlmSettings } from "../llm";

const LIVE = process.env.LIVE_PROVIDERS === "1";
const suite = LIVE ? describe : describe.skip;

/* global fetch */
const nodeFetch: FetchLike = async (url, init) => {
  const r = await fetch(url, init as any);
  return { ok: r.ok, status: r.status, text: () => r.text() };
};
const deps = { fetch: nodeFetch };

interface Case {
  id: string;
  key: string | undefined;
  model: string;
  embedModel?: string;
}

const CHAT_CASES: Case[] = [
  { id: "groq", key: process.env.GROQ_API_KEY, model: "llama-3.3-70b-versatile" },
  { id: "cohere", key: process.env.COHERE_API_KEY, model: "command-r-08-2024" },
  { id: "huggingface", key: process.env.HF_TOKEN, model: "meta-llama/Llama-3.3-70B-Instruct" },
  { id: "openrouter", key: process.env.OPENROUTER_API_KEY, model: "meta-llama/llama-3.3-70b-instruct" },
  { id: "nebius", key: process.env.NEBIUS_API_KEY, model: "meta-llama/Llama-3.3-70B-Instruct" },
  { id: "mistral", key: process.env.MISTRAL_API_KEY, model: "mistral-small-latest" },
  { id: "together", key: process.env.TOGETHER_API_KEY, model: "meta-llama/Llama-3.3-70B-Instruct-Turbo-Free" },
  { id: "cerebras", key: process.env.CEREBRAS_API_KEY, model: "llama-3.3-70b" },
];

suite("LIVE providers", () => {
  jest.setTimeout(40000);

  describe("chat (runPrompt)", () => {
    for (const c of CHAT_CASES) {
      const t = c.key ? test : test.skip;
      t(`${c.id} returns a completion`, async () => {
        const settings: LlmSettings = { provider: c.id, model: c.model, apiKey: c.key };
        const out = await runPrompt("Reply with exactly one word: pong", settings, deps);
        expect(typeof out).toBe("string");
        expect(out.trim().length).toBeGreaterThan(0);
      });
    }
  });

  describe("listModels", () => {
    for (const c of CHAT_CASES) {
      const t = c.key ? test : test.skip;
      t(`${c.id} returns a non-empty model list`, async () => {
        const models = await listModels({ provider: c.id, model: c.model, apiKey: c.key }, deps);
        expect(Array.isArray(models)).toBe(true);
        expect(models.length).toBeGreaterThan(0);
      });
    }
  });

  describe("embeddings", () => {
    // Nebius embeddings were verified working; Gemini exposes /embeddings too.
    const EMBED_CASES: Case[] = [
      { id: "nebius", key: process.env.NEBIUS_API_KEY, model: "", embedModel: "Qwen/Qwen3-Embedding-8B" },
    ];
    for (const c of EMBED_CASES) {
      const t = c.key ? test : test.skip;
      t(`${c.id} returns a numeric vector`, async () => {
        const vec = await embed("hello world", c.embedModel!, { provider: c.id, model: "", apiKey: c.key }, deps);
        expect(Array.isArray(vec)).toBe(true);
        expect(vec.length).toBeGreaterThan(0);
        expect(typeof vec[0]).toBe("number");
      });
    }
  });
});
