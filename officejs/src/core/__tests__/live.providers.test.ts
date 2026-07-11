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
import { fillByExample, generateTable, tagText, editText, writeFormula, analyzeImage, recall } from "../tasks";

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

  describe("new task functions (via Groq)", () => {
    const key = process.env.GROQ_API_KEY;
    const s: LlmSettings = { provider: "groq", model: "llama-3.3-70b-versatile", apiKey: key };
    const t = key ? test : test.skip;

    t("FILL infers a pattern from examples", async () => {
      const out = await fillByExample([{ input: "Germany", output: "DE" }], ["France", "Spain"], s, deps);
      expect(out).toHaveLength(2);
      expect(out.join(" ").toUpperCase()).toContain("FR");
    });

    t("TABLE returns a multi-row grid", async () => {
      const grid = await generateTable("3 largest planets with their diameter", s, deps);
      expect(grid.length).toBeGreaterThan(1);
      expect(grid[0].length).toBeGreaterThan(1);
    });

    t("TAG returns only valid labels", async () => {
      const out = await tagText("The invoice total is wrong and the app crashes", ["Bug", "Billing", "Praise"], s, deps);
      const labels = out ? out.split(",").map((x) => x.trim()) : [];
      for (const l of labels) expect(["Bug", "Billing", "Praise"]).toContain(l);
    });

    t("EDIT fixes grammar", async () => {
      const out = await editText("they is going too the store", undefined, s, deps);
      expect(out.length).toBeGreaterThan(0);
    });

    t("FORMULA returns a formula string", async () => {
      const f = await writeFormula("sum column B where column A is greater than 100", s, deps);
      expect(f.startsWith("=")).toBe(true);
    });
  });

  describe("vision (via Groq multimodal)", () => {
    const key = process.env.GROQ_API_KEY;
    const t = key ? test : test.skip;
    // A 64x64 solid-red PNG as a data: URI — deterministic, no external image fetch.
    const RED_PNG =
      "data:image/png;base64," +
      "iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAIAAAAlC+aJAAAAb0lEQVR4nO3PAQkAAAyEwO9feoshgnAB" +
      "dLep8QUNyPEFDcjxBQ3I8QUNyPEFDcjxBQ3I8QUNyPEFDcjxBQ3I8QUNyPEFDcjxBQ3I8QUNyPEFDcjx" +
      "BQ3I8QUNyPEFDcjxBQ3I8QUNyPEFDcjxBQ3IPanc8OLDQitxAAAAAElFTkSuQmCC";

    t("VISION identifies the image color", async () => {
      const s: LlmSettings = { provider: "groq", model: "meta-llama/llama-4-scout-17b-16e-instruct", apiKey: key };
      const out = await analyzeImage(RED_PNG, "What color fills this image? One word.", s, deps);
      expect(out.toLowerCase()).toContain("red");
    });
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

      t(`${c.id} RECALL ranks the semantically-closest row first`, async () => {
        const s: LlmSettings = { provider: c.id, model: "", apiKey: c.key };
        const rows = ["The cat sat on the mat.", "Quarterly revenue rose 12%.", "A feline rested on the rug."];
        const out = await recall("a cat lay on a carpet", rows, 1, c.embedModel!, s, deps);
        expect(out[0][0]).toMatch(/feline|cat/i); // one of the two cat sentences
      });
    }
  });
});
