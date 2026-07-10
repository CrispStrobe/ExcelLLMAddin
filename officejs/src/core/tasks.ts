// Higher-level LLM operations built on runPrompt. Each sets a task-specific
// system prompt and post-processes the reply. Pure + testable (fetch injected).

import { runPrompt, LlmSettings, Deps } from "./llm";

function withSystem(settings: LlmSettings, system: string): LlmSettings {
  return { ...settings, systemPrompt: system };
}

/** Classify text into exactly one of the supplied categories. */
export async function classify(
  text: string,
  categories: string[],
  settings: LlmSettings,
  deps: Deps
): Promise<string> {
  const cats = categories.map((c) => String(c).trim()).filter(Boolean);
  if (cats.length === 0) return "Error: no categories provided";
  const system =
    "You are a precise text classifier. Respond with EXACTLY ONE of the " +
    `following labels and nothing else: ${cats.join(", ")}.`;
  const out = await runPrompt(`Classify this text:\n${text}`, withSystem(settings, system), deps);
  return matchCategory(out, cats);
}

/** Extract a requested value from text; empty string if absent. */
export async function extract(
  text: string,
  instruction: string,
  settings: LlmSettings,
  deps: Deps
): Promise<string> {
  const system =
    "Extract the requested information from the text. Output only the extracted " +
    "value as plain text — no labels, quotes, or explanation. If it is not present, output nothing.";
  const out = await runPrompt(
    `From the following text, extract: ${instruction}\n\nText:\n${text}`,
    withSystem(settings, system),
    deps
  );
  return out.trim();
}

/** Translate text into a target language. */
export async function translate(
  text: string,
  targetLanguage: string,
  settings: LlmSettings,
  deps: Deps
): Promise<string> {
  const system = "You are a translator. Output only the translation, with no notes or quotes.";
  const out = await runPrompt(
    `Translate the following into ${targetLanguage}:\n\n${text}`,
    withSystem(settings, system),
    deps
  );
  return out.trim();
}

/** Summarize text, optionally capped to maxWords. */
export async function summarize(
  text: string,
  maxWords: number | undefined,
  settings: LlmSettings,
  deps: Deps
): Promise<string> {
  const limit = maxWords && maxWords > 0 ? ` in at most ${maxWords} words` : "";
  const system = "You are a concise summarizer. Output only the summary.";
  const out = await runPrompt(
    `Summarize the following${limit}:\n\n${text}`,
    withSystem(settings, system),
    deps
  );
  return out.trim();
}

export interface MapOptions {
  /** Number of chunks processed concurrently. */
  concurrency?: number;
  /** Cells per model call. >1 batches (one call returns a JSON array). */
  batchSize?: number;
}

/**
 * Apply an instruction to every non-empty cell of a 2D range, preserving shape.
 * Cells are batched (batchSize per call) to cut cost/latency: each call asks for
 * a JSON array of results. If the array doesn't line up (wrong length / not
 * JSON), that chunk falls back to reliable per-cell calls. Empty cells pass
 * through; per-cell failures become "Error: …" without failing the whole batch.
 */
export async function mapRange(
  values: unknown[][],
  instruction: string,
  settings: LlmSettings,
  deps: Deps,
  options: MapOptions = {}
): Promise<string[][]> {
  const batchSize = Math.max(1, options.batchSize ?? 20);
  const concurrency = Math.max(1, options.concurrency ?? 4);

  const result: string[][] = values.map((row) => row.map(() => ""));
  const jobs: Array<{ r: number; c: number; text: string }> = [];
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const cell = values[r][c];
      const text = cell == null ? "" : String(cell);
      if (text.trim() !== "") jobs.push({ r, c, text });
    }
  }

  const chunks = chunkArray(jobs, batchSize);
  await runPool(chunks, concurrency, async (group) => {
    const outputs = await mapChunk(group.map((j) => j.text), instruction, settings, deps);
    group.forEach((j, i) => {
      result[j.r][j.c] = outputs[i];
    });
  });

  return result;
}

async function mapOne(
  text: string,
  instruction: string,
  settings: LlmSettings,
  deps: Deps
): Promise<string> {
  const system =
    "Apply the user's instruction to the single input value. " +
    "Output only the result for that value, as plain text.";
  const out = await runPrompt(
    `Instruction: ${instruction}\n\nInput: ${text}`,
    withSystem(settings, system),
    deps
  );
  return out.trim();
}

async function mapChunk(
  inputs: string[],
  instruction: string,
  settings: LlmSettings,
  deps: Deps
): Promise<string[]> {
  if (inputs.length === 1) {
    try {
      return [await mapOne(inputs[0], instruction, settings, deps)];
    } catch (e) {
      return ["Error: " + errMsg(e)];
    }
  }

  const system =
    "Apply the instruction to each numbered input. Return ONLY a JSON array of " +
    "strings — exactly one result per input, in the same order, no commentary or code fences.";
  const numbered = inputs.map((t, i) => `${i + 1}. ${t}`).join("\n");
  const user =
    `Instruction: ${instruction}\n\nInputs:\n${numbered}\n\n` +
    `Return a JSON array of exactly ${inputs.length} strings.`;

  try {
    const raw = await runPrompt(user, withSystem(settings, system), deps);
    const arr = parseStringArray(raw);
    if (arr && arr.length === inputs.length) return arr.map((v) => v.trim());
  } catch {
    /* fall through to per-cell */
  }

  // Fallback: reliable, order-safe per-cell calls.
  return Promise.all(
    inputs.map((t) => mapOne(t, instruction, settings, deps).catch((e) => "Error: " + errMsg(e)))
  );
}

function parseStringArray(raw: string): string[] | null {
  let s = raw.trim();
  const fence = s.match(/```(?:json)?\s*([\s\S]*?)```/i);
  if (fence) s = fence[1].trim();
  const start = s.indexOf("[");
  const end = s.lastIndexOf("]");
  if (start === -1 || end <= start) return null;
  try {
    const parsed = JSON.parse(s.slice(start, end + 1));
    if (!Array.isArray(parsed)) return null;
    return parsed.map((v) => (v == null ? "" : typeof v === "string" ? v : JSON.stringify(v)));
  } catch {
    return null;
  }
}

function chunkArray<T>(arr: T[], size: number): T[][] {
  const out: T[][] = [];
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
}

function errMsg(e: unknown): string {
  return e instanceof Error ? e.message : String(e);
}

// ---- helpers ----------------------------------------------------------------

export function matchCategory(output: string, categories: string[]): string {
  const lower = output.trim().toLowerCase();
  const exact = categories.find((c) => c.toLowerCase() === lower);
  if (exact) return exact;
  const contained = categories.find((c) => lower.includes(c.toLowerCase()));
  if (contained) return contained;
  return output.trim();
}

async function runPool<T>(items: T[], limit: number, worker: (t: T) => Promise<void>): Promise<void> {
  let next = 0;
  const size = Math.max(1, Math.min(limit, items.length || 1));
  const runners = Array.from({ length: size }, async () => {
    while (next < items.length) {
      const idx = next++;
      await worker(items[idx]);
    }
  });
  await Promise.all(runners);
}
