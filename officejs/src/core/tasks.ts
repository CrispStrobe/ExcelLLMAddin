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

/**
 * Apply an instruction to every non-empty cell of a 2D range, preserving shape.
 * Runs with bounded concurrency. Empty cells pass through as empty strings;
 * per-cell failures become "Error: …" in that cell without failing the batch.
 */
export async function mapRange(
  values: unknown[][],
  instruction: string,
  settings: LlmSettings,
  deps: Deps,
  concurrency = 4
): Promise<string[][]> {
  const result: string[][] = values.map((row) => row.map(() => ""));
  const jobs: Array<{ r: number; c: number; text: string }> = [];

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const cell = values[r][c];
      const text = cell == null ? "" : String(cell);
      if (text.trim() !== "") jobs.push({ r, c, text });
    }
  }

  const system =
    "Apply the user's instruction to the single input value. " +
    "Output only the result for that value, as plain text.";

  await runPool(jobs, concurrency, async (job) => {
    try {
      const out = await runPrompt(
        `Instruction: ${instruction}\n\nInput: ${job.text}`,
        withSystem(settings, system),
        deps
      );
      result[job.r][job.c] = out.trim();
    } catch (e) {
      result[job.r][job.c] = "Error: " + (e instanceof Error ? e.message : String(e));
    }
  });

  return result;
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
