// Higher-level LLM operations built on runPrompt. Each sets a task-specific
// system prompt and post-processes the reply. Pure + testable (fetch injected).

import { runPrompt, embed, visionPrompt, LlmSettings, Deps } from "./llm";

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

/** Classify text sentiment into Positive / Neutral / Negative. */
export async function sentiment(
  text: string,
  settings: LlmSettings,
  deps: Deps
): Promise<string> {
  return classify(text, ["Positive", "Neutral", "Negative"], settings, deps);
}

/** Apply ALL matching labels from a set (multi-label). Returns a comma-joined subset. */
export async function tagText(
  text: string,
  categories: string[],
  settings: LlmSettings,
  deps: Deps
): Promise<string> {
  const cats = categories.map((c) => String(c).trim()).filter(Boolean);
  if (cats.length === 0) return "Error: no categories provided";
  const system =
    `Apply labels to the text. Choose ALL that apply from: ${cats.join(", ")}. ` +
    "Return only the matching labels as a comma-separated list, in the given order. If none apply, return nothing.";
  const out = await runPrompt(`Text:\n${text}`, withSystem(settings, system), deps);
  // Keep only recognized labels (order-preserving) so the result is always clean.
  const lc = out.toLowerCase();
  return cats.filter((c) => lc.includes(c.toLowerCase())).join(", ");
}

/** Rewrite/edit text per an instruction (default: fix spelling & grammar). */
export async function editText(
  text: string,
  instruction: string | undefined,
  settings: LlmSettings,
  deps: Deps
): Promise<string> {
  const what = instruction && instruction.trim() ? instruction.trim() : "Fix spelling and grammar";
  const system = "You are an editor. Apply the requested edit and output ONLY the revised text — no notes or quotes.";
  const out = await runPrompt(`Edit instruction: ${what}\n\nText:\n${text}`, withSystem(settings, system), deps);
  return out.trim();
}

/** Generate a 2D table from a prompt; first row is headers. Spills as a grid. */
export async function generateTable(
  prompt: string,
  settings: LlmSettings,
  deps: Deps
): Promise<string[][]> {
  const system =
    "Generate tabular data as a JSON array of arrays: each inner array is one row " +
    "of string cells, and the first row is the column headers. Keep every row the " +
    "same length. Return ONLY the JSON — no commentary or code fences.";
  const raw = await runPrompt(prompt, withSystem(settings, system), deps);
  const grid = parseGrid(raw);
  return grid ?? [["Error: could not parse a table from the response"]];
}

/**
 * Infer the pattern from example input→output pairs and apply it to new inputs
 * (Numerous-style "fill by example"). Batches into one call with a per-input fallback.
 */
export async function fillByExample(
  examples: Array<{ input: string; output: string }>,
  inputs: string[],
  settings: LlmSettings,
  deps: Deps
): Promise<string[]> {
  const ex = examples.filter((e) => e.input.trim() !== "" && e.output.trim() !== "");
  if (ex.length === 0) return inputs.map(() => "Error: need at least one example (input, output) pair");
  if (inputs.length === 0) return [];

  const system =
    "Infer the transformation from the examples and apply it to each new input. " +
    "Return ONLY a JSON array of output strings, exactly one per input, in order. No commentary.";
  const exBlock = ex.map((e, i) => `${i + 1}. IN: ${e.input}  =>  OUT: ${e.output}`).join("\n");
  const inBlock = inputs.map((t, i) => `${i + 1}. ${t}`).join("\n");
  const user =
    `Examples:\n${exBlock}\n\n` +
    `Apply the same transformation to these inputs and return a JSON array of exactly ${inputs.length} strings:\n${inBlock}`;

  try {
    const raw = await runPrompt(user, withSystem(settings, system), deps);
    const arr = parseStringArray(raw);
    if (arr && arr.length === inputs.length) return arr.map((v) => v.trim());
  } catch {
    /* fall through to per-input */
  }
  return Promise.all(inputs.map((t) => fillOne(ex, t, settings, deps).catch((e) => "Error: " + errMsg(e))));
}

async function fillOne(
  examples: Array<{ input: string; output: string }>,
  input: string,
  settings: LlmSettings,
  deps: Deps
): Promise<string> {
  const system =
    "Infer the transformation from the examples and output ONLY the result for the " +
    "new input — plain text, no labels or quotes.";
  const exBlock = examples.map((e) => `IN: ${e.input}  =>  OUT: ${e.output}`).join("\n");
  const out = await runPrompt(`${exBlock}\nIN: ${input}  =>  OUT:`, withSystem(settings, system), deps);
  return out.trim();
}

/** Ask a question about an image (URL or data: URI). Needs a vision-capable model. */
export async function analyzeImage(
  image: string,
  question: string,
  settings: LlmSettings,
  deps: Deps
): Promise<string> {
  const url = String(image).trim();
  if (!url) return "Error: no image URL or data: URI provided";
  const q = question && question.trim() ? question.trim() : "Describe this image.";
  const out = await visionPrompt(q, url, settings, deps);
  return out.trim();
}

/** Write an Excel formula from a natural-language description. Returns "=…". */
export async function writeFormula(
  description: string,
  settings: LlmSettings,
  deps: Deps
): Promise<string> {
  const system =
    "You write Microsoft Excel formulas. Output ONLY a single Excel formula that " +
    "starts with '=' — no explanation, no code fences, no surrounding text. Prefer " +
    "standard, widely-supported functions.";
  const out = await runPrompt(`Write an Excel formula that: ${description}`, withSystem(settings, system), deps);
  return cleanFormula(out);
}

/** Explain what an Excel formula does, in plain English. */
export async function explainFormula(
  formula: string,
  settings: LlmSettings,
  deps: Deps
): Promise<string> {
  const system =
    "Explain what the given Excel formula does, in plain English and concisely " +
    "(2-3 sentences max). Output only the explanation.";
  const out = await runPrompt(`Explain this Excel formula:\n${formula}`, withSystem(settings, system), deps);
  return out.trim();
}

/** Pull a clean "=…" formula out of a model reply (strip fences/prose). */
function cleanFormula(raw: string): string {
  let s = raw.trim();
  const fence = s.match(/```(?:[a-z]*)?\s*([\s\S]*?)```/i);
  if (fence) s = fence[1].trim();
  s = s.replace(/^`+|`+$/g, "").trim();
  const eq = s.indexOf("=");
  if (eq >= 0) s = s.slice(eq); // drop any leading prose before the '='
  s = s.split("\n")[0].trim(); // formula is a single line
  if (s && !s.startsWith("=")) s = "=" + s;
  return s;
}

/**
 * Ask the model for a list and return it as items. Prefers a JSON array; falls
 * back to splitting lines and stripping bullet/number prefixes.
 */
export async function listItems(
  prompt: string,
  count: number | undefined,
  settings: LlmSettings,
  deps: Deps
): Promise<string[]> {
  const n = count && count > 0 ? Math.floor(count) : undefined;
  const ask = n ? `${prompt}\n\nReturn exactly ${n} items.` : prompt;
  const system =
    "Answer as a JSON array of short strings — no commentary, no code fences.";
  const raw = await runPrompt(ask, withSystem(settings, system), deps);

  const arr = parseStringArray(raw);
  const items =
    arr ??
    raw
      .split(/\r?\n/)
      .map((l) => l.replace(/^\s*(?:[-*•]|\d+[.)])\s*/, "").trim())
      .filter(Boolean);

  return n ? items.slice(0, n) : items;
}

/** Extract several fields from text; returns one value per field, in order. */
export async function extractFields(
  text: string,
  fields: string[],
  settings: LlmSettings,
  deps: Deps
): Promise<string[]> {
  const fs = fields.map((f) => String(f).trim()).filter(Boolean);
  if (fs.length === 0) return [];

  const system =
    "Extract the requested fields from the text. Return ONLY a JSON array of " +
    "string values, one per field, in the given order. Use an empty string for " +
    "any field not present. No commentary or code fences.";
  const numbered = fs.map((f, i) => `${i + 1}. ${f}`).join("\n");
  const user = `Fields:\n${numbered}\n\nText:\n${text}\n\nReturn a JSON array of exactly ${fs.length} strings.`;

  try {
    const raw = await runPrompt(user, withSystem(settings, system), deps);
    const arr = parseStringArray(raw);
    if (arr && arr.length === fs.length) return arr.map((v) => v.trim());
  } catch {
    /* fall through to per-field */
  }

  return Promise.all(fs.map((f) => extract(text, f, settings, deps).catch((e) => "Error: " + errMsg(e))));
}

/** Answer a question using only the supplied context text. */
export async function ask(
  question: string,
  context: string,
  settings: LlmSettings,
  deps: Deps
): Promise<string> {
  const system =
    "Answer the question using only the provided context. If the answer is not " +
    "in the context, say so briefly. Output plain text suitable for a cell.";
  const out = await runPrompt(`Context:\n${context}\n\nQuestion: ${question}`, withSystem(settings, system), deps);
  return out.trim();
}

/** Cosine similarity of two texts' embeddings (1 = identical meaning). */
export async function similarity(
  a: string,
  b: string,
  model: string,
  settings: LlmSettings,
  deps: Deps
): Promise<number> {
  const [va, vb] = await Promise.all([
    embed(a, model, settings, deps),
    embed(b, model, settings, deps),
  ]);
  return cosine(va, vb);
}

export function cosine(a: number[], b: number[]): number {
  const n = Math.min(a.length, b.length);
  let dot = 0;
  let na = 0;
  let nb = 0;
  for (let i = 0; i < n; i++) {
    dot += a[i] * b[i];
    na += a[i] * a[i];
    nb += b[i] * b[i];
  }
  if (na === 0 || nb === 0) return 0;
  return dot / (Math.sqrt(na) * Math.sqrt(nb));
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
  const parsed = tolerantJsonArray(s.slice(start, end + 1));
  if (!Array.isArray(parsed)) return null;
  return parsed.map((v) => (v == null ? "" : typeof v === "string" ? v : JSON.stringify(v)));
}

/** Parse a JSON 2D array (rows of cells) from possibly-fenced output; null if none. */
function parseGrid(raw: string): string[][] | null {
  let s = raw.trim();
  const fence = s.match(/```(?:json)?\s*([\s\S]*?)```/i);
  if (fence) s = fence[1].trim();
  const start = s.indexOf("[");
  const end = s.lastIndexOf("]");
  if (start === -1 || end <= start) return null;
  const parsed = tolerantJsonArray(s.slice(start, end + 1));
  if (!Array.isArray(parsed) || parsed.length === 0) return null;
  const cell = (v: any) =>
    v == null ? "" : typeof v === "string" ? v : typeof v === "object" ? JSON.stringify(v) : String(v);
  // Rows-of-arrays, or a flat array we treat as a single column.
  if (Array.isArray((parsed as any[])[0])) {
    return (parsed as any[]).map((row) => (Array.isArray(row) ? row.map(cell) : [cell(row)]));
  }
  return (parsed as any[]).map((v) => [cell(v)]);
}

// Small/local models frequently emit *almost*-JSON arrays. Strict JSON.parse is
// tried first; only on failure do we repair the two most common quirks (a trailing
// comma before ] and single-quoted strings) and retry once. If the repair doesn't
// yield valid JSON we return null and the caller takes its reliable per-cell path,
// so a bad repair can never be worse than the un-repaired behavior.
function tolerantJsonArray(slice: string): unknown {
  try {
    return JSON.parse(slice);
  } catch {
    try {
      const repaired = slice.replace(/,\s*]/g, "]").replace(/'([^'\\]*)'/g, '"$1"');
      return JSON.parse(repaired);
    } catch {
      return null;
    }
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
