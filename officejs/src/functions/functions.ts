// Excel custom functions. JSDoc @customfunction tags are read at build time by
// custom-functions-metadata-plugin to generate functions.json. Each function is
// explicitly associated by id so registration is deterministic.
//
// Runtime note: the custom-functions runtime provides a global `fetch`; we adapt
// it to the core's injectable FetchLike so the same tested code path runs here.

import { runPrompt, listModels, LlmSettings } from "../core/llm";
import {
  classify,
  extract,
  translate,
  summarize,
  mapRange,
  sentiment,
  listItems,
  extractFields,
  ask,
  similarity,
} from "../core/tasks";
import { loadSettings } from "../core/config";
import { resilientFetch as fetchLike } from "../browserFetch";
import { createLruCache } from "../core/cache";
import { streamChat } from "../stream";

/* global CustomFunctions */

// Shared deps for all custom functions. The session cache means identical
// (provider, model, prompt) calls don't re-hit the API on every Excel recalc;
// errors are never cached. It resets when the functions runtime reloads.
const deps = { fetch: fetchLike, cache: createLruCache(500) };

async function currentSettings(provider?: string, model?: string): Promise<LlmSettings> {
  const s = await loadSettings();
  return {
    ...s,
    provider: (provider && provider.trim()) || s.provider,
    model: (model && model.trim()) || s.model,
    baseUrl: s.baseUrl || undefined,
    apiKey: s.apiKey || undefined,
    proxyUrl: s.proxyUrl || undefined,
    systemPrompt: s.systemPrompt || undefined,
    embedModel: s.embedModel || undefined,
  };
}

function errorText(e: unknown): string {
  return "Error: " + (e instanceof Error ? e.message : String(e));
}

/**
 * Sends a prompt to the configured LLM and returns its reply.
 * @customfunction PROMPT
 * @param text The prompt text (or a cell reference).
 * @param provider Optional provider id (openai, mistral, nebius, scaleway, openrouter, ollama).
 * @param model Optional model name.
 * @returns The model's reply.
 */
export async function prompt(text: string, provider?: string, model?: string): Promise<string> {
  try {
    return await runPrompt(text, await currentSettings(provider, model), deps);
  } catch (e) {
    return errorText(e);
  }
}

/**
 * Lists available models for a provider, spilling one per row.
 * @customfunction LIST_MODELS
 * @param provider Optional provider id; defaults to the configured provider.
 * @returns A single-column list of model ids.
 */
export async function listModelsFn(provider?: string): Promise<string[][]> {
  try {
    const models = await listModels(await currentSettings(provider), deps);
    return models.length ? models.map((m) => [m]) : [["(no models)"]];
  } catch (e) {
    return [[errorText(e)]];
  }
}

/**
 * Shows the current provider and model.
 * @customfunction CONFIG
 * @returns Provider and model summary.
 */
export async function config(): Promise<string> {
  const s = await loadSettings();
  return `Provider: ${s.provider} | Model: ${s.model}` + (s.proxyUrl ? " | via proxy" : "");
}

/**
 * Classifies text into exactly one of the given category labels.
 * @customfunction CLASSIFY
 * @param text The text to classify (or a cell reference).
 * @param categories A range or array of candidate labels.
 * @returns The chosen label.
 */
export async function classifyFn(text: string, categories: string[][]): Promise<string> {
  try {
    return await classify(text, flatten(categories), await currentSettings(), deps);
  } catch (e) {
    return errorText(e);
  }
}

/**
 * Extracts a requested value from text.
 * @customfunction EXTRACT
 * @param text The source text (or a cell reference).
 * @param instruction What to extract, e.g. "the email address".
 * @returns The extracted value, or empty if not found.
 */
export async function extractFn(text: string, instruction: string): Promise<string> {
  try {
    return await extract(text, instruction, await currentSettings(), deps);
  } catch (e) {
    return errorText(e);
  }
}

/**
 * Translates text into a target language.
 * @customfunction TRANSLATE
 * @param text The text to translate (or a cell reference).
 * @param targetLanguage The target language, e.g. "German".
 * @returns The translation.
 */
export async function translateFn(text: string, targetLanguage: string): Promise<string> {
  try {
    return await translate(text, targetLanguage, await currentSettings(), deps);
  } catch (e) {
    return errorText(e);
  }
}

/**
 * Summarizes text, optionally capped to a word count.
 * @customfunction SUMMARIZE
 * @param text The text to summarize (or a cell reference).
 * @param maxWords Optional maximum number of words.
 * @returns The summary.
 */
export async function summarizeFn(text: string, maxWords?: number): Promise<string> {
  try {
    return await summarize(text, maxWords, await currentSettings(), deps);
  } catch (e) {
    return errorText(e);
  }
}

/**
 * Applies an instruction to every cell of a range, spilling the results.
 * Cells are batched into few calls (with a per-cell fallback) to keep it fast.
 * @customfunction MAP
 * @param range The range of input values.
 * @param instruction The instruction to apply to each cell.
 * @returns A range of results with the same shape as the input.
 */
export async function mapFn(range: string[][], instruction: string): Promise<string[][]> {
  try {
    return await mapRange(range, instruction, await currentSettings(), deps);
  } catch (e) {
    return [[errorText(e)]];
  }
}

function flatten(grid: string[][]): string[] {
  const out: string[] = [];
  for (const row of grid || []) {
    for (const cell of row || []) {
      const s = String(cell ?? "").trim();
      if (s === "") continue;
      // Allow a single "a, b, c" cell as well as a range of labels.
      if (s.includes(",")) out.push(...s.split(",").map((x) => x.trim()).filter(Boolean));
      else out.push(s);
    }
  }
  return out;
}

/**
 * Classifies the sentiment of text as Positive, Neutral, or Negative.
 * @customfunction SENTIMENT
 * @param text The text (or a cell reference).
 * @returns Positive, Neutral, or Negative.
 */
export async function sentimentFn(text: string): Promise<string> {
  try {
    return await sentiment(text, await currentSettings(), deps);
  } catch (e) {
    return errorText(e);
  }
}

/**
 * Generates a list from a prompt and spills it down a column.
 * @customfunction LIST
 * @param prompt What to list, e.g. "synonyms of happy".
 * @param count Optional number of items.
 * @returns A single-column list.
 */
export async function listFn(prompt: string, count?: number): Promise<string[][]> {
  try {
    const items = await listItems(prompt, count, await currentSettings(), deps);
    return items.length ? items.map((i) => [i]) : [["(no items)"]];
  } catch (e) {
    return [[errorText(e)]];
  }
}

/**
 * Extracts several fields from text, spilling one value per field across a row.
 * @customfunction FIELDS
 * @param text The source text (or a cell reference).
 * @param fields A range or list of field names/descriptions to extract.
 * @returns A single row of extracted values, aligned to the fields.
 */
export async function fieldsFn(text: string, fields: string[][]): Promise<string[][]> {
  try {
    const values = await extractFields(text, flatten(fields), await currentSettings(), deps);
    return [values.length ? values : ["(no fields)"]];
  } catch (e) {
    return [[errorText(e)]];
  }
}

/**
 * Answers a question using a range of cells as context.
 * @customfunction ASK
 * @param question The question to answer.
 * @param context A range whose contents are the context.
 * @returns The answer.
 */
export async function askFn(question: string, context: string[][]): Promise<string> {
  try {
    const ctx = (context || [])
      .map((row) => (row || []).map((c) => String(c ?? "")).join("\t"))
      .join("\n");
    return await ask(question, ctx, await currentSettings(), deps);
  } catch (e) {
    return errorText(e);
  }
}

/**
 * Semantic similarity of two texts (1 = same meaning), via embeddings.
 * Uses the embedding model from LLM Settings unless one is given.
 * @customfunction SIMILARITY
 * @param a First text (or a cell reference).
 * @param b Second text (or a cell reference).
 * @param model Optional embedding model id.
 * @returns Cosine similarity, typically 0..1.
 */
export async function similarityFn(a: string, b: string, model?: string): Promise<number | string> {
  try {
    const s = await currentSettings();
    const m = (model && model.trim()) || s.embedModel || "";
    return await similarity(a, b, m, s, deps);
  } catch (e) {
    return errorText(e);
  }
}

/**
 * Streams the model's reply into the cell as it is generated.
 * @customfunction STREAM
 * @param text The prompt text (or a cell reference).
 * @param provider Optional provider id.
 * @param model Optional model name.
 * @param invocation
 * @returns The streamed reply.
 */
export function streamFn(
  text: string,
  provider: string,
  model: string,
  invocation: CustomFunctions.StreamingInvocation<string>
): void {
  let canceled = false;
  invocation.onCanceled = () => {
    canceled = true;
  };
  invocation.setResult("…");
  (async () => {
    try {
      const s = await currentSettings(provider, model);
      const full = await streamChat(text, s, (partial) => {
        if (!canceled) invocation.setResult(partial);
      });
      if (!canceled) invocation.setResult(full || "(empty)");
    } catch (e) {
      if (!canceled) invocation.setResult(errorText(e));
    }
  })();
}

CustomFunctions.associate("PROMPT", prompt);
CustomFunctions.associate("STREAM", streamFn);
CustomFunctions.associate("LIST_MODELS", listModelsFn);
CustomFunctions.associate("CONFIG", config);
CustomFunctions.associate("CLASSIFY", classifyFn);
CustomFunctions.associate("EXTRACT", extractFn);
CustomFunctions.associate("TRANSLATE", translateFn);
CustomFunctions.associate("SUMMARIZE", summarizeFn);
CustomFunctions.associate("MAP", mapFn);
CustomFunctions.associate("SENTIMENT", sentimentFn);
CustomFunctions.associate("LIST", listFn);
CustomFunctions.associate("FIELDS", fieldsFn);
CustomFunctions.associate("ASK", askFn);
CustomFunctions.associate("SIMILARITY", similarityFn);
