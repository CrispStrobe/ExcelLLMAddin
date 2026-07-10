// Excel custom functions. JSDoc @customfunction tags are read at build time by
// custom-functions-metadata-plugin to generate functions.json. Each function is
// explicitly associated by id so registration is deterministic.
//
// Runtime note: the custom-functions runtime provides a global `fetch`; we adapt
// it to the core's injectable FetchLike so the same tested code path runs here.

import { runPrompt, listModels, LlmSettings, FetchLike } from "../core/llm";
import { loadSettings } from "../core/config";

/* global CustomFunctions, fetch */

const fetchLike: FetchLike = (url, init) =>
  fetch(url, init as RequestInit).then((r) => ({
    ok: r.ok,
    status: r.status,
    text: () => r.text(),
  }));

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
    return await runPrompt(text, await currentSettings(provider, model), { fetch: fetchLike });
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
    const models = await listModels(await currentSettings(provider), { fetch: fetchLike });
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

CustomFunctions.associate("PROMPT", prompt);
CustomFunctions.associate("LIST_MODELS", listModelsFn);
CustomFunctions.associate("CONFIG", config);
