// Pure LLM client. No Office / DOM / global-fetch dependency: `fetch` is
// injected, so this whole module is unit-testable with a mock (see __tests__).
//
// Two transport modes:
//   - direct: call the provider API from the browser. Works for browser-friendly
//     providers (OpenRouter, local Ollama); most others block it via CORS.
//   - proxy:  POST a normalized envelope to a serverless proxy that holds the
//     API keys server-side and adds CORS. Recommended for production + secrets.

import {
  ProviderSpec,
  ProviderStyle,
  getProvider,
  chatEndpoint,
  modelsEndpoint,
  embeddingsEndpoint,
} from "./providers";
import { parseUsage, TokenUsage } from "./usage";

/** Report token usage parsed from a raw (direct) response body, if any. */
function reportDirectUsage(deps: Deps, text: string, style: ProviderStyle): void {
  if (!deps.onUsage) return;
  const u = parseUsage(text, style);
  if (u) deps.onUsage(u);
}

/** Report token usage from a proxy envelope's `usage` object, if present. */
function reportProxyUsage(deps: Deps, usage: unknown): void {
  if (!deps.onUsage || !usage || typeof usage !== "object") return;
  const u = usage as Record<string, unknown>;
  const p = Number(u.prompt_tokens) || 0;
  const c = Number(u.completion_tokens) || 0;
  const t = Number(u.total_tokens) || p + c;
  if (p || c || t) deps.onUsage({ promptTokens: p, completionTokens: c, totalTokens: t });
}

export class LlmError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "LlmError";
  }
}

export interface LlmSettings {
  provider: string;
  model: string;
  /** Override the provider's default base URL. */
  baseUrl?: string;
  /** Used only for direct (non-proxy) calls. */
  apiKey?: string;
  /** If set, route through this proxy; keys are held server-side. */
  proxyUrl?: string;
  systemPrompt?: string;
  /** Embedding model (for SIMILARITY); provider-specific, e.g. Nebius Qwen/Qwen3-Embedding-8B. */
  embedModel?: string;
  /** Optional remote MCP server URL; its tools are added to the agent. */
  mcpUrl?: string;
  /** Black Forest Labs API key for =IMAGE (direct calls; CORS may require the proxy). */
  imageApiKey?: string;
  /** BFL image model slug for =IMAGE, e.g. flux-dev. */
  imageModel?: string;
}

export type FetchLike = (
  input: string,
  init?: {
    method?: string;
    headers?: Record<string, string>;
    body?: string;
  }
) => Promise<{
  ok: boolean;
  status: number;
  text: () => Promise<string>;
}>;

/** Optional response cache. get returns undefined on a miss. */
export interface LlmCache {
  get(key: string): string | undefined;
  set(key: string, value: string): void;
  clear?(): void;
}

export interface Deps {
  fetch: FetchLike;
  /** If present, successful prompt results are served/stored here. */
  cache?: LlmCache;
  /** Called with token usage parsed from each direct (non-proxy) response. */
  onUsage?: (usage: TokenUsage) => void;
}

export const DEFAULT_SYSTEM_PROMPT =
  "You are a helpful assistant embedded in a spreadsheet. " +
  "Answer concisely and return plain text suitable for a cell unless asked otherwise.";

export async function runPrompt(
  promptText: string,
  settings: LlmSettings,
  deps: Deps
): Promise<string> {
  const spec = requireProvider(settings.provider);

  const cacheKey = deps.cache ? promptCacheKey(settings, promptText) : "";
  if (deps.cache) {
    const hit = deps.cache.get(cacheKey);
    if (hit !== undefined) return hit;
  }

  let result: string;
  if (settings.proxyUrl) {
    const data = await callProxy("chat", promptText, settings, spec, deps);
    result = String(data.content ?? "");
    reportProxyUsage(deps, data.usage);
  } else {
    const baseUrl = settings.baseUrl || spec.defaultBaseUrl;
    if (spec.requiresKey && !settings.apiKey) {
      throw new LlmError(`No API key configured for ${spec.label}.`);
    }
    const url = chatEndpoint(spec, baseUrl);
    const body = buildChatBody(spec, settings.model, promptText, settings.systemPrompt);
    const resp = await deps.fetch(url, {
      method: "POST",
      headers: directHeaders(spec, settings.apiKey),
      body: JSON.stringify(body),
    });
    const text = await resp.text();
    if (!resp.ok) {
      throw new LlmError(parseErrorMessage(text) ?? `HTTP ${resp.status} from ${url}`);
    }
    result = extractChatContent(text);
    reportDirectUsage(deps, text, spec.style);
  }

  // Only successful results reach here (errors throw above), so errors aren't cached.
  if (deps.cache) deps.cache.set(cacheKey, result);
  return result;
}

function promptCacheKey(settings: LlmSettings, promptText: string): string {
  return JSON.stringify([settings.provider, settings.model, settings.systemPrompt ?? "", promptText]);
}

/**
 * Ask a question about an image (multimodal). Uses the OpenAI-compatible
 * content-array shape, so it works for openai-style providers with a vision model
 * (OpenAI, Gemini, OpenRouter, ...). Direct only — the proxy/Ollama paths use a
 * different multimodal shape and aren't supported here.
 */
export async function visionPrompt(
  question: string,
  imageUrl: string,
  settings: LlmSettings,
  deps: Deps
): Promise<string> {
  const spec = requireProvider(settings.provider);
  if (settings.proxyUrl) throw new LlmError("VISION requires a direct provider + key (not the proxy).");
  if (spec.style === "ollama") throw new LlmError("VISION isn't supported for Ollama yet (different image format).");
  const baseUrl = settings.baseUrl || spec.defaultBaseUrl;
  if (spec.requiresKey && !settings.apiKey) throw new LlmError(`No API key configured for ${spec.label}.`);

  const url = chatEndpoint(spec, baseUrl);
  const body = {
    model: settings.model,
    messages: [
      { role: "system", content: settings.systemPrompt ?? DEFAULT_SYSTEM_PROMPT },
      {
        role: "user",
        content: [
          { type: "text", text: question },
          { type: "image_url", image_url: { url: imageUrl } },
        ],
      },
    ],
  };
  const resp = await deps.fetch(url, {
    method: "POST",
    headers: directHeaders(spec, settings.apiKey),
    body: JSON.stringify(body),
  });
  const text = await resp.text();
  if (!resp.ok) throw new LlmError(parseErrorMessage(text) ?? `HTTP ${resp.status} from ${url}`);
  return extractChatContent(text);
}

export async function listModels(
  settings: LlmSettings,
  deps: Deps
): Promise<string[]> {
  const spec = requireProvider(settings.provider);

  if (settings.proxyUrl) {
    const data = await callProxy("models", "", settings, spec, deps);
    return Array.isArray(data.models) ? data.models.map(String) : [];
  }

  const baseUrl = settings.baseUrl || spec.defaultBaseUrl;
  if (spec.requiresKey && !settings.apiKey) {
    throw new LlmError(`No API key configured for ${spec.label}.`);
  }

  const url = modelsEndpoint(spec, baseUrl);
  const resp = await deps.fetch(url, {
    method: "GET",
    headers: directHeaders(spec, settings.apiKey),
  });
  const text = await resp.text();
  if (!resp.ok) {
    throw new LlmError(parseErrorMessage(text) ?? `HTTP ${resp.status} from ${url}`);
  }
  return extractModelList(spec, text);
}

/** Get an embedding vector for text. Uses settings.embedModel unless `model` is given. */
export async function embed(
  text: string,
  model: string,
  settings: LlmSettings,
  deps: Deps
): Promise<number[]> {
  const spec = requireProvider(settings.provider);
  if (!model) {
    throw new LlmError("No embedding model set — pass one to SIMILARITY or set it in LLM Settings.");
  }

  const cacheKey = deps.cache ? JSON.stringify(["embed", settings.provider, model, text]) : "";
  if (deps.cache) {
    const hit = deps.cache.get(cacheKey);
    if (hit !== undefined) return JSON.parse(hit) as number[];
  }

  let vec: number[];
  if (settings.proxyUrl) {
    const data = await callProxy("embed", text, { ...settings, model }, spec, deps);
    if (!Array.isArray(data.embedding)) throw new LlmError("Proxy returned no embedding.");
    vec = (data.embedding as unknown[]).map(Number);
    reportProxyUsage(deps, data.usage);
  } else {
    const baseUrl = settings.baseUrl || spec.defaultBaseUrl;
    if (spec.requiresKey && !settings.apiKey) {
      throw new LlmError(`No API key configured for ${spec.label}.`);
    }
    const url = embeddingsEndpoint(spec, baseUrl);
    const body = spec.style === "ollama" ? { model, prompt: text } : { model, input: text };
    const resp = await deps.fetch(url, {
      method: "POST",
      headers: directHeaders(spec, settings.apiKey),
      body: JSON.stringify(body),
    });
    const t = await resp.text();
    if (!resp.ok) throw new LlmError(parseErrorMessage(t) ?? `HTTP ${resp.status} from ${url}`);
    vec = extractEmbedding(t);
    reportDirectUsage(deps, t, spec.style);
  }

  if (deps.cache) deps.cache.set(cacheKey, JSON.stringify(vec));
  return vec;
}

function embedCacheKey(settings: LlmSettings, model: string, text: string): string {
  return JSON.stringify(["embed", settings.provider, model, text]);
}

/**
 * Embed many texts, preferring ONE request over N. OpenAI-style embedding
 * endpoints accept an `input` array and return `data[i].embedding`, so a whole
 * range collapses to a single call (or a few, if the caller chunks). Ollama has
 * no array-input shape in the endpoint we use, so it falls back to cache-aware
 * single calls. The per-text cache is honored: only cache misses are requested,
 * and results are written back, so this stays consistent with embed().
 */
export async function embedBatch(
  texts: string[],
  model: string,
  settings: LlmSettings,
  deps: Deps
): Promise<number[][]> {
  const spec = requireProvider(settings.provider);
  if (!model) {
    throw new LlmError("No embedding model set — pass one to SIMILARITY or set it in LLM Settings.");
  }
  if (texts.length === 0) return [];

  // Serve per-text cache hits; collect the indices that still need a request.
  const out: number[][] = new Array(texts.length);
  const missIdx: number[] = [];
  for (let i = 0; i < texts.length; i++) {
    const hit = deps.cache ? deps.cache.get(embedCacheKey(settings, model, texts[i])) : undefined;
    if (hit !== undefined) out[i] = JSON.parse(hit) as number[];
    else missIdx.push(i);
  }
  if (missIdx.length === 0) return out;

  // Ollama: no array-input endpoint here, so embed the misses one by one
  // (embed() is itself cache-aware, so this also fills the cache).
  if (spec.style === "ollama" && !settings.proxyUrl) {
    for (const i of missIdx) out[i] = await embed(texts[i], model, settings, deps);
    return out;
  }

  const missTexts = missIdx.map((i) => texts[i]);
  let vecs: number[][];
  if (settings.proxyUrl) {
    const data = await callProxy("embed", "", { ...settings, model }, spec, deps, { inputs: missTexts });
    if (!Array.isArray(data.embeddings)) throw new LlmError("Proxy returned no embeddings.");
    vecs = (data.embeddings as unknown[]).map((v) => (v as unknown[]).map(Number));
    reportProxyUsage(deps, data.usage);
  } else {
    const baseUrl = settings.baseUrl || spec.defaultBaseUrl;
    if (spec.requiresKey && !settings.apiKey) {
      throw new LlmError(`No API key configured for ${spec.label}.`);
    }
    const url = embeddingsEndpoint(spec, baseUrl);
    const resp = await deps.fetch(url, {
      method: "POST",
      headers: directHeaders(spec, settings.apiKey),
      body: JSON.stringify({ model, input: missTexts }),
    });
    const t = await resp.text();
    if (!resp.ok) throw new LlmError(parseErrorMessage(t) ?? `HTTP ${resp.status} from ${url}`);
    vecs = extractEmbeddingList(t, missTexts.length);
    reportDirectUsage(deps, t, spec.style);
  }

  if (vecs.length !== missTexts.length) {
    throw new LlmError(`Embedding count mismatch: asked for ${missTexts.length}, got ${vecs.length}.`);
  }
  for (let k = 0; k < missIdx.length; k++) {
    out[missIdx[k]] = vecs[k];
    if (deps.cache) deps.cache.set(embedCacheKey(settings, model, missTexts[k]), JSON.stringify(vecs[k]));
  }
  return out;
}

/** Parse an OpenAI-style batch embeddings body into vectors, ordered by index. */
function extractEmbeddingList(text: string, expected: number): number[][] {
  const data = safeJson(text);
  if (data && (data as any).error) throw new LlmError(errorMessage((data as any).error));
  const rows = (data as any)?.data;
  if (!Array.isArray(rows)) throw new LlmError("No embeddings found in response.");
  // Providers return the rows in request order but also carry an explicit
  // `index`; honor it when present so a reordered response still lines up.
  const ordered = rows.every((r: any) => typeof r?.index === "number")
    ? [...rows].sort((a: any, b: any) => a.index - b.index)
    : rows;
  const vecs = ordered.map((r: any) => {
    if (!Array.isArray(r?.embedding)) throw new LlmError("No embedding found in a response row.");
    return (r.embedding as unknown[]).map(Number);
  });
  return vecs;
}

function extractEmbedding(text: string): number[] {
  const data = safeJson(text);
  if (data && (data as any).error) throw new LlmError(errorMessage((data as any).error));
  // OpenAI: {data:[{embedding:[...]}]} ; Ollama: {embedding:[...]} or {embeddings:[[...]]}
  const v =
    (data as any)?.data?.[0]?.embedding ??
    (data as any)?.embedding ??
    (data as any)?.embeddings?.[0];
  if (!Array.isArray(v)) throw new LlmError("No embedding found in response.");
  return v.map(Number);
}

// ---- request building -------------------------------------------------------

export function buildChatBody(
  spec: ProviderSpec,
  model: string,
  prompt: string,
  system?: string
): Record<string, unknown> {
  const messages: Array<{ role: string; content: string }> = [];
  messages.push({ role: "system", content: system ?? DEFAULT_SYSTEM_PROMPT });
  messages.push({ role: "user", content: prompt });
  const body: Record<string, unknown> = { model, messages };
  if (spec.style === "ollama") body.stream = false;
  return body;
}

export function directHeaders(spec: ProviderSpec, apiKey?: string): Record<string, string> {
  const h: Record<string, string> = { "Content-Type": "application/json" };
  if (apiKey) h["Authorization"] = `Bearer ${apiKey}`;
  if (spec.id === "openrouter") {
    h["HTTP-Referer"] = "https://excel-llm-addin";
    h["X-Title"] = "Excel LLM Add-in";
  }
  return h;
}

// ---- response parsing (native JSON — no hand-rolled scanning) ----------------

export function extractChatContent(text: string): string {
  const data = safeJson(text);
  if (data && (data as any).error) {
    throw new LlmError(errorMessage((data as any).error));
  }
  const choice = (data as any)?.choices?.[0];
  if (choice) {
    if (choice.message?.content != null) return String(choice.message.content);
    if (choice.text != null) return String(choice.text);
  }
  if ((data as any)?.message?.content != null) {
    return String((data as any).message.content);
  }
  throw new LlmError("No content found in response.");
}

export function extractModelList(spec: ProviderSpec, text: string): string[] {
  const data = safeJson(text);
  if (data && (data as any).error) {
    throw new LlmError(errorMessage((data as any).error));
  }
  if (spec.style === "ollama") {
    const models = (data as any)?.models;
    if (Array.isArray(models)) return models.map((m: any) => String(m?.name)).filter(Boolean);
  } else {
    // OpenAI shape is {data:[{id}]}; some providers (e.g. Together AI) return a
    // bare array of model objects instead. Accept either, and read id or name.
    const rows = Array.isArray(data) ? data : (data as any)?.data;
    if (Array.isArray(rows)) {
      return rows.map((m: any) => String(m?.id ?? m?.name ?? "")).filter((s) => s && s !== "undefined");
    }
  }
  return [];
}

// ---- proxy ------------------------------------------------------------------

async function callProxy(
  op: "chat" | "models" | "embed",
  prompt: string,
  settings: LlmSettings,
  spec: ProviderSpec,
  deps: Deps,
  extra?: Record<string, unknown>
): Promise<any> {
  const resp = await deps.fetch(settings.proxyUrl!, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      op,
      provider: spec.id,
      model: settings.model,
      prompt,
      system: settings.systemPrompt,
      baseUrl: settings.baseUrl,
      ...extra,
    }),
  });
  const text = await resp.text();
  const data = safeJson(text) ?? {};
  if (!resp.ok || (data as any).error) {
    throw new LlmError((data as any).error ? errorMessage((data as any).error) : `Proxy HTTP ${resp.status}`);
  }
  return data;
}

// ---- helpers ----------------------------------------------------------------

function requireProvider(id: string): ProviderSpec {
  const spec = getProvider(id);
  if (!spec) throw new LlmError(`Unknown provider '${id}'.`);
  return spec;
}

function safeJson(text: string): unknown {
  try {
    return JSON.parse(text);
  } catch {
    return null;
  }
}

function parseErrorMessage(text: string): string | undefined {
  const data = safeJson(text);
  if (data && (data as any).error) return errorMessage((data as any).error);
  return undefined;
}

function errorMessage(err: unknown): string {
  if (typeof err === "string") return err;
  if (!err || typeof err !== "object") return "Unknown error";
  const e = err as any;
  let msg = typeof e.message === "string" ? e.message : JSON.stringify(err);
  // Gateways like OpenRouter wrap the real upstream reason in error.metadata
  // (e.g. "Provider returned error" with the actual cause in metadata.raw).
  const meta = e.metadata;
  if (meta) {
    const raw = meta.raw;
    const detail = typeof raw === "string" ? raw : raw ? JSON.stringify(raw) : undefined;
    if (detail && !msg.includes(detail)) msg += ` — ${detail}`;
    else if (typeof meta.provider_name === "string") msg += ` (via ${meta.provider_name})`;
  }
  if (e.code != null && !msg.includes(String(e.code))) msg += ` [${e.code}]`;
  return msg;
}
