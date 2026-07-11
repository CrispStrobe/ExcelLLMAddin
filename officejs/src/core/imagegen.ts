// Image generation via Black Forest Labs (FLUX). BFL is an async submit-then-poll
// API (not OpenAI-compatible), so it lives outside the chat providers. Returns a
// hosted image URL — pair it with Excel's =IMAGE() to show the picture in a cell.
//
// fetch + sleep are injected, so the whole protocol is unit-testable with no
// network or timers. Direct calls are CORS-blocked in a browser, so the task pane
// routes through the proxy (op:"image"); Node/tests and the VBA client call direct.

import { FetchLike } from "./llm";

/* global setTimeout */

export interface ImageDeps {
  fetch: FetchLike;
  /** Injected for tests; defaults to a real timer. */
  sleep?: (ms: number) => Promise<void>;
}

export interface ImageSettings {
  /** BFL API key (direct calls). */
  apiKey?: string;
  /** BFL model slug, e.g. flux-dev, flux-pro-1.1. Default flux-dev. */
  model?: string;
  /** If set, POST {op:"image",...} here instead of calling BFL directly. */
  proxyUrl?: string;
  width?: number;
  height?: number;
  /** Poll attempts (default 30) and interval ms (default 1500). */
  maxPolls?: number;
  pollMs?: number;
}

const defaultSleep = (ms: number): Promise<void> => new Promise((r) => setTimeout(r, ms));

/** Generate an image and return its URL. Throws on error/timeout. */
export async function generateImage(prompt: string, settings: ImageSettings, deps: ImageDeps): Promise<string> {
  const sleep = deps.sleep ?? defaultSleep;
  if (!prompt || !prompt.trim()) throw new Error("No prompt provided for image generation.");

  if (settings.proxyUrl) {
    const resp = await deps.fetch(settings.proxyUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        op: "image",
        prompt,
        model: settings.model,
        width: settings.width,
        height: settings.height,
      }),
    });
    const data = (safeJson(await resp.text()) ?? {}) as any;
    if (!resp.ok || data.error) throw new Error(data.error ? String(data.error) : `Proxy HTTP ${resp.status}`);
    if (!data.url) throw new Error("Proxy returned no image url.");
    return String(data.url);
  }

  if (!settings.apiKey) throw new Error("No BFL API key configured for image generation.");
  const model = settings.model || "flux-dev";

  // 1) Submit the job.
  const submit = await deps.fetch(`https://api.bfl.ai/v1/${model}`, {
    method: "POST",
    headers: { "x-key": settings.apiKey, "Content-Type": "application/json" },
    body: JSON.stringify({ prompt, width: settings.width ?? 1024, height: settings.height ?? 768 }),
  });
  const subText = await submit.text();
  if (!submit.ok) throw new Error(bflError(subText) ?? `Image request failed: HTTP ${submit.status}`);
  const sub = (safeJson(subText) ?? {}) as any;
  const pollUrl: string | undefined = sub.polling_url;
  if (!pollUrl) throw new Error("BFL returned no polling_url.");

  // 2) Poll until the result is Ready.
  const maxPolls = settings.maxPolls ?? 30;
  const pollMs = settings.pollMs ?? 1500;
  for (let i = 0; i < maxPolls; i++) {
    await sleep(pollMs);
    const pr = await deps.fetch(pollUrl, { method: "GET", headers: { "x-key": settings.apiKey } });
    const pd = (safeJson(await pr.text()) ?? {}) as any;
    const status = String(pd.status ?? "");
    if (status === "Ready") {
      const url = pd.result?.sample;
      if (!url) throw new Error("BFL result contained no image.");
      return String(url);
    }
    if (status === "Error" || status === "Failed" || status === "Content Moderated" || status === "Request Moderated") {
      throw new Error(`Image generation ${status.toLowerCase()}.`);
    }
  }
  throw new Error(`Image generation timed out after ${maxPolls} polls.`);
}

function safeJson(text: string): unknown {
  try {
    return JSON.parse(text);
  } catch {
    return null;
  }
}

function bflError(text: string): string | undefined {
  const d = safeJson(text) as any;
  if (d && typeof d === "object") {
    if (typeof d.detail === "string") return d.detail;
    if (d.error && typeof d.error === "string") return d.error;
    if (d.error?.message) return String(d.error.message);
  }
  return undefined;
}
