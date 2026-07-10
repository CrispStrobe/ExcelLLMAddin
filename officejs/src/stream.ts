// Streaming chat driver (browser/runtime). Uses the real fetch + a ReadableStream
// reader, so it lives outside the pure core. Falls back to a single non-streaming
// read if the runtime/provider can't stream, so callers never break.

import { getProvider, chatEndpoint } from "./core/providers";
import { LlmSettings, LlmError, buildChatBody, directHeaders } from "./core/llm";
import { createStreamParser } from "./core/streamParser";

/* global fetch, TextDecoder, ReadableStream */

export async function streamChat(
  promptText: string,
  settings: LlmSettings,
  onDelta: (partial: string) => void
): Promise<string> {
  const spec = getProvider(settings.provider);
  if (!spec) throw new LlmError(`Unknown provider '${settings.provider}'.`);
  const baseUrl = settings.baseUrl || spec.defaultBaseUrl;
  if (spec.requiresKey && !settings.apiKey) {
    throw new LlmError(`No API key configured for ${spec.label}.`);
  }

  const url = chatEndpoint(spec, baseUrl);
  const body = { ...buildChatBody(spec, settings.model, promptText, settings.systemPrompt), stream: true };

  const resp = await fetch(url, {
    method: "POST",
    headers: directHeaders(spec, settings.apiKey),
    body: JSON.stringify(body),
  });

  if (!resp.ok) {
    throw new LlmError(errorFrom(await resp.text()) ?? `HTTP ${resp.status}`);
  }

  // Fallback: runtime/provider didn't give a readable stream — read it all at once.
  const stream = resp.body as ReadableStream<Uint8Array> | null;
  if (!stream || typeof stream.getReader !== "function") {
    const full = createStreamParser(spec.style).push((await resp.text()) + "\n");
    onDelta(full);
    return full;
  }

  const reader = stream.getReader();
  const decoder = new TextDecoder();
  const parser = createStreamParser(spec.style);
  let full = "";
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    const delta = parser.push(decoder.decode(value, { stream: true }));
    if (delta) {
      full += delta;
      onDelta(full);
    }
  }
  return full;
}

function errorFrom(text: string): string | undefined {
  try {
    const j = JSON.parse(text) as any;
    if (typeof j.error === "string") return j.error;
    if (j.error && typeof j.error.message === "string") return j.error.message;
  } catch {
    /* ignore */
  }
  return undefined;
}
