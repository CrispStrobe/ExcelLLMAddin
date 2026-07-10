// Pure incremental parser for streaming chat responses. Feed it raw text chunks
// (as they arrive off the network) and it returns the new content to append.
// Handles OpenAI-style SSE ("data: {json}\n\n" with choices[0].delta.content)
// and Ollama-style NDJSON (one JSON object per line with message.content).
// No network/DOM — fully unit-testable.

import { ProviderStyle } from "./providers";

export interface StreamParser {
  /** Feed a raw text chunk; returns the newly-decoded content (may be ""). */
  push(chunk: string): string;
}

export function createStreamParser(style: ProviderStyle): StreamParser {
  let buffer = "";

  return {
    push(chunk: string): string {
      buffer += chunk;
      const lines = buffer.split("\n");
      buffer = lines.pop() ?? ""; // keep the trailing incomplete line

      let out = "";
      for (const raw of lines) {
        const line = raw.trim();
        if (line === "") continue;

        if (style === "ollama") {
          out += contentFrom(line, "ollama");
        } else {
          if (!line.startsWith("data:")) continue;
          const payload = line.slice(5).trim();
          if (payload === "[DONE]") continue;
          out += contentFrom(payload, "openai");
        }
      }
      return out;
    },
  };
}

function contentFrom(jsonText: string, style: ProviderStyle): string {
  try {
    const j = JSON.parse(jsonText) as any;
    if (style === "ollama") return String(j?.message?.content ?? "");
    return String(j?.choices?.[0]?.delta?.content ?? j?.choices?.[0]?.message?.content ?? "");
  } catch {
    return "";
  }
}
