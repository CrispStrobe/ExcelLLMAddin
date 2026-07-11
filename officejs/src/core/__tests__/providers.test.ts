import {
  getProvider,
  chatEndpoint,
  modelsEndpoint,
  PROVIDERS,
} from "../providers";

describe("getProvider", () => {
  test("is case- and whitespace-insensitive", () => {
    expect(getProvider("  OpenAI ")?.id).toBe("openai");
    expect(getProvider("OLLAMA")?.id).toBe("ollama");
  });

  test("returns undefined for unknown", () => {
    expect(getProvider("banana")).toBeUndefined();
    expect(getProvider("")).toBeUndefined();
  });
});

describe("endpoints", () => {
  test("openai style", () => {
    const p = PROVIDERS.openai;
    expect(chatEndpoint(p, "https://api.openai.com/v1")).toBe("https://api.openai.com/v1/chat/completions");
    expect(modelsEndpoint(p, "https://api.openai.com/v1")).toBe("https://api.openai.com/v1/models");
  });

  test("ollama style", () => {
    const p = PROVIDERS.ollama;
    expect(chatEndpoint(p, "http://localhost:11434")).toBe("http://localhost:11434/api/chat");
    expect(modelsEndpoint(p, "http://localhost:11434")).toBe("http://localhost:11434/api/tags");
  });

  test("tolerates a trailing slash on the base url", () => {
    expect(chatEndpoint(PROVIDERS.openai, "https://x.ai/v1/")).toBe("https://x.ai/v1/chat/completions");
  });

  test("OpenAI-compat cloud providers use openai-style paths", () => {
    for (const id of ["groq", "together", "cerebras", "gemini", "cohere", "huggingface", "requesty"]) {
      const p = PROVIDERS[id];
      expect(p.style).toBe("openai");
      expect(chatEndpoint(p, p.defaultBaseUrl)).toBe(`${p.defaultBaseUrl}/chat/completions`);
      expect(modelsEndpoint(p, p.defaultBaseUrl)).toBe(`${p.defaultBaseUrl}/models`);
    }
  });
});

describe("catalog invariants", () => {
  test("every provider that requires a key is openai-style, ollama needs none", () => {
    expect(PROVIDERS.ollama.requiresKey).toBe(false);
    expect(PROVIDERS.openai.requiresKey).toBe(true);
  });

  test("keys of the record match each spec id", () => {
    for (const [key, spec] of Object.entries(PROVIDERS)) {
      expect(spec.id).toBe(key);
    }
  });

  test("cloud providers that block the browser are marked not browserFriendly", () => {
    // These need the proxy; the UI keys its CORS hint off this flag.
    for (const id of ["groq", "together", "cerebras", "gemini", "cohere", "huggingface", "requesty"]) {
      expect(PROVIDERS[id].browserFriendly).toBe(false);
      expect(PROVIDERS[id].requiresKey).toBe(true);
    }
  });
});
