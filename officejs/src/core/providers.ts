// Provider catalog. Mirrors the VBA add-in's provider list but expresses the
// two request "styles" explicitly so the client stays declarative.

export type ProviderStyle = "openai" | "ollama";

export interface ProviderSpec {
  id: string;
  label: string;
  defaultBaseUrl: string;
  requiresKey: boolean;
  style: ProviderStyle;
  /** True if the provider's API is callable directly from a browser (CORS). */
  browserFriendly: boolean;
}

export const PROVIDERS: Record<string, ProviderSpec> = {
  openai: {
    id: "openai", label: "OpenAI",
    defaultBaseUrl: "https://api.openai.com/v1",
    requiresKey: true, style: "openai", browserFriendly: false,
  },
  mistral: {
    id: "mistral", label: "Mistral",
    defaultBaseUrl: "https://api.mistral.ai/v1",
    requiresKey: true, style: "openai", browserFriendly: false,
  },
  nebius: {
    id: "nebius", label: "Nebius",
    defaultBaseUrl: "https://api.studio.nebius.com/v1",
    // Verified: Nebius returns access-control-allow-origin:* — works direct from the browser.
    requiresKey: true, style: "openai", browserFriendly: true,
  },
  scaleway: {
    id: "scaleway", label: "Scaleway",
    defaultBaseUrl: "https://api.scaleway.ai/v1",
    requiresKey: true, style: "openai", browserFriendly: false,
  },
  openrouter: {
    id: "openrouter", label: "OpenRouter",
    defaultBaseUrl: "https://openrouter.ai/api/v1",
    requiresKey: true, style: "openai", browserFriendly: true,
  },
  groq: {
    id: "groq", label: "Groq",
    // OpenAI-compatible; very fast. Bearer auth. No embeddings endpoint.
    defaultBaseUrl: "https://api.groq.com/openai/v1",
    requiresKey: true, style: "openai", browserFriendly: false,
  },
  together: {
    id: "together", label: "Together AI",
    defaultBaseUrl: "https://api.together.xyz/v1",
    requiresKey: true, style: "openai", browserFriendly: false,
  },
  cerebras: {
    id: "cerebras", label: "Cerebras",
    defaultBaseUrl: "https://api.cerebras.ai/v1",
    requiresKey: true, style: "openai", browserFriendly: false,
  },
  gemini: {
    id: "gemini", label: "Google Gemini",
    // Google's OpenAI-compatible endpoint: /chat/completions, /models, /embeddings
    // all normalized to OpenAI shapes, Bearer auth. CORS-blocked → proxy.
    defaultBaseUrl: "https://generativelanguage.googleapis.com/v1beta/openai",
    requiresKey: true, style: "openai", browserFriendly: false,
  },
  cohere: {
    id: "cohere", label: "Cohere",
    defaultBaseUrl: "https://api.cohere.ai/compatibility/v1",
    requiresKey: true, style: "openai", browserFriendly: false,
  },
  huggingface: {
    id: "huggingface", label: "Hugging Face",
    // Inference-provider router; fronts many models with one OpenAI-compat API.
    defaultBaseUrl: "https://router.huggingface.co/v1",
    requiresKey: true, style: "openai", browserFriendly: false,
  },
  requesty: {
    id: "requesty", label: "Requesty",
    defaultBaseUrl: "https://router.requesty.ai/v1",
    requiresKey: true, style: "openai", browserFriendly: false,
  },
  ollama: {
    id: "ollama", label: "Ollama (local)",
    defaultBaseUrl: "http://localhost:11434",
    requiresKey: false, style: "ollama", browserFriendly: true,
  },
};

export function getProvider(id: string): ProviderSpec | undefined {
  return PROVIDERS[(id || "").trim().toLowerCase()];
}

function trimTrailingSlash(u: string): string {
  return u.replace(/\/+$/, "");
}

export function chatEndpoint(spec: ProviderSpec, baseUrl: string): string {
  const base = trimTrailingSlash(baseUrl);
  return spec.style === "ollama" ? `${base}/api/chat` : `${base}/chat/completions`;
}

export function modelsEndpoint(spec: ProviderSpec, baseUrl: string): string {
  const base = trimTrailingSlash(baseUrl);
  return spec.style === "ollama" ? `${base}/api/tags` : `${base}/models`;
}

export function embeddingsEndpoint(spec: ProviderSpec, baseUrl: string): string {
  const base = trimTrailingSlash(baseUrl);
  return spec.style === "ollama" ? `${base}/api/embeddings` : `${base}/embeddings`;
}
