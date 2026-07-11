// AUTO-GENERATED from shared/providers.json by tools/gen-providers.cjs.
// Do not edit by hand — run `npm run gen:providers` to regenerate.
import type { ProviderSpec } from "./providers";

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
