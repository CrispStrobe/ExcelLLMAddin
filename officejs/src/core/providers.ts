// Provider catalog. The data lives in shared/providers.json (single source of
// truth across the TS, proxy, and VBA editions) and is generated into
// providers.generated.ts by `npm run gen:providers`; the request "styles" and
// endpoint helpers below stay hand-written here.

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

export { PROVIDERS } from "./providers.generated";
import { PROVIDERS } from "./providers.generated";

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
