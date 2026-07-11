// Settings persistence. OfficeRuntime.storage is shared across the add-in's
// runtimes (task pane writes, custom-functions runtime reads), and persists
// per-user. Kept separate from llm.ts so the core stays Office-free + testable.

import { LlmSettings } from "./llm";

const STORAGE_KEY = "excelllm.settings.v1";

// Every custom-function cell resolves settings before it runs, so a sheet full of
// =LLM.PROMPT cells would hit OfficeRuntime.storage once per cell on every recalc.
// A short-lived cache collapses a bulk recalc to ~one storage read. The task pane
// and functions share one runtime, so saveSettings updates the cache in place and
// changes take effect immediately; the TTL is only a backstop for a write made by
// another runtime instance.
const SETTINGS_TTL_MS = 3000;
let settingsCache: { value: LlmSettings; at: number } | null = null;

/** Drop the cached settings (e.g. in tests, or to force a fresh read). */
export function clearSettingsCache(): void {
  settingsCache = null;
}

export function defaultSettings(): LlmSettings {
  return {
    provider: "ollama",
    model: "llama3.2",
    baseUrl: "",
    apiKey: "",
    proxyUrl: "",
    systemPrompt: "",
    embedModel: "",
    mcpUrl: "",
    imageApiKey: "",
    imageModel: "flux-dev",
  };
}

export async function loadSettings(): Promise<LlmSettings> {
  if (settingsCache && Date.now() - settingsCache.at < SETTINGS_TTL_MS) {
    return settingsCache.value;
  }
  try {
    const raw = await OfficeRuntime.storage.getItem(STORAGE_KEY);
    const value = raw ? { ...defaultSettings(), ...JSON.parse(raw) } : defaultSettings();
    settingsCache = { value, at: Date.now() };
    return value;
  } catch {
    // Don't cache the error-fallback, so storage recovers on the next call.
    return defaultSettings();
  }
}

export async function saveSettings(settings: LlmSettings): Promise<void> {
  await OfficeRuntime.storage.setItem(STORAGE_KEY, JSON.stringify(settings));
  settingsCache = { value: settings, at: Date.now() }; // keep the shared runtime consistent
}
