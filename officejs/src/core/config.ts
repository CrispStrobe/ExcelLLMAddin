// Settings persistence. OfficeRuntime.storage is shared across the add-in's
// runtimes (task pane writes, custom-functions runtime reads), and persists
// per-user. Kept separate from llm.ts so the core stays Office-free + testable.

import { LlmSettings } from "./llm";

const STORAGE_KEY = "excelllm.settings.v1";

export function defaultSettings(): LlmSettings {
  return {
    provider: "ollama",
    model: "llama3.2",
    baseUrl: "",
    apiKey: "",
    proxyUrl: "",
    systemPrompt: "",
    embedModel: "",
  };
}

export async function loadSettings(): Promise<LlmSettings> {
  try {
    const raw = await OfficeRuntime.storage.getItem(STORAGE_KEY);
    if (raw) return { ...defaultSettings(), ...JSON.parse(raw) };
  } catch {
    /* fall through to defaults */
  }
  return defaultSettings();
}

export async function saveSettings(settings: LlmSettings): Promise<void> {
  await OfficeRuntime.storage.setItem(STORAGE_KEY, JSON.stringify(settings));
}
