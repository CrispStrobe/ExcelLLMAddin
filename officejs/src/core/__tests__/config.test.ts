// Tests for settings persistence over OfficeRuntime.storage. We install a fake
// storage backed by a plain object, covering: defaults, partial-over-default
// merge (so a settings blob written by an older/newer build never drops keys),
// resilience to corrupt or unreadable storage, and save→load round-trip.

import { defaultSettings, loadSettings, saveSettings } from "../config";

const KEY = "excelllm.settings.v1";

function installStorage(store: Record<string, string>, opts: { throwOnGet?: boolean } = {}) {
  (global as any).OfficeRuntime = {
    storage: {
      getItem: async (k: string) => {
        if (opts.throwOnGet) throw new Error("storage unavailable");
        return k in store ? store[k] : null;
      },
      setItem: async (k: string, v: string) => {
        store[k] = v;
      },
    },
  };
  return store;
}

describe("settings persistence", () => {
  test("defaultSettings ships an offline-first (Ollama) profile", () => {
    const d = defaultSettings();
    expect(d.provider).toBe("ollama");
    expect(d).toHaveProperty("apiKey", "");
    expect(d).toHaveProperty("mcpUrl", "");
  });

  test("loadSettings returns defaults when nothing is stored", async () => {
    installStorage({});
    expect(await loadSettings()).toEqual(defaultSettings());
  });

  test("loadSettings merges a partial blob over the defaults", async () => {
    installStorage({ [KEY]: JSON.stringify({ provider: "openai", model: "gpt-4o-mini" }) });
    const s = await loadSettings();
    expect(s.provider).toBe("openai");
    expect(s.model).toBe("gpt-4o-mini");
    expect(s.mcpUrl).toBe(""); // key absent from the blob still present from defaults
  });

  test("loadSettings falls back to defaults on corrupt JSON", async () => {
    installStorage({ [KEY]: "{not valid json" });
    expect(await loadSettings()).toEqual(defaultSettings());
  });

  test("loadSettings falls back to defaults when storage throws", async () => {
    installStorage({}, { throwOnGet: true });
    expect(await loadSettings()).toEqual(defaultSettings());
  });

  test("saveSettings then loadSettings round-trips", async () => {
    const store = installStorage({});
    const s = { ...defaultSettings(), provider: "nebius", model: "custom", apiKey: "k" };
    await saveSettings(s);
    expect(JSON.parse(store[KEY])).toEqual(s);
    expect(await loadSettings()).toEqual(s);
  });
});
