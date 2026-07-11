// Tests for settings persistence over OfficeRuntime.storage. We install a fake
// storage backed by a plain object, covering: defaults, partial-over-default
// merge (so a settings blob written by an older/newer build never drops keys),
// resilience to corrupt or unreadable storage, and save→load round-trip.

import { defaultSettings, loadSettings, saveSettings, clearSettingsCache } from "../config";

const KEY = "excelllm.settings.v1";

let getCalls = 0;

function installStorage(store: Record<string, string>, opts: { throwOnGet?: boolean } = {}) {
  getCalls = 0;
  clearSettingsCache(); // a fresh store means a fresh read; don't serve a prior test's cache
  (global as any).OfficeRuntime = {
    storage: {
      getItem: async (k: string) => {
        getCalls++;
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

describe("settings read cache", () => {
  test("repeated loads within the TTL hit storage once (bulk-recalc path)", async () => {
    installStorage({ [KEY]: JSON.stringify({ provider: "openai", model: "gpt-4o-mini" }) });
    const a = await loadSettings();
    const b = await loadSettings();
    const c = await loadSettings();
    expect(a.provider).toBe("openai");
    expect(b).toEqual(a);
    expect(c).toEqual(a);
    expect(getCalls).toBe(1); // 200 cells → 1 storage read, not 200
  });

  test("saveSettings makes the new value visible without another storage read", async () => {
    installStorage({});
    await loadSettings(); // warm cache (1 read)
    await saveSettings({ ...defaultSettings(), provider: "mistral", model: "m" });
    const s = await loadSettings();
    expect(s.provider).toBe("mistral");
    expect(getCalls).toBe(1); // the save updated the cache in place; no re-read
  });

  test("clearSettingsCache forces a fresh read", async () => {
    installStorage({ [KEY]: JSON.stringify({ provider: "openai", model: "x" }) });
    await loadSettings();
    clearSettingsCache();
    await loadSettings();
    expect(getCalls).toBe(2);
  });

  test("the error-fallback is not cached (storage can recover next call)", async () => {
    const store = installStorage({}, { throwOnGet: true });
    expect(await loadSettings()).toEqual(defaultSettings()); // storage throws → defaults
    // Storage recovers; the next load must read it, not serve a cached default.
    (global as any).OfficeRuntime.storage.getItem = async (k: string) => {
      getCalls++;
      return k in store ? store[k] : null;
    };
    store[KEY] = JSON.stringify({ provider: "nebius", model: "n" });
    expect((await loadSettings()).provider).toBe("nebius");
  });
});
