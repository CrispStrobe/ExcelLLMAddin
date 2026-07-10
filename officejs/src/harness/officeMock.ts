// Minimal mock of the Office bits the task pane uses, so it can run in a plain
// browser (Safari/Chrome) with no Excel. Storage is backed by localStorage so
// settings persist across reloads. Loaded ONLY by the dev harness.

/* global localStorage */

const STORE_KEY = "__harness_office_storage";

function readAll(): Record<string, string> {
  try {
    return JSON.parse(localStorage.getItem(STORE_KEY) || "{}");
  } catch {
    return {};
  }
}
function writeAll(data: Record<string, string>): void {
  try {
    localStorage.setItem(STORE_KEY, JSON.stringify(data));
  } catch {
    /* ignore */
  }
}

const g = globalThis as any;

g.OfficeRuntime = {
  storage: {
    getItem: async (key: string) => {
      const d = readAll();
      return key in d ? d[key] : null;
    },
    setItem: async (key: string, value: string) => {
      const d = readAll();
      d[key] = value;
      writeAll(d);
    },
    removeItem: async (key: string) => {
      const d = readAll();
      delete d[key];
      writeAll(d);
    },
    getItems: async (keys: string[]) => {
      const d = readAll();
      const out: Record<string, string | null> = {};
      for (const k of keys) out[k] = k in d ? d[k] : null;
      return out;
    },
    setItems: async (items: Record<string, string>) => {
      writeAll({ ...readAll(), ...items });
    },
    removeItems: async (keys: string[]) => {
      const d = readAll();
      for (const k of keys) delete d[k];
      writeAll(d);
    },
  },
};

g.Office = {
  HostType: { Excel: "Excel" },
  onReady: (cb?: (info: { host: string; platform: string }) => void) => {
    const info = { host: "Excel", platform: "harness" };
    if (cb) Promise.resolve().then(() => cb(info));
    return Promise.resolve(info);
  },
};
