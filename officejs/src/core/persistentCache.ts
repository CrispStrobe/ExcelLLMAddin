// A response cache that survives custom-functions-runtime reloads. The plain LRU
// (cache.ts) lives only for a session, so identical prompts re-hit the API after
// every reload. This wraps an in-memory LRU with best-effort persistence to an
// async key/value store (OfficeRuntime.storage), so repeat prompts across
// sessions are served for free.
//
// Design constraints:
//   - The LlmCache interface is synchronous (get/set), and runPrompt calls it
//     synchronously, so the in-memory Map stays the source of truth. Persistence
//     is write-behind (debounced) and hydration is best-effort on startup.
//   - Large values (embedding vectors are thousands of floats) are kept in memory
//     but NOT persisted, so a RECALL over a big range can't blow the storage
//     quota. Small prompt/task results — the common case — persist fine.
//   - Any storage failure degrades silently to pure in-memory behavior.

import { LlmCache } from "./llm";

/** Minimal async key/value store. OfficeRuntime.storage satisfies this shape. */
export interface AsyncKeyValueStore {
  getItem(key: string): Promise<string | null>;
  setItem(key: string, value: string): Promise<void>;
}

export interface PersistentCacheOptions {
  storageKey?: string;
  /** Max in-memory entries (LRU eviction beyond this). */
  max?: number;
  /** Values longer than this are cached in memory but not persisted (skips embeddings). */
  maxValueBytes?: number;
  /** Debounce window before a write-behind flush. */
  debounceMs?: number;
  /** Injectable scheduler (defaults to setTimeout) so flush timing is testable. */
  schedule?: (fn: () => void, ms: number) => void;
}

export interface PersistentCache extends LlmCache {
  /** Always present here (LlmCache leaves it optional). */
  clear(): void;
  /** Resolves once stored entries have been merged into memory. */
  ready: Promise<void>;
  /** Force an immediate write of the current entries. */
  flush(): Promise<void>;
}

const DEFAULTS = { storageKey: "excelllm.cache.v1", max: 500, maxValueBytes: 8192, debounceMs: 1500 };

export function createPersistentCache(
  store: AsyncKeyValueStore,
  options: PersistentCacheOptions = {}
): PersistentCache {
  const opt = { ...DEFAULTS, ...options };
  const schedule = options.schedule ?? ((fn: () => void, ms: number) => setTimeout(fn, ms));
  const map = new Map<string, string>(); // insertion order = LRU (oldest first)
  let dirty = false;
  let flushQueued = false;

  function touch(key: string, value: string): void {
    if (map.has(key)) map.delete(key);
    map.set(key, value);
    while (map.size > opt.max) {
      const oldest = map.keys().next().value;
      if (oldest === undefined) break;
      map.delete(oldest);
    }
  }

  async function doFlush(): Promise<void> {
    flushQueued = false;
    if (!dirty) return;
    dirty = false;
    // Persist in recency order, skipping oversized values to bound storage size.
    const entries: Array<[string, string]> = [];
    for (const [k, v] of map) {
      if (v.length <= opt.maxValueBytes) entries.push([k, v]);
    }
    try {
      await store.setItem(opt.storageKey, JSON.stringify(entries));
    } catch {
      /* storage unavailable — remain in-memory only */
    }
  }

  function scheduleFlush(): void {
    dirty = true;
    if (flushQueued) return;
    flushQueued = true;
    schedule(() => void doFlush(), opt.debounceMs);
  }

  const ready = (async () => {
    try {
      const raw = await store.getItem(opt.storageKey);
      if (!raw) return;
      const entries = JSON.parse(raw);
      if (!Array.isArray(entries)) return;
      for (const e of entries) {
        // Don't clobber anything written while we were hydrating.
        if (Array.isArray(e) && typeof e[0] === "string" && typeof e[1] === "string" && !map.has(e[0])) {
          touch(e[0], e[1]);
        }
      }
    } catch {
      /* absent or corrupt cache — start empty */
    }
  })();

  return {
    ready,
    get(key: string): string | undefined {
      const v = map.get(key);
      if (v !== undefined) {
        map.delete(key);
        map.set(key, v); // refresh recency
      }
      return v;
    },
    set(key: string, value: string): void {
      touch(key, value);
      scheduleFlush();
    },
    clear(): void {
      map.clear();
      dirty = true;
      void doFlush();
    },
    flush(): Promise<void> {
      return doFlush();
    },
  };
}
