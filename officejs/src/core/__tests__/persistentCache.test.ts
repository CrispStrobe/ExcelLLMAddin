import { createPersistentCache, AsyncKeyValueStore } from "../persistentCache";

/** In-memory store with call spying and injectable failures. */
function memStore(seed: Record<string, string> = {}) {
  const data: Record<string, string> = { ...seed };
  let failSet = false;
  const store: AsyncKeyValueStore = {
    getItem: async (k) => (k in data ? data[k] : null),
    setItem: async (k, v) => {
      if (failSet) throw new Error("quota exceeded");
      data[k] = v;
    },
  };
  return { store, data, fail: () => (failSet = true) };
}

// Run scheduled flushes immediately so behavior is deterministic in tests.
const now = (fn: () => void) => fn();
const KEY = "excelllm.cache.v1";

describe("createPersistentCache", () => {
  test("get/set behave like a cache in memory", async () => {
    const { store } = memStore();
    const c = createPersistentCache(store, { schedule: now });
    expect(c.get("a")).toBeUndefined();
    c.set("a", "1");
    expect(c.get("a")).toBe("1");
  });

  test("persists small values to the store", async () => {
    const { store, data } = memStore();
    const c = createPersistentCache(store, { schedule: now });
    c.set("k", "hello");
    await c.flush();
    expect(JSON.parse(data[KEY])).toEqual([["k", "hello"]]);
  });

  test("hydrates prior entries from the store on startup", async () => {
    const { store } = memStore({ [KEY]: JSON.stringify([["prev", "value"]]) });
    const c = createPersistentCache(store, { schedule: now });
    await c.ready;
    expect(c.get("prev")).toBe("value");
  });

  test("keeps large values in memory but does not persist them", async () => {
    const { store, data } = memStore();
    const c = createPersistentCache(store, { schedule: now, maxValueBytes: 8 });
    const big = "x".repeat(50); // e.g. a serialized embedding
    c.set("small", "ok");
    c.set("big", big);
    await c.flush();
    expect(c.get("big")).toBe(big); // served from memory
    expect(JSON.parse(data[KEY])).toEqual([["small", "ok"]]); // big not persisted
  });

  test("evicts the oldest entry past max (LRU)", async () => {
    const { store } = memStore();
    const c = createPersistentCache(store, { schedule: now, max: 2 });
    c.set("a", "1");
    c.set("b", "2");
    c.get("a"); // refresh a's recency, so b is now oldest
    c.set("c", "3"); // evicts b
    expect(c.get("a")).toBe("1");
    expect(c.get("c")).toBe("3");
    expect(c.get("b")).toBeUndefined();
  });

  test("a failing setItem does not throw (degrades to in-memory)", async () => {
    const { store, fail } = memStore();
    const c = createPersistentCache(store, { schedule: now });
    fail();
    c.set("k", "v");
    await expect(c.flush()).resolves.toBeUndefined();
    expect(c.get("k")).toBe("v");
  });

  test("ignores corrupt stored data", async () => {
    const { store } = memStore({ [KEY]: "not json{" });
    const c = createPersistentCache(store, { schedule: now });
    await expect(c.ready).resolves.toBeUndefined();
    expect(c.get("anything")).toBeUndefined();
  });

  test("does not clobber a value written before hydration finishes", async () => {
    // Slow getItem so set() runs first; hydration must not overwrite the fresh value.
    const store: AsyncKeyValueStore = {
      getItem: async (_k) => {
        await Promise.resolve();
        return JSON.stringify([["k", "stale"]]);
      },
      setItem: async () => {},
    };
    const c = createPersistentCache(store, { schedule: now });
    c.set("k", "fresh");
    await c.ready;
    expect(c.get("k")).toBe("fresh");
  });

  test("clear empties memory and persists the empty set", async () => {
    const { store, data } = memStore();
    const c = createPersistentCache(store, { schedule: now });
    c.set("k", "v");
    await c.flush();
    c.clear();
    await c.flush();
    expect(c.get("k")).toBeUndefined();
    expect(JSON.parse(data[KEY])).toEqual([]);
  });
});
