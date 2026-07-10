import { createLruCache } from "../cache";

describe("createLruCache", () => {
  test("stores and retrieves", () => {
    const c = createLruCache();
    expect(c.get("a")).toBeUndefined();
    c.set("a", "1");
    expect(c.get("a")).toBe("1");
  });

  test("evicts the oldest beyond max", () => {
    const c = createLruCache(2);
    c.set("a", "1");
    c.set("b", "2");
    c.set("c", "3"); // evicts "a"
    expect(c.get("a")).toBeUndefined();
    expect(c.get("b")).toBe("2");
    expect(c.get("c")).toBe("3");
  });

  test("get refreshes recency so the touched key survives", () => {
    const c = createLruCache(2);
    c.set("a", "1");
    c.set("b", "2");
    c.get("a"); // "a" is now most-recent
    c.set("c", "3"); // evicts "b", not "a"
    expect(c.get("b")).toBeUndefined();
    expect(c.get("a")).toBe("1");
    expect(c.get("c")).toBe("3");
  });

  test("clear empties the cache", () => {
    const c = createLruCache();
    c.set("a", "1");
    c.clear?.();
    expect(c.get("a")).toBeUndefined();
  });
});
