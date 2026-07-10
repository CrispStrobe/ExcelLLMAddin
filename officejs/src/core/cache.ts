import { LlmCache } from "./llm";

/**
 * Small LRU cache backed by a Map (Map preserves insertion order, so the first
 * key is the oldest). Bounds memory in the long-lived custom-functions runtime.
 */
export function createLruCache(max = 500): LlmCache {
  const map = new Map<string, string>();
  return {
    get(key: string): string | undefined {
      const v = map.get(key);
      if (v !== undefined) {
        map.delete(key);
        map.set(key, v); // refresh recency
      }
      return v;
    },
    set(key: string, value: string): void {
      if (map.has(key)) map.delete(key);
      map.set(key, value);
      if (map.size > max) {
        const oldest = map.keys().next().value;
        if (oldest !== undefined) map.delete(oldest);
      }
    },
    clear(): void {
      map.clear();
    },
  };
}
