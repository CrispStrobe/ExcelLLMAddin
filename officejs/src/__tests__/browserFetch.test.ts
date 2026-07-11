// Tests for the browser fetch adapter: response pass-through, the abort→friendly
// timeout mapping (the reason it exists — an unreachable Ollama shouldn't hang the
// UI for the OS default), and that non-abort errors propagate unchanged.

import { browserFetch } from "../browserFetch";

describe("browserFetch", () => {
  afterEach(() => {
    jest.useRealTimers();
  });

  test("passes through ok, status, and text", async () => {
    (global as any).fetch = async () => ({ ok: true, status: 201, text: async () => "hi" });
    const r = await browserFetch("http://x", { method: "GET" } as any);
    expect(r.ok).toBe(true);
    expect(r.status).toBe(201);
    expect(await r.text()).toBe("hi");
  });

  test("maps an abort to a friendly timeout message", async () => {
    jest.useFakeTimers();
    (global as any).fetch = (_url: string, init: any) =>
      new Promise((_resolve, reject) => {
        init.signal.addEventListener("abort", () => {
          const e = new Error("aborted");
          e.name = "AbortError";
          reject(e);
        });
      });
    const p = browserFetch("http://x", {} as any);
    jest.advanceTimersByTime(45000);
    await expect(p).rejects.toThrow(/timed out after 45s/);
  });

  test("rethrows non-abort errors unchanged", async () => {
    (global as any).fetch = async () => {
      throw new Error("ECONNREFUSED");
    };
    await expect(browserFetch("http://x", {} as any)).rejects.toThrow("ECONNREFUSED");
  });
});
