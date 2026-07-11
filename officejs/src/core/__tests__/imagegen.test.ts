// Tests for the BFL image-generation protocol: submit -> poll -> Ready, plus the
// proxy path and error/timeout handling. sleep is injected (instant) and fetch is
// scripted, so there is no network and no real timer.

import { generateImage, ImageDeps, ImageSettings } from "../imagegen";
import { FetchLike } from "../llm";

const instant = { sleep: async () => {} };

/** Scripted fetch: each queued entry is {ok?, status?, body}. Records requests. */
function scripted(steps: Array<{ ok?: boolean; status?: number; body: any }>): {
  deps: ImageDeps;
  calls: Array<{ url: string; init: any }>;
} {
  const calls: Array<{ url: string; init: any }> = [];
  let i = 0;
  const fetch: FetchLike = async (url, init) => {
    calls.push({ url, init });
    const s = steps[Math.min(i++, steps.length - 1)];
    return { ok: s.ok ?? true, status: s.status ?? 200, text: async () => JSON.stringify(s.body) };
  };
  return { deps: { fetch, ...instant }, calls };
}

const bfl: ImageSettings = { apiKey: "bfl-key", model: "flux-dev", pollMs: 1, maxPolls: 5 };

describe("generateImage (BFL direct)", () => {
  test("submits, polls past Pending, and returns the image url", async () => {
    const { deps, calls } = scripted([
      { body: { id: "1", polling_url: "https://poll/x" } },
      { body: { status: "Pending" } },
      { body: { status: "Ready", result: { sample: "https://img/red.png" } } },
    ]);
    const url = await generateImage("a red square", bfl, deps);
    expect(url).toBe("https://img/red.png");
    // First call is the submit with the x-key header and the prompt.
    expect(calls[0].url).toBe("https://api.bfl.ai/v1/flux-dev");
    expect(calls[0].init.headers["x-key"]).toBe("bfl-key");
    expect(JSON.parse(calls[0].init.body).prompt).toBe("a red square");
    // Later calls poll the polling_url.
    expect(calls[1].url).toBe("https://poll/x");
  });

  test("throws on a moderated result", async () => {
    const { deps } = scripted([
      { body: { polling_url: "https://poll/x" } },
      { body: { status: "Content Moderated" } },
    ]);
    await expect(generateImage("x", bfl, deps)).rejects.toThrow(/moderated/i);
  });

  test("times out if never Ready", async () => {
    const { deps } = scripted([
      { body: { polling_url: "https://poll/x" } },
      { body: { status: "Pending" } },
    ]);
    await expect(generateImage("x", { ...bfl, maxPolls: 3 }, deps)).rejects.toThrow(/timed out/i);
  });

  test("surfaces a submit error (e.g. bad key)", async () => {
    const { deps } = scripted([{ ok: false, status: 401, body: { detail: "invalid key" } }]);
    await expect(generateImage("x", bfl, deps)).rejects.toThrow("invalid key");
  });

  test("requires an api key when not using a proxy", async () => {
    const { deps, calls } = scripted([{ body: {} }]);
    await expect(generateImage("x", { model: "flux-dev" }, deps)).rejects.toThrow(/BFL API key/);
    expect(calls.length).toBe(0);
  });

  test("rejects an empty prompt before any call", async () => {
    const { deps, calls } = scripted([{ body: {} }]);
    await expect(generateImage("   ", bfl, deps)).rejects.toThrow(/No prompt/);
    expect(calls.length).toBe(0);
  });
});

describe("generateImage (via proxy)", () => {
  test("posts an image envelope and returns the url", async () => {
    const { deps, calls } = scripted([{ body: { url: "https://img/fromproxy.png" } }]);
    const s: ImageSettings = { proxyUrl: "https://proxy.example/api", width: 512, height: 512 };
    const url = await generateImage("a cat", s, deps);
    expect(url).toBe("https://img/fromproxy.png");
    const env = JSON.parse(calls[0].init.body);
    expect(env).toMatchObject({ op: "image", prompt: "a cat", width: 512, height: 512 });
  });

  test("throws when the proxy returns no url", async () => {
    const { deps } = scripted([{ body: { error: "no image backend configured" } }]);
    await expect(generateImage("x", { proxyUrl: "https://p/api" }, deps)).rejects.toThrow(/no image backend/);
  });
});
