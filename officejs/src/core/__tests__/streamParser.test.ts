import { createStreamParser } from "../streamParser";

describe("createStreamParser (openai SSE)", () => {
  test("accumulates deltas across events", () => {
    const p = createStreamParser("openai");
    let out = p.push('data: {"choices":[{"delta":{"content":"Hel"}}]}\n');
    out += p.push('data: {"choices":[{"delta":{"content":"lo"}}]}\n\n');
    out += p.push("data: [DONE]\n\n");
    expect(out).toBe("Hello");
  });

  test("handles a chunk boundary in the middle of a line", () => {
    const p = createStreamParser("openai");
    let out = p.push('data: {"choices":[{"delta":{"content":"A'); // incomplete line, buffered
    out += p.push('BC"}}]}\n'); // completes it
    expect(out).toBe("ABC");
  });

  test("ignores blank lines and SSE comments", () => {
    const p = createStreamParser("openai");
    expect(p.push("\n\n: keep-alive\n")).toBe("");
  });

  test("falls back to message.content when a chunk has no delta", () => {
    // Some providers emit a whole message object in the final SSE event.
    const p = createStreamParser("openai");
    expect(p.push('data: {"choices":[{"message":{"content":"whole"}}]}\n')).toBe("whole");
  });

  test("skips a malformed data line without throwing", () => {
    const p = createStreamParser("openai");
    expect(p.push("data: {not valid json}\n")).toBe("");
  });
});

describe("createStreamParser (ollama NDJSON)", () => {
  test("accumulates message.content per line", () => {
    const p = createStreamParser("ollama");
    let out = p.push('{"message":{"content":"foo"}}\n');
    out += p.push('{"message":{"content":"bar"},"done":true}\n');
    expect(out).toBe("foobar");
  });
});
