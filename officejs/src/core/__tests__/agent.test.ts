import { runAgent, chatWithTools, createApprovalExecutor, ToolSchema } from "../agent";
import { LlmSettings, Deps, FetchLike } from "../llm";

/** Fetch double that returns queued response bodies in order. */
function queueMock(bodies: string[]): { deps: Deps; calls: Array<{ url: string; init: any }> } {
  const calls: Array<{ url: string; init: any }> = [];
  let i = 0;
  const fetch: FetchLike = async (url, init) => {
    calls.push({ url, init });
    const body = bodies[Math.min(i++, bodies.length - 1)];
    return { ok: true, status: 200, text: async () => body };
  };
  return { deps: { fetch }, calls };
}

const settings: LlmSettings = { provider: "openai", model: "gpt-4o-mini", apiKey: "sk-test" };

const TOOLS: ToolSchema[] = [
  {
    name: "write_range",
    description: "Write values to a range",
    parameters: { type: "object", properties: { address: { type: "string" }, values: { type: "array" } } },
  },
];

function toolCallResponse(name: string, args: object) {
  return JSON.stringify({
    choices: [
      {
        message: {
          role: "assistant",
          content: null,
          tool_calls: [{ id: "call_1", type: "function", function: { name, arguments: JSON.stringify(args) } }],
        },
      },
    ],
  });
}
function finalResponse(content: string) {
  return JSON.stringify({ choices: [{ message: { role: "assistant", content } }] });
}

describe("runAgent", () => {
  test("executes a tool call, feeds the result back, then returns final text", async () => {
    const { deps, calls } = queueMock([
      toolCallResponse("write_range", { address: "A1", values: [[42]] }),
      finalResponse("Done — wrote 42 to A1."),
    ]);
    const executed: Array<{ name: string; args: any }> = [];
    const execute = async (name: string, args: any) => {
      executed.push({ name, args });
      return "wrote 1 cell";
    };

    const res = await runAgent("put 42 in A1", TOOLS, settings, deps, execute);

    expect(executed).toEqual([{ name: "write_range", args: { address: "A1", values: [[42]] } }]);
    expect(res.finalText).toContain("Done");
    expect(res.steps).toHaveLength(1);
    expect(res.steps[0]).toMatchObject({ tool: "write_range", result: "wrote 1 cell" });
    // 2 LLM turns: the tool call, then the final message.
    expect(calls).toHaveLength(2);
    // The second request must include the tool result message.
    const secondBody = JSON.parse(calls[1].init.body);
    expect(secondBody.messages.some((m: any) => m.role === "tool" && m.content === "wrote 1 cell")).toBe(true);
  });

  test("handles tool-call arguments as an object (Ollama), not just a string", async () => {
    const ollamaStyle = JSON.stringify({
      choices: [
        {
          message: {
            role: "assistant",
            content: null,
            tool_calls: [
              { id: "c1", type: "function", function: { name: "write_range", arguments: { address: "A1", values: [[42]] } } },
            ],
          },
        },
      ],
    });
    const { deps } = queueMock([ollamaStyle, finalResponse("done")]);
    const executed: Array<{ name: string; args: any }> = [];
    await runAgent("x", TOOLS, settings, deps, async (name, args) => {
      executed.push({ name, args });
      return "ok";
    });
    expect(executed).toEqual([{ name: "write_range", args: { address: "A1", values: [[42]] } }]);
  });

  test("returns immediately when the model calls no tools", async () => {
    const { deps } = queueMock([finalResponse("Nothing to do.")]);
    const res = await runAgent("hi", TOOLS, settings, deps, async () => "unused");
    expect(res.finalText).toBe("Nothing to do.");
    expect(res.steps).toHaveLength(0);
  });

  test("captures tool errors as observations without aborting", async () => {
    const { deps } = queueMock([
      toolCallResponse("write_range", { address: "A1" }),
      finalResponse("Reported the error."),
    ]);
    const res = await runAgent("do it", TOOLS, settings, deps, async () => {
      throw new Error("range locked");
    });
    expect(res.steps[0].result).toMatch(/Error: range locked/);
    expect(res.finalText).toContain("Reported");
  });

  test("stops at the step limit", async () => {
    // Always returns a tool call -> never finishes.
    const { deps } = queueMock([toolCallResponse("write_range", { address: "A1", values: [[1]] })]);
    const res = await runAgent("loop", TOOLS, settings, deps, async () => "ok", { maxSteps: 3 });
    expect(res.steps).toHaveLength(3);
    expect(res.finalText).toMatch(/step limit/i);
  });
});

describe("createApprovalExecutor", () => {
  test("queues writes and passes reads through", async () => {
    const real = jest.fn(async (name: string) => `real:${name}`);
    const { executor, pending } = createApprovalExecutor(real, (n) => n.startsWith("write"));
    expect(await executor("read_range", { a: 1 })).toBe("real:read_range");
    expect(await executor("write_range", { address: "A1" })).toMatch(/Queued/);
    expect(pending).toEqual([{ name: "write_range", args: { address: "A1" } }]);
    expect(real).toHaveBeenCalledTimes(1);
  });

  test("runAgent in approval mode defers the write", async () => {
    const { deps } = queueMock([
      toolCallResponse("write_range", { address: "A1", values: [[1]] }),
      finalResponse("Planned the write."),
    ]);
    const real = jest.fn(async () => "applied");
    const { executor, pending } = createApprovalExecutor(real, (n) => n === "write_range");
    const res = await runAgent("write 1", TOOLS, settings, deps, executor);
    expect(pending).toHaveLength(1);
    expect(real).not.toHaveBeenCalled();
    expect(res.finalText).toContain("Planned");
  });
});

describe("chatWithTools", () => {
  test("sends tools in OpenAI format and parses tool_calls", async () => {
    const { deps, calls } = queueMock([toolCallResponse("write_range", { address: "B2", values: [[7]] })]);
    const out = await chatWithTools([{ role: "user", content: "x" }], TOOLS, settings, deps);
    expect(out.toolCalls).toHaveLength(1);
    expect(out.toolCalls[0]).toMatchObject({ name: "write_range" });
    const body = JSON.parse(calls[0].init.body);
    expect(body.tools[0]).toMatchObject({ type: "function", function: { name: "write_range" } });
    expect(body.tool_choice).toBe("auto");
  });

  test("reports token usage via onUsage", async () => {
    const body = JSON.stringify({
      choices: [{ message: { role: "assistant", content: "done" } }],
      usage: { prompt_tokens: 20, completion_tokens: 8, total_tokens: 28 },
    });
    const { deps } = queueMock([body]);
    const seen: any[] = [];
    await chatWithTools([{ role: "user", content: "x" }], TOOLS, settings, { ...deps, onUsage: (u) => seen.push(u) });
    expect(seen).toEqual([{ promptTokens: 20, completionTokens: 8, totalTokens: 28 }]);
  });

  test("throws before fetching when a required key is missing", async () => {
    const { deps, calls } = queueMock([finalResponse("unused")]);
    await expect(
      chatWithTools([{ role: "user", content: "x" }], TOOLS, { provider: "openai", model: "m" }, deps)
    ).rejects.toThrow(/No API key/);
    expect(calls).toHaveLength(0);
  });

  test("throws on an unknown provider", async () => {
    const { deps } = queueMock([finalResponse("unused")]);
    await expect(
      chatWithTools([{ role: "user", content: "x" }], TOOLS, { provider: "nope", model: "m" }, deps)
    ).rejects.toThrow(/Unknown provider/);
  });
});

/** Fetch double with control over ok/status, for error-path coverage. */
function errMock(body: string, ok: boolean, status = 200): Deps {
  const fetch: FetchLike = async () => ({ ok, status, text: async () => body });
  return { fetch };
}

describe("chatWithTools error handling", () => {
  test("surfaces a provider error message on an HTTP failure", async () => {
    const deps = errMock('{"error":{"message":"server boom"}}', false, 500);
    await expect(chatWithTools([], TOOLS, settings, deps)).rejects.toThrow("server boom");
  });

  test("falls back to the status code when the error body is unparseable", async () => {
    const deps = errMock("<html>502 Bad Gateway</html>", false, 502);
    await expect(chatWithTools([], TOOLS, settings, deps)).rejects.toThrow(/HTTP 502/);
  });

  test("surfaces an error object in an otherwise-ok body", async () => {
    const deps = errMock('{"error":"model missing"}', true, 200);
    await expect(chatWithTools([], TOOLS, settings, deps)).rejects.toThrow("model missing");
  });
});

describe("runAgent argument parsing", () => {
  test("malformed tool-call arguments degrade to an empty object, not a crash", async () => {
    const badArgs = JSON.stringify({
      choices: [
        {
          message: {
            role: "assistant",
            content: null,
            tool_calls: [{ id: "c1", type: "function", function: { name: "write_range", arguments: "not json{" } }],
          },
        },
      ],
    });
    const { deps } = queueMock([badArgs, finalResponse("recovered")]);
    const executed: any[] = [];
    const res = await runAgent("x", TOOLS, settings, deps, async (name, args) => {
      executed.push({ name, args });
      return "ok";
    });
    expect(executed).toEqual([{ name: "write_range", args: {} }]);
    expect(res.finalText).toContain("recovered");
  });
});
