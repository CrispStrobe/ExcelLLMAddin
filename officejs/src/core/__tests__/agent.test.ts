import { runAgent, chatWithTools, ToolSchema } from "../agent";
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
});
