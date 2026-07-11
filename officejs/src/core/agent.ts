// Tool-calling agent loop. Pure and testable: the LLM transport is the injected
// Deps.fetch, and side effects (actually touching Excel) happen through an
// injected `execute` function. The Excel tool implementations live in
// ../excelTools (browser/Office.js); MCP or other tool sources can be added the
// same way — this loop only cares about tool schemas + an executor.

import {
  getProvider,
  chatEndpoint,
} from "./providers";
import { LlmSettings, LlmError, directHeaders, Deps } from "./llm";
import { parseUsage } from "./usage";

export interface ToolSchema {
  name: string;
  description: string;
  /** JSON Schema for the tool's arguments. */
  parameters: Record<string, unknown>;
}

export interface ToolCall {
  id: string;
  name: string;
  arguments: string; // raw JSON string from the model
}

export interface AgentStep {
  tool: string;
  args: unknown;
  result: string;
}

export interface AgentResult {
  finalText: string;
  steps: AgentStep[];
}

export type ToolExecutor = (name: string, args: any) => Promise<string>;

export interface AgentOptions {
  maxSteps?: number;
  system?: string;
}

export const DEFAULT_AGENT_SYSTEM =
  "You are an assistant operating on the user's Excel worksheet via tools. " +
  "Read what you need, then make the requested changes in small, verifiable steps. " +
  "Prefer writing formulas over hardcoded values when it fits. Address ranges use A1 " +
  "notation (e.g. Sheet1!B2:B10). When the task is done, reply with a short summary and " +
  "no further tool call.";

/** One turn: send messages + tools, get back the assistant's content and tool calls. */
export async function chatWithTools(
  messages: any[],
  tools: ToolSchema[],
  settings: LlmSettings,
  deps: Deps
): Promise<{ message: any; toolCalls: ToolCall[] }> {
  const spec = getProvider(settings.provider);
  if (!spec) throw new LlmError(`Unknown provider '${settings.provider}'.`);
  const baseUrl = settings.baseUrl || spec.defaultBaseUrl;
  if (spec.requiresKey && !settings.apiKey) {
    throw new LlmError(`No API key configured for ${spec.label}.`);
  }

  const url = chatEndpoint(spec, baseUrl);
  const body = {
    model: settings.model,
    messages,
    tools: tools.map((t) => ({ type: "function", function: t })),
    tool_choice: "auto",
  };
  const resp = await deps.fetch(url, {
    method: "POST",
    headers: directHeaders(spec, settings.apiKey),
    body: JSON.stringify(body),
  });
  const text = await resp.text();
  if (!resp.ok) throw new LlmError(errorMessage(text) ?? `HTTP ${resp.status}`);
  if (deps.onUsage) {
    const usage = parseUsage(text, spec.style);
    if (usage) deps.onUsage(usage);
  }

  const data = safeJson(text) as any;
  if (data?.error) throw new LlmError(typeof data.error === "string" ? data.error : data.error.message);

  const message = data?.choices?.[0]?.message ?? { role: "assistant", content: "" };
  const toolCalls: ToolCall[] = (message.tool_calls || []).map((tc: any) => {
    // OpenAI-style APIs return arguments as a JSON string; Ollama returns it as an
    // already-parsed object. Normalize to a string so runAgent can JSON.parse it.
    const a = tc.function?.arguments;
    return {
      id: tc.id,
      name: tc.function?.name,
      arguments: typeof a === "string" ? a : JSON.stringify(a ?? {}),
    };
  });
  return { message, toolCalls };
}

/** Run the agent loop until the model stops calling tools or maxSteps is hit. */
export async function runAgent(
  instruction: string,
  tools: ToolSchema[],
  settings: LlmSettings,
  deps: Deps,
  execute: ToolExecutor,
  options: AgentOptions = {}
): Promise<AgentResult> {
  const maxSteps = options.maxSteps ?? 8;
  const messages: any[] = [
    { role: "system", content: options.system ?? DEFAULT_AGENT_SYSTEM },
    { role: "user", content: instruction },
  ];
  const steps: AgentStep[] = [];

  for (let step = 0; step < maxSteps; step++) {
    const { message, toolCalls } = await chatWithTools(messages, tools, settings, deps);

    if (toolCalls.length === 0) {
      return { finalText: String(message.content ?? ""), steps };
    }

    messages.push(message);
    for (const call of toolCalls) {
      let args: any = {};
      try {
        args = JSON.parse(call.arguments || "{}");
      } catch {
        args = {};
      }
      let result: string;
      try {
        result = await execute(call.name, args);
      } catch (e) {
        result = "Error: " + (e instanceof Error ? e.message : String(e));
      }
      steps.push({ tool: call.name, args, result });
      messages.push({ role: "tool", tool_call_id: call.id, content: result });
    }
  }

  return { finalText: "Stopped: reached the step limit before finishing.", steps };
}

export interface PendingAction {
  name: string;
  args: any;
}

/**
 * Wrap a real executor so that mutating ("write") tools are queued for approval
 * instead of run. Read tools pass through (the model still sees real data). The
 * returned `pending` array collects the deferred actions to apply later.
 */
export function createApprovalExecutor(
  real: ToolExecutor,
  isWriteTool: (name: string) => boolean
): { executor: ToolExecutor; pending: PendingAction[] } {
  const pending: PendingAction[] = [];
  const executor: ToolExecutor = async (name, args) => {
    if (isWriteTool(name)) {
      pending.push({ name, args });
      return `Queued ${name} for approval (not applied yet).`;
    }
    return real(name, args);
  };
  return { executor, pending };
}

function safeJson(text: string): unknown {
  try {
    return JSON.parse(text);
  } catch {
    return null;
  }
}

function errorMessage(text: string): string | undefined {
  const d = safeJson(text) as any;
  if (d?.error) return typeof d.error === "string" ? d.error : d.error.message;
  return undefined;
}
