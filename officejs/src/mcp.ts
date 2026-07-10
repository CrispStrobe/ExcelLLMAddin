// Minimal MCP (Model Context Protocol) client over Streamable HTTP (JSON-RPC 2.0).
// Lets the agent use tools from a remote MCP server in addition to the Excel
// tools. A browser add-in can't do stdio MCP or host a server, but it can be an
// HTTP MCP *client*. The JSON-RPC construction + response parsing are pure and
// unit-tested; connectMcp wires them to fetch.

import { ToolSchema, ToolExecutor } from "./core/agent";

/* global fetch */

export interface McpConfig {
  url: string;
  headers?: Record<string, string>;
}

/** Fetch shape the client needs (adds response header access for the session id). */
export type McpFetch = (
  url: string,
  init: { method: string; headers: Record<string, string>; body: string }
) => Promise<{ ok: boolean; status: number; text: () => Promise<string>; header: (name: string) => string | null }>;

export interface McpConnection {
  tools: ToolSchema[];
  executor: ToolExecutor;
}

/** Extract the JSON-RPC message for `id` from a plain-JSON or SSE response body. */
export function parseRpcResponse(text: string, id: number): any {
  const t = text.trim();
  let candidates: any[] = [];
  if (t.startsWith("{") || t.startsWith("[")) {
    try {
      const j = JSON.parse(t);
      candidates = Array.isArray(j) ? j : [j];
    } catch {
      return null;
    }
  } else {
    for (const line of t.split("\n")) {
      const s = line.trim();
      if (!s.startsWith("data:")) continue;
      const p = s.slice(5).trim();
      if (p === "[DONE]") continue;
      try {
        candidates.push(JSON.parse(p));
      } catch {
        /* skip */
      }
    }
  }
  return (
    candidates.find((m) => m && m.id === id) ??
    candidates.find((m) => m && (m.result !== undefined || m.error !== undefined)) ??
    null
  );
}

/** Connect to an MCP server: initialize, list tools, and return an executor. */
export async function connectMcp(cfg: McpConfig, fetchImpl: McpFetch): Promise<McpConnection> {
  let sessionId: string | null = null;
  let id = 0;

  async function call(method: string, params: unknown): Promise<any> {
    id += 1;
    const myId = id;
    const headers: Record<string, string> = {
      "Content-Type": "application/json",
      Accept: "application/json, text/event-stream",
      ...(cfg.headers || {}),
    };
    if (sessionId) headers["Mcp-Session-Id"] = sessionId;

    const resp = await fetchImpl(cfg.url, {
      method: "POST",
      headers,
      body: JSON.stringify({ jsonrpc: "2.0", id: myId, method, params }),
    });
    const sid = resp.header("Mcp-Session-Id");
    if (sid) sessionId = sid;

    const msg = parseRpcResponse(await resp.text(), myId);
    if (!msg) throw new Error(`MCP: no JSON-RPC response for ${method} (HTTP ${resp.status}).`);
    if (msg.error) throw new Error(`MCP ${method}: ${msg.error.message ?? JSON.stringify(msg.error)}`);
    return msg.result;
  }

  await call("initialize", {
    protocolVersion: "2024-11-05",
    capabilities: {},
    clientInfo: { name: "excel-llm-addin", version: "1.0" },
  });

  const listed = await call("tools/list", {});
  const tools: ToolSchema[] = (listed?.tools || []).map((t: any) => ({
    name: t.name,
    description: t.description || "",
    parameters: t.inputSchema || { type: "object", properties: {} },
  }));

  const executor: ToolExecutor = async (name, args) => {
    const result = await call("tools/call", { name, arguments: args });
    const content = result?.content;
    if (Array.isArray(content)) return content.map((c: any) => c.text ?? JSON.stringify(c)).join("\n");
    return typeof result === "string" ? result : JSON.stringify(result);
  };

  return { tools, executor };
}

/** Browser McpFetch backed by the global fetch. */
export const browserMcpFetch: McpFetch = async (url, init) => {
  const r = await fetch(url, init as RequestInit);
  return {
    ok: r.ok,
    status: r.status,
    text: () => r.text(),
    header: (name: string) => r.headers.get(name),
  };
};
