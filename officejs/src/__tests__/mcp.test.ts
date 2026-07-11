import { parseRpcResponse, connectMcp, browserMcpFetch, McpFetch } from "../mcp";

function mcpMock(responses: string[], sessionId = "sess-1"): { fetchImpl: McpFetch; reqs: any[] } {
  const reqs: any[] = [];
  let i = 0;
  const fetchImpl: McpFetch = async (url, init) => {
    reqs.push({ url, body: JSON.parse(init.body), headers: init.headers });
    const text = responses[Math.min(i++, responses.length - 1)];
    return {
      ok: true,
      status: 200,
      text: async () => text,
      header: (n: string) => (n === "Mcp-Session-Id" ? sessionId : null),
    };
  };
  return { fetchImpl, reqs };
}

describe("parseRpcResponse", () => {
  test("plain JSON body", () => {
    const r = parseRpcResponse('{"jsonrpc":"2.0","id":2,"result":{"ok":1}}', 2);
    expect(r.result).toEqual({ ok: 1 });
  });

  test("SSE body (picks the data line)", () => {
    const sse = 'event: message\ndata: {"jsonrpc":"2.0","id":3,"result":{"x":2}}\n\n';
    expect(parseRpcResponse(sse, 3).result).toEqual({ x: 2 });
  });

  test("returns null when no message matches", () => {
    expect(parseRpcResponse("not json", 1)).toBeNull();
  });

  test("returns null on a malformed JSON body", () => {
    expect(parseRpcResponse('{"id":1,', 1)).toBeNull();
  });

  test("falls back to any result/error message when the id does not match", () => {
    // Some servers reply with a different id; still take the result-bearing message.
    const r = parseRpcResponse('{"jsonrpc":"2.0","id":99,"result":{"x":1}}', 1);
    expect(r.result).toEqual({ x: 1 });
  });
});

describe("connectMcp", () => {
  test("initializes, lists tools, and routes tools/call", async () => {
    const init = '{"jsonrpc":"2.0","id":1,"result":{"protocolVersion":"2024-11-05"}}';
    const list =
      '{"jsonrpc":"2.0","id":2,"result":{"tools":[{"name":"search","description":"web search","inputSchema":{"type":"object","properties":{"q":{"type":"string"}}}}]}}';
    const callResp = '{"jsonrpc":"2.0","id":3,"result":{"content":[{"type":"text","text":"result text"}]}}';
    const { fetchImpl, reqs } = mcpMock([init, list, callResp]);

    const conn = await connectMcp({ url: "https://mcp.example/rpc" }, fetchImpl);

    expect(reqs[0].body.method).toBe("initialize");
    expect(reqs[1].body.method).toBe("tools/list");
    expect(conn.tools).toEqual([
      { name: "search", description: "web search", parameters: { type: "object", properties: { q: { type: "string" } } } },
    ]);
    // session id from initialize propagates to later requests
    expect(reqs[1].headers["Mcp-Session-Id"]).toBe("sess-1");

    const out = await conn.executor("search", { q: "cats" });
    expect(out).toBe("result text");
    expect(reqs[2].body.method).toBe("tools/call");
    expect(reqs[2].body.params).toEqual({ name: "search", arguments: { q: "cats" } });
  });

  test("throws on a JSON-RPC error", async () => {
    const init = '{"jsonrpc":"2.0","id":1,"result":{}}';
    const err = '{"jsonrpc":"2.0","id":2,"error":{"code":-32601,"message":"Method not found"}}';
    const { fetchImpl } = mcpMock([init, err]);
    await expect(connectMcp({ url: "https://mcp.example/rpc" }, fetchImpl)).rejects.toThrow(/Method not found/);
  });

  test("stringifies a non-array tool result", async () => {
    const init = '{"jsonrpc":"2.0","id":1,"result":{}}';
    const list = '{"jsonrpc":"2.0","id":2,"result":{"tools":[{"name":"ping","description":"","inputSchema":{}}]}}';
    // result has no content array -> executor returns the JSON-stringified result.
    const callResp = '{"jsonrpc":"2.0","id":3,"result":{"value":42}}';
    const { fetchImpl } = mcpMock([init, list, callResp]);
    const conn = await connectMcp({ url: "https://mcp.example/rpc" }, fetchImpl);
    expect(await conn.executor("ping", {})).toBe('{"value":42}');
  });
});

describe("browserMcpFetch", () => {
  test("adapts the global fetch, exposing response headers", async () => {
    let seen: any;
    (global as any).fetch = async (url: string, init: any) => {
      seen = { url, init };
      return {
        ok: true,
        status: 200,
        text: async () => "body",
        headers: { get: (n: string) => (n === "Mcp-Session-Id" ? "sess-9" : null) },
      };
    };
    const r = await browserMcpFetch("https://mcp.example/rpc", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: "{}",
    });
    expect(r.ok).toBe(true);
    expect(r.status).toBe(200);
    expect(await r.text()).toBe("body");
    expect(r.header("Mcp-Session-Id")).toBe("sess-9");
    expect(seen.url).toBe("https://mcp.example/rpc");
  });
});
