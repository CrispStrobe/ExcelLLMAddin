import { parseRpcResponse, connectMcp, McpFetch } from "../mcp";

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
});
