// Minimal Cloudflare Worker proxy for the Excel LLM add-in.
//
// Why: most LLM provider APIs reject direct browser calls (no CORS), and you do
// not want API keys living in the workbook. This worker holds the keys as
// environment secrets, adds CORS, and speaks the add-in's normalized envelope:
//
//   POST { op: "chat"|"models", provider, model, prompt, system, baseUrl? }
//   ->  { content: string }              (op = chat)
//   ->  { models: string[] }             (op = models)
//   ->  { error: string }                (on failure, plus non-2xx status)
//
// Deploy: `npm i -g wrangler && wrangler deploy`, set secrets with
// `wrangler secret put OPENAI_API_KEY` (etc.). Point the add-in's Proxy URL at
// the deployed https URL. Lock ALLOWED_ORIGIN down to your add-in's origin.

const ALLOWED_ORIGIN = "*"; // tighten to your add-in origin in production

const PROVIDERS = {
  openai: { baseUrl: "https://api.openai.com/v1", style: "openai", keyEnv: "OPENAI_API_KEY" },
  mistral: { baseUrl: "https://api.mistral.ai/v1", style: "openai", keyEnv: "MISTRAL_API_KEY" },
  nebius: { baseUrl: "https://api.studio.nebius.com/v1", style: "openai", keyEnv: "NEBIUS_API_KEY" },
  scaleway: { baseUrl: "https://api.scaleway.ai/v1", style: "openai", keyEnv: "SCALEWAY_API_KEY" },
  openrouter: { baseUrl: "https://openrouter.ai/api/v1", style: "openai", keyEnv: "OPENROUTER_API_KEY" },
  ollama: { baseUrl: "http://localhost:11434", style: "ollama", keyEnv: null },
};

const SYSTEM_DEFAULT =
  "You are a helpful assistant embedded in a spreadsheet. " +
  "Answer concisely and return plain text suitable for a cell unless asked otherwise.";

function cors(headers = {}) {
  return {
    "Access-Control-Allow-Origin": ALLOWED_ORIGIN,
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
    "Content-Type": "application/json",
    ...headers,
  };
}

function json(body, status = 200) {
  return new Response(JSON.stringify(body), { status, headers: cors() });
}

export default {
  async fetch(request, env) {
    if (request.method === "OPTIONS") return new Response(null, { headers: cors() });
    if (request.method !== "POST") return json({ error: "POST only" }, 405);

    let req;
    try {
      req = await request.json();
    } catch {
      return json({ error: "Invalid JSON body" }, 400);
    }

    const spec = PROVIDERS[String(req.provider || "").toLowerCase()];
    if (!spec) return json({ error: `Unknown provider '${req.provider}'` }, 400);

    const baseUrl = (req.baseUrl || spec.baseUrl).replace(/\/+$/, "");
    const apiKey = spec.keyEnv ? env[spec.keyEnv] : undefined;
    if (spec.keyEnv && !apiKey) return json({ error: `Missing secret ${spec.keyEnv}` }, 500);

    const headers = { "Content-Type": "application/json" };
    if (apiKey) headers["Authorization"] = `Bearer ${apiKey}`;
    if (spec.style === "openai" && req.provider === "openrouter") {
      headers["HTTP-Referer"] = "https://excel-llm-addin";
      headers["X-Title"] = "Excel LLM Add-in";
    }

    try {
      if (req.op === "models") {
        const url = spec.style === "ollama" ? `${baseUrl}/api/tags` : `${baseUrl}/models`;
        const r = await fetch(url, { headers });
        const data = await r.json();
        if (data.error) return json({ error: errMsg(data.error) }, r.status || 502);
        const models =
          spec.style === "ollama"
            ? (data.models || []).map((m) => m.name)
            : (data.data || []).map((m) => m.id);
        return json({ models });
      }

      // op === "chat"
      const messages = [
        { role: "system", content: req.system || SYSTEM_DEFAULT },
        { role: "user", content: String(req.prompt ?? "") },
      ];
      const body = { model: req.model, messages };
      if (spec.style === "ollama") body.stream = false;

      const url = spec.style === "ollama" ? `${baseUrl}/api/chat` : `${baseUrl}/chat/completions`;
      const r = await fetch(url, { method: "POST", headers, body: JSON.stringify(body) });
      const data = await r.json();
      if (data.error) return json({ error: errMsg(data.error) }, r.status || 502);

      const content =
        data.choices?.[0]?.message?.content ??
        data.choices?.[0]?.text ??
        data.message?.content;
      if (content == null) return json({ error: "No content in provider response" }, 502);
      return json({ content: String(content) });
    } catch (e) {
      return json({ error: e && e.message ? e.message : String(e) }, 502);
    }
  },
};

function errMsg(err) {
  if (typeof err === "string") return err;
  if (err && typeof err === "object" && typeof err.message === "string") return err.message;
  return JSON.stringify(err);
}
