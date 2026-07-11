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

// Provider table generated from shared/providers.json (npm run gen:providers).
const PROVIDERS = {
  // <providers:generated>
  openai: { baseUrl: "https://api.openai.com/v1", style: "openai", keyEnv: "OPENAI_API_KEY" },
  mistral: { baseUrl: "https://api.mistral.ai/v1", style: "openai", keyEnv: "MISTRAL_API_KEY" },
  nebius: { baseUrl: "https://api.studio.nebius.com/v1", style: "openai", keyEnv: "NEBIUS_API_KEY" },
  scaleway: { baseUrl: "https://api.scaleway.ai/v1", style: "openai", keyEnv: "SCALEWAY_API_KEY" },
  openrouter: { baseUrl: "https://openrouter.ai/api/v1", style: "openai", keyEnv: "OPENROUTER_API_KEY" },
  groq: { baseUrl: "https://api.groq.com/openai/v1", style: "openai", keyEnv: "GROQ_API_KEY" },
  together: { baseUrl: "https://api.together.xyz/v1", style: "openai", keyEnv: "TOGETHER_API_KEY" },
  cerebras: { baseUrl: "https://api.cerebras.ai/v1", style: "openai", keyEnv: "CEREBRAS_API_KEY" },
  gemini: { baseUrl: "https://generativelanguage.googleapis.com/v1beta/openai", style: "openai", keyEnv: "GEMINI_API_KEY" },
  cohere: { baseUrl: "https://api.cohere.ai/compatibility/v1", style: "openai", keyEnv: "COHERE_API_KEY" },
  huggingface: { baseUrl: "https://router.huggingface.co/v1", style: "openai", keyEnv: "HF_TOKEN" },
  requesty: { baseUrl: "https://router.requesty.ai/v1", style: "openai", keyEnv: "REQUESTY_API_KEY" },
  ollama: { baseUrl: "http://localhost:11434", style: "ollama", keyEnv: null },
  // </providers:generated>
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

    // Image generation (Black Forest Labs / FLUX): submit then poll, server-side.
    // Not OpenAI-compatible, so it's handled before the chat-provider lookup.
    if (req.op === "image") {
      const bflKey = env.BFL_API_KEY;
      if (!bflKey) return json({ error: "Missing secret BFL_API_KEY" }, 500);
      const model = req.model || "flux-dev";
      try {
        const sub = await fetch(`https://api.bfl.ai/v1/${model}`, {
          method: "POST",
          headers: { "x-key": bflKey, "Content-Type": "application/json" },
          body: JSON.stringify({ prompt: String(req.prompt || ""), width: req.width || 1024, height: req.height || 768 }),
        });
        const subData = await sub.json();
        if (!sub.ok) return json({ error: errMsg(subData.detail || subData.error || `HTTP ${sub.status}`) }, sub.status || 502);
        const pollUrl = subData.polling_url;
        if (!pollUrl) return json({ error: "No polling_url from BFL" }, 502);
        for (let i = 0; i < 30; i++) {
          await new Promise((r) => setTimeout(r, 1500));
          const pr = await fetch(pollUrl, { headers: { "x-key": bflKey } });
          const pd = await pr.json();
          if (pd.status === "Ready") {
            const url = pd.result && pd.result.sample;
            return url ? json({ url }) : json({ error: "No image in BFL result" }, 502);
          }
          if (["Error", "Failed", "Content Moderated", "Request Moderated"].includes(pd.status)) {
            return json({ error: "Image generation " + pd.status }, 502);
          }
        }
        return json({ error: "Image generation timed out" }, 504);
      } catch (e) {
        return json({ error: e && e.message ? e.message : String(e) }, 502);
      }
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
      if (req.op === "embed") {
        const url = spec.style === "ollama" ? `${baseUrl}/api/embeddings` : `${baseUrl}/embeddings`;
        // Batch path: `inputs` is an array → return `embeddings: number[][]` in
        // order. Single path keeps `prompt` → `embedding`. Ollama has no array
        // shape here, so batch requests fan out to sequential single calls.
        if (Array.isArray(req.inputs)) {
          if (spec.style === "ollama") {
            const embeddings = [];
            for (const t of req.inputs) {
              const r = await fetch(url, {
                method: "POST",
                headers,
                body: JSON.stringify({ model: req.model, prompt: t }),
              });
              const data = await r.json();
              if (data.error) return json({ error: errMsg(data.error) }, r.status || 502);
              const emb = data.embedding ?? (data.embeddings && data.embeddings[0]);
              if (!Array.isArray(emb)) return json({ error: "No embedding in provider response" }, 502);
              embeddings.push(emb);
            }
            return json({ embeddings });
          }
          const r = await fetch(url, {
            method: "POST",
            headers,
            body: JSON.stringify({ model: req.model, input: req.inputs }),
          });
          const data = await r.json();
          if (data.error) return json({ error: errMsg(data.error) }, r.status || 502);
          const rows = data.data;
          if (!Array.isArray(rows)) return json({ error: "No embeddings in provider response" }, 502);
          const ordered = rows.every((x) => typeof x?.index === "number")
            ? [...rows].sort((a, b) => a.index - b.index)
            : rows;
          const embeddings = ordered.map((x) => x.embedding);
          if (embeddings.some((e) => !Array.isArray(e))) return json({ error: "Malformed embeddings row" }, 502);
          return json({ embeddings });
        }
        const body =
          spec.style === "ollama" ? { model: req.model, prompt: req.prompt } : { model: req.model, input: req.prompt };
        const r = await fetch(url, { method: "POST", headers, body: JSON.stringify(body) });
        const data = await r.json();
        if (data.error) return json({ error: errMsg(data.error) }, r.status || 502);
        const emb = data.data?.[0]?.embedding ?? data.embedding ?? (data.embeddings && data.embeddings[0]);
        if (!Array.isArray(emb)) return json({ error: "No embedding in provider response" }, 502);
        return json({ embedding: emb });
      }

      if (req.op === "models") {
        const url = spec.style === "ollama" ? `${baseUrl}/api/tags` : `${baseUrl}/models`;
        const r = await fetch(url, { headers });
        const data = await r.json();
        if (data.error) return json({ error: errMsg(data.error) }, r.status || 502);
        // OpenAI shape {data:[{id}]}, or a bare array (Together AI). Accept both.
        const rows = spec.style === "ollama" ? data.models || [] : Array.isArray(data) ? data : data.data || [];
        const models = rows.map((m) => (spec.style === "ollama" ? m.name : m.id || m.name)).filter(Boolean);
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
      // Pass token usage back so the task-pane meter works on the proxy path too.
      let usage;
      if (data.usage) {
        usage = {
          prompt_tokens: data.usage.prompt_tokens || 0,
          completion_tokens: data.usage.completion_tokens || 0,
          total_tokens: data.usage.total_tokens || 0,
        };
      } else if (spec.style === "ollama" && (data.prompt_eval_count || data.eval_count)) {
        usage = {
          prompt_tokens: data.prompt_eval_count || 0,
          completion_tokens: data.eval_count || 0,
          total_tokens: (data.prompt_eval_count || 0) + (data.eval_count || 0),
        };
      }
      return json({ content: String(content), usage });
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
