# Excel LLM Add-in (Office.js)

The modern, cross-platform version of the add-in: an Office Web Add-in that runs
**identically on Excel for Mac, Windows, Web, and iPad**. It adds worksheet
functions like `=LLM.PROMPT("…")` plus a settings task pane — no VBA, no curl, no
temp files, and none of the encoding pain of the legacy `.xlam`.

> A second edition — the VBA `.xlam` at the repo root — is the **fully offline /
> air-gapped** build (no hosting, no web server; pair it with local Ollama) and has
> near-parity with this one. This `officejs/` project is the cross-platform edition
> (Mac · Windows · Web · iPad) and adds AppSource, streaming, and the browser
> harness. See the root `README.md` for the comparison.

## What you get

- `=LLM.PROMPT(text, [provider], [model])` — ask an LLM from a cell
- `=LLM.STREAM(text, [provider], [model])` — like PROMPT, but streams into the cell live
- `=LLM.CLASSIFY(text, categories)` — pick one label from a range/list
- `=LLM.EXTRACT(text, instruction)` — pull a value out of text
- `=LLM.TRANSLATE(text, language)` — translate a cell
- `=LLM.SUMMARIZE(text, [maxWords])` — summarize a cell
- `=LLM.SENTIMENT(text)` — Positive / Neutral / Negative
- `=LLM.LIST(prompt, [count])` — generate a list, spilled down a column
- `=LLM.FIELDS(text, fields)` — extract fields into a spilled row (text → table)
- `=LLM.ASK(question, context)` — answer a question using a range as context
- `=LLM.SIMILARITY(a, b, [model])` — semantic similarity (0..1) via embeddings
- `=LLM.MAP(range, instruction)` — apply an instruction to every cell, spilling results
- `=LLM.LIST_MODELS([provider])` — spill available models
- `=LLM.CONFIG()` — show the active provider/model
- **Agent** — describe an edit in plain English; the model reads/writes ranges, formulas, and formatting on your sheet via tool-calling (extensible with remote MCP servers)
- A task pane (Home ▸ **LLM Settings**) to pick provider, model, key, or proxy
- Providers: OpenAI, Mistral, Nebius, Scaleway, OpenRouter, Groq, Together AI, Cerebras, Google Gemini, Cohere, Hugging Face, Requesty, Ollama (local)

To ship it to users (hosting, org deployment, AppSource), see **[PUBLISHING.md](PUBLISHING.md)**.

## Prerequisites

- Node.js 18+ (that's all you need to **build and test** — no Excel required)
- Excel (Mac/Windows/Web) only when you want to **run** the add-in

## Run it (dev / sideload)

```bash
cd officejs
npm install
npm start        # builds, starts the https dev-server, and sideloads into Excel
```

`npm start` opens Excel with the add-in loaded. Then:

1. On the **Home** tab, click **LLM Settings**.
2. Pick a provider, enter a model (and API key or proxy URL), click **Save**.
3. In any cell: `=LLM.PROMPT("Say hello")`.

That's the whole loop. To stop: `npm stop`.

### CORS reality (important)

Browsers block direct calls to most cloud LLM APIs, and you shouldn't put keys in
a workbook anyway. Two supported setups:

- **No proxy** — works out of the box for **OpenRouter** and **local Ollama**
  (both browser-friendly). Enter the key (OpenRouter) or just the model (Ollama).
- **Proxy (recommended for OpenAI/Mistral/…)** — deploy `proxy/worker.js` (a
  Cloudflare Worker), keep keys as server secrets, and set the task pane's
  **Proxy URL**. Keys never touch the workbook. See `proxy/worker.js` header for
  deploy steps.

## Agent (edit the sheet in plain English)

The task pane has an **Agent** box. Describe a change and the model operates your
workbook via tool-calling — `read_range`, `write_range`, `write_formula`,
`set_format`, `add_worksheet`, `create_chart`, `get_selection`, `list_sheets` — looping until done.

- **Approve-before-apply (default):** reads run live so the model sees your data,
  but writes are queued and shown as **Apply N changes** — you click to apply. A
  checkbox opts into auto-apply.
- **MCP (optional):** set an **MCP server URL** in Advanced. The add-in connects
  over HTTP (JSON-RPC: `initialize` → `tools/list`) and merges that server's tools
  with the Excel tools. (A sandboxed add-in can't do stdio MCP or host a server,
  but it can be an HTTP MCP client.)

Examples: *"In D1 put the sum of B2:B10, then bold anything over 100"*,
*"add a column classifying each row of my selection as high/low"*. Needs a
tool-calling-capable model (gpt-4o-mini, Claude, Llama-3.3, …).

## Test it (on any platform, no Excel)

```bash
npm test         # 160+ Jest unit + functional tests (~99% line coverage)
npm run typecheck
```

The core (`src/core/*`) is Office-free and tested with a mocked `fetch`, so the
request-build → parse → error pipeline is verified deterministically — this is
the cross-platform CI gate. What's covered, all without Excel or a network:

- **Transport** (`llm.ts`, `stream.ts`, `retry.ts`): direct + proxy
  chat/models/embeddings, SSE + NDJSON streaming, provider selection + routing
  (incl. every OpenAI-compat provider), headers, retry/backoff, and every error path.
- **Tasks** (`tasks.ts`): all worksheet functions incl. `MAP` batching/fallback
  and `SIMILARITY`/cosine.
- **Agent** (`agent.ts`, `excelTools.ts`): the tool-calling loop, approve-before-
  apply, arguments as string (OpenAI) *or* object (Ollama), and the Excel tool
  handlers driven against a fake `Excel` global (address parsing, resize, formula
  matrices, formatting).
- **MCP** (`mcp.ts`): JSON-RPC build + plain-JSON/SSE response parsing.
- **Config** (`config.ts`): settings persistence over a faked `OfficeRuntime.storage`.

### Live provider tests (real network, opt-in)

`src/core/__tests__/live.providers.test.ts` drives the **real** `runPrompt` /
`listModels` / `embed` against live endpoints — the same code the add-in runs. It
is **skipped by default** (no keys in CI) and each provider self-skips when its key
is absent. Enable it with your keys:

```bash
LIVE_PROVIDERS=1 GROQ_API_KEY=… OPENROUTER_API_KEY=… npx jest live.providers
```

Verified live via this suite: Groq, OpenRouter, Nebius, Mistral, Cohere, and
Hugging Face (chat + model listing), plus Nebius embeddings.

Excel-only behaviour (custom-function registration, `=LLM.PROMPT` in a live cell)
is checked separately — see `docs/MANUAL_TEST_CHECKLIST.md`.

## Dev harness (iterate on the task pane without Excel)

The task pane also runs in a normal browser with Office mocked — far faster than
the Excel reload loop. With the dev server running (`npm start` or `npm run
dev-server`), open:

```
https://localhost:3000/harness.html
```

Provider calls are **real** fetches, so CORS-friendly providers (OpenRouter,
Nebius, local Ollama) work end to end right in the browser — you can Save, Test,
and load models without Excel. (The harness is dev-only; it is never built into
`dist/` or shipped.)

Headless smoke test (drives the harness in Chrome, runs a real OpenRouter call):

```bash
OPENROUTER_API_KEY=sk-or-... npm run harness:smoke
```

Excel-only behaviour (custom-function registration, `=LLM.PROMPT` in a cell) still
has to be checked in Excel — but everything else iterates here in seconds.

## Build for production

```bash
npm run build    # emits dist/
```

Host `dist/` on any static HTTPS origin (GitHub Pages, Azure Static Web Apps,
Cloudflare Pages), set `urlProd` in `webpack.config.js` to that origin, rebuild,
and distribute `dist/manifest.xml`. For a true <5-click, auto-updating install,
publish to **AppSource** or your org's **Integrated Apps / add-in catalog**;
users then get it via **Insert ▸ Get Add-ins ▸ Add**.

## Layout

```
officejs/
  manifest.xml            # the add-in manifest (shared runtime; sideload/publish this)
  src/core/               # pure, unit-tested TS: providers, llm, tasks, agent,
                          #   cache, streamParser, config
  src/core/__tests__/     # Jest tests for the pure core (no Office/network)
  src/__tests__/          # Jest tests for the Office-facing edge (excelTools,
                          #   stream, mcp, browserFetch) via fake Excel/OfficeRuntime
  src/functions/          # custom functions (=LLM.PROMPT, ... via CustomFunctions)
  src/taskpane/           # settings UI + Agent panel
  src/excelTools.ts       # Excel.run tool handlers the agent calls
  src/mcp.ts              # optional MCP-over-HTTP client
  src/stream.ts           # streaming driver (=LLM.STREAM)
  src/harness/            # dev-only: run the task pane in a plain browser
  src/site/               # landing page + privacy/terms (deployed to Pages)
  proxy/worker.js         # optional serverless key-custody + CORS proxy
  tools/                  # icon generator, harness smoke test
  webpack.config.js       # build (generates functions.json from JSDoc)
```

Runs on a **shared runtime** — the task pane and custom functions share one
long-lived runtime, so opening the pane warms the functions.
