# Excel LLM Add-in (Office.js)

The modern, cross-platform version of the add-in: an Office Web Add-in that runs
**identically on Excel for Mac, Windows, Web, and iPad**. It adds worksheet
functions like `=LLM.PROMPT("…")` plus a settings task pane — no VBA, no curl, no
temp files, and none of the encoding pain of the legacy `.xlam`.

> The legacy VBA add-in still lives at the repo root and is kept as a stopgap.
> This `officejs/` project is the going-forward product.

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
- A task pane (Home ▸ **LLM Settings**) to pick provider, model, key, or proxy
- Providers: OpenAI, Mistral, Nebius, Scaleway, OpenRouter, Ollama (local)

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

## Test it (on any platform, no Excel)

```bash
npm test         # Jest unit + functional tests over the core logic
npm run typecheck
```

The core (`src/core/*`) is Office-free and tested with a mocked `fetch`, so the
request-build → parse → error pipeline is verified deterministically — this is
the cross-platform CI gate.

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
  manifest.xml            # the add-in manifest (this is what you sideload/publish)
  src/core/               # pure TS: providers, llm client, config  (unit-tested)
  src/core/__tests__/     # Jest tests (no Office/network)
  src/functions/          # custom functions (=LLM.PROMPT, ...)
  src/taskpane/           # settings UI
  proxy/worker.js         # optional serverless key-custody + CORS proxy
  webpack.config.js       # build (also generates functions.json from JSDoc)
```
