# Excel LLM Add-in

Call LLMs straight from Excel cells — and let an **agent edit your sheet** — across
Excel for **Mac, Windows, Web, and iPad**. Bring your own provider (OpenAI,
Mistral, Nebius, Scaleway, OpenRouter, or local Ollama).

**▶ Live: https://crispstrobe.github.io/ExcelLLMAddin/** — install by sideloading
[`manifest.xml`](https://crispstrobe.github.io/ExcelLLMAddin/manifest.xml) (see
[install steps](officejs/PUBLISHING.md)).

There are two editions in this repo:

| Edition | Where | Best for |
|---|---|---|
| **Office.js add-in** (`officejs/`) | Mac · Windows · Web · iPad | store/AppSource, the widest reach, streaming |
| **VBA `.xlam`** (repo root `.bas`) | Mac · Windows | **fully offline / air-gapped**, single-file install |

Both have the function set *and* the sheet-editing agent. The `.xlam` runs with no
hosting and no internet (pair it with local Ollama); the Office.js edition adds
Web/iPad, AppSource, and `STREAM`.

## What it does

Worksheet functions (namespace `LLM`):

| Function | Purpose |
|---|---|
| `=LLM.PROMPT(text, [provider], [model])` | Ask an LLM from a cell |
| `=LLM.STREAM(text, …)` | Like PROMPT, streamed into the cell live |
| `=LLM.CLASSIFY(text, categories)` | Pick one label from a range/list |
| `=LLM.EXTRACT(text, instruction)` | Pull a single value out of text |
| `=LLM.FIELDS(text, fields)` | Extract many fields into a row (text → table) |
| `=LLM.TRANSLATE(text, language)` | Translate a cell |
| `=LLM.SUMMARIZE(text, [maxWords])` | Summarize a cell |
| `=LLM.SENTIMENT(text)` | Positive / Neutral / Negative |
| `=LLM.LIST(prompt, [count])` | Generate a list, spilled down a column |
| `=LLM.ASK(question, context)` | Answer a question using a range as context |
| `=LLM.SIMILARITY(a, b, [model])` | Semantic similarity (0–1) via embeddings |
| `=LLM.MAP(range, instruction)` | Apply an instruction to every cell (batched) |
| `=LLM.LIST_MODELS([provider])` · `=LLM.CONFIG()` | List models · show config |

Plus an **Agent**: describe a change in plain English in the task pane and the model
reads/edits your workbook (ranges, formulas, formatting, sheets) via tool-calling —
with **approve-before-apply** by default. Its toolset can be extended with remote
**MCP** servers.

## Quick start

**Use it (any user):** sideload the hosted manifest — see
[officejs/PUBLISHING.md](officejs/PUBLISHING.md) for Mac / org-deployment / AppSource.

**Develop it:**
```bash
cd officejs
npm install
npm start        # builds, starts the https dev-server, sideloads into Excel
npm test         # 70+ unit/functional tests, no Excel needed
```
Details, the browser dev-harness, and the CORS/proxy notes are in
[officejs/README.md](officejs/README.md).

## Providers

OpenAI · Mistral · Nebius · Scaleway · OpenRouter · Ollama (local). OpenRouter,
Nebius, and local Ollama work directly from the browser; others use the optional
key-custody proxy (`officejs/proxy/worker.js`).

## Offline VBA edition (`.xlam`)

The VBA add-in (repo root `.bas`/`.cls`) is **fully self-contained and offline** —
no hosting, no web server. It has near-parity with the Office.js edition:

- Functions: `=PROMPT`, `=CLASSIFY`, `=EXTRACT`, `=TRANSLATE`, `=SUMMARIZE`,
  `=SENTIMENT`, `=ASK`, `=LIST`, `=FIELDS`, `=MAP`, `=SIMILARITY`, `=LIST_MODELS`,
  `=LLM_CONFIG` (`modLLMFunctions.bas`, `modTasks.bas`).
- **Agent** (`modAgent.bas`) — run the `RunAgent` macro; the model edits the sheet
  via native `Range` tools with approve-before-apply. With local Ollama, this is a
  fully air-gapped AI that edits your workbook.
- **MCP** (`modMcp.bas`) — run `SetMcpServer` to add a remote MCP server's tools
  to the agent (best-effort JSON-RPC over HTTP; targets stateless servers).
- Solid plumbing: injected `IHttpClient` (WinHTTP/curl), real JSON (vendored
  VBA-JSON), UTF-8, a response cache, and a `RunAllTests` self-test harness.
- Only unported feature: `STREAM` (VBA UDFs are synchronous — no live cell updates).

**Build:** `pwsh tools/Build-Addin.ps1` on Windows+Excel, or import the modules in
the VBA editor and Save As `.xlam` (Excel is required to compile VBA — it can't be
built on Mac or GitHub-hosted CI). **Install offline:** `installer/mac-xlam/`
builds a `.pkg` that drops the `.xlam` into Excel's add-ins folder. See
`installer/README.md`.

## License

AGPL-3.0.
