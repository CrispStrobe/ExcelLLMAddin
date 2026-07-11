# Build & validate the VBA `.xlam` (Mac)

Excel-for-Mac has no VBA-project automation, so the `.xlam` is built and validated
by hand once (or on a Windows runner via `tools/Build-Addin.ps1`). This is the
Mac path — it takes ~10 minutes and gives you the fully offline, feature-parity
add-in.

## 0. Import the modules (VBA editor)

Open Excel, then **Tools ▸ Macro ▸ Visual Basic Editor** (or ⌥F11). Select your
workbook's project, **File ▸ Import File…**, and import all of these (order doesn't
matter):

- `vendor/Dictionary.cls`, `vendor/JsonConverter.bas`
- `modText.bas`, `IHttpClient.cls`, `MockHttpClient.cls`, `WinHttpClient.cls`, `CurlClient.cls`, `modHttp.bas`
- `modConfig.bas`, `modLLMFunctions.bas`
- `modTasks.bas`, `modAgent.bas`, `modMcp.bas`
- `modMenu.bas`, `modTests.bas`

## 1. Run the self-tests (no network)

- [ ] In the Immediate window (⌘G), run: `?RunAllTests(True)`
- [ ] Dialog reports **`FAIL: 0`** (~43 tests: UTF-8, JSON, request/parse, cache,
      task helpers, embeddings, agent, MCP). If anything fails, read the Immediate
      window log and send it over.

Then **File ▸ Save As ▸ Excel Add-in (.xlam)** as `ExcelLLMAddin.xlam`. Enable it
via **Tools ▸ Excel Add-ins**. (Or `installer/mac-xlam/` builds a `.pkg`.)

## 2. Core, in cells (needs a provider — Ollama offline, or a cloud key)

Configure first: **Tools ▸ Macro ▸ ShowSettings** (pick provider/model/key).

- [ ] `=LLM_CONFIG()` shows the provider/model
- [ ] `=LIST_MODELS("ollama")` spills your models
- [ ] `=PROMPT("Say hello")` returns text (not `Error:`)

## 3. Task functions

- [ ] `=CLASSIFY("I love this", "Positive,Negative")` → `Positive`
- [ ] `=TRANSLATE("good morning", "German")` → `Guten Morgen`
- [ ] `=SENTIMENT("this is terrible")` → `Negative`
- [ ] `=SUMMARIZE(A1, 10)` (A1 = a paragraph) → short summary
- [ ] `=EXTRACT("call bob@x.com", "the email")` → `bob@x.com`
- [ ] `=LIST("3 primary colors", 3)` → spills 3 rows
- [ ] `=FIELDS("Bob, 30, NYC", "name,age,city")` → spills a row of 3
- [ ] `=MAP(A1:A5, "translate to French")` → spills results per cell
- [ ] `=ASK("how many?", A1:B5)` → answer using the range as context
- [ ] `=SIMILARITY("cat","kitten","<embedding-model>")` → a number near 1
      (needs an embeddings-capable provider, e.g. Nebius `Qwen/Qwen3-Embedding-8B`)

## 4. Agent (the offline AI editor)

- [ ] **Tools ▸ Macro ▸ RunAgent**, enter: *"In D1 put the sum of B2:B10, then bold
      anything over 100"* (with numbers in B2:B10)
- [ ] It shows an action log, then a Yes/No **"Apply N changes?"** prompt
- [ ] Click **Yes** → D1 gets `=SUM(B2:B10)` and the formatting is applied
- [ ] (Optional) `SetMcpServer` with an MCP URL, then RunAgent again — the server's
      tools appear alongside the Excel tools

## 5. Transport correctness (the historic bug classes)

- [ ] Umlauts: `=PROMPT("Antworte nur mit: Grüße")` → `Grüße` (no `√º` mojibake)
- [ ] Injection-safe: `=PROMPT("Echo verbatim: '; ls / #  \"quotes\" $HOME")` →
      returned as data, no shell interpretation
- [ ] Concurrency: fill `A1:A10` with `=PROMPT("Return "&ROW())` → each cell its own
      number (no cross-talk)

## 6. Fully offline (Ollama)

- [ ] Disconnect from the network, `ollama serve` running with a pulled model
- [ ] Configure provider **ollama**, then `=PROMPT(...)`, a task function, and
      **RunAgent** all work with no internet.

## Record the run

Note macOS + Excel version and any failures in the PR/commit so we track which
builds are verified.
