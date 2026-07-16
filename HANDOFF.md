# Handoff — optimization pass (embeddings · cache · provider de-dup)

Status for the next agent, especially one on a **Windows + Desktop Excel** machine
(needed to build/test the VBA `.xlam`, which can't be done on Mac/Linux).

All work below is **committed and pushed to `main`** (`origin/main`). Pull first.

## What changed (5 commits on `main`)

| Commit | Summary |
|---|---|
| `perf(embeddings)` | `RECALL`/`SIMILARITY` batch candidate embeddings into a few requests instead of one-per-row — TS, proxy, **and VBA**. |
| `perf(cache)` | Office.js prompt cache now persists across functions-runtime reloads (`persistentCache.ts`). |
| `refactor(providers)` | `shared/providers.json` is the single source of truth; TS + proxy tables are generated from it; a guard test also checks the VBA URLs. Fixed a live drift: VBA had the dead `api.studio.nebius.ai`, now `.com`. |
| `docs` | READMEs updated. |
| `perf(config)` | Settings-read cache so a bulk recalc of many `=LLM.PROMPT` cells does ~1 storage read, not one per cell. |
| `fix(usage)` | Embedding tokens (`=SIMILARITY`/`=RECALL`) now count in the task-pane meter; proxy passes embeddings `usage` back. **TS/proxy only — no VBA impact.** |

## Verified on this (macOS) machine — Office.js edition

All green; safe to trust:

```bash
cd officejs
npm ci                 # if node_modules is absent
npm run typecheck      # tsc --noEmit — clean
npm test               # 242 pass, 25 skipped (skipped = opt-in live-provider suite)
npm run check:providers # provider tables in sync with shared/providers.json
npm run build          # webpack production build compiles (functions.js, taskpane.js)
```

## NOT verified here — needs Windows + Desktop Excel

The VBA edition was edited but **not built or run** (no Excel on this host). Do this on Windows:

1. **Rebuild the add-in** from the `.bas`/`.cls` sources (the `.xlam` is a build artifact):
   ```powershell
   pwsh tools/Build-Addin.ps1      # imports modules in order → ExcelLLMAddin.xlam
   ```
   Requires Desktop Excel with **"Trust access to the VBA project object model"** on
   (the scripts enable it for the current user).

2. **Run the VBA test suite** (headless; injects `MockHttpClient`, no network needed):
   ```powershell
   pwsh tools/Run-Tests.ps1        # runs RunAllTests, writes test-results.xml (JUnit)
   ```
   Expect **`FAIL: 0`**. New test added this pass: `Test_Task_EmbedBatch`
   (`modTests.bas`) — covers the single-request batch path + index reordering +
   the Ollama fallback. Entry point is `RunAllTests` in `modTests.bas`.

3. **Live smoke in a real cell** (needs an embeddings-capable provider, e.g. Nebius
   `Qwen/Qwen3-Embedding-8B`) — see `docs/MANUAL_TEST_CHECKLIST.md`:
   - `=SIMILARITY("cat","kitten","<embed-model>")` → ~1
   - `=RECALL("query", A1:A50, 5, "<embed-model>")` over a range of ≥20 rows —
     confirm it returns and is fast; this now issues **one** `/embeddings` request
     for all candidates (watch the network tab / provider dashboard to confirm the
     batching if you can).

## VBA changes to review (the batched path)

- `modLLMFunctions.bas` — new `EmbedVectorsBatch(texts, embModel, provider)`: one
  `/embeddings` call with an `input` array (openai-style); returns `Nothing` for
  Ollama or on any failure so the caller falls back to per-item `EmbedVector`.
- `modTasks.bas` — `RECALL` now calls `EmbedVectorsBatch` first, per-item fallback
  when it returns `Nothing`. `SIMILARITY` unchanged (only 2 calls).
- `modConfig.bas`, `modMenu.bas` — Nebius base URL `…nebius.ai` → `…nebius.com`.
- `modTests.bas` — `Test_Task_EmbedBatch` added and registered in the run list.

## Parked follow-up (optional, deliberately not done)

VBA provider base URLs are currently **guarded** against drift (the jest test
`officejs/src/core/__tests__/providers.gen.test.ts` fails CI if `modConfig.bas` /
`modMenu.bas` disagree with `shared/providers.json`) but **not generated**. Fully
generating them (marker-delimited blocks in `modConfig.bas`, emitted by
`officejs/tools/gen-providers.cjs`) would finish the de-dup story. Left out because
it's an invasive edit to hand-tuned VBA that can't be validated without Excel — do
it on Windows if desired, then re-run `Build-Addin.ps1` + `Run-Tests.ps1`.

## Single source of truth for providers

Edit **`shared/providers.json`** only, then regenerate the Office.js tables:

```bash
cd officejs && npm run gen:providers   # rewrites src/core/providers.generated.ts + proxy/worker.js block
```

`src/core/providers.ts` keeps the types/endpoint helpers and re-exports the
generated `PROVIDERS`. The VBA edition reads the same list conceptually but its
strings are hand-maintained (see the parked follow-up).
