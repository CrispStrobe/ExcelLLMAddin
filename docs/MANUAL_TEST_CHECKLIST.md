# Mac Manual Smoke-Test Checklist

Automated Excel-in-CI isn't practical on macOS (no headless Excel COM), so Mac
verification is a documented manual pass. The **logic** is already covered by the
automated Windows CI (`tools/Run-Tests.ps1`, which uses `MockHttpClient` and runs
identically on both platforms); this checklist covers the **platform boundary**
that CI can't reach on Mac: the real curl transport, temp-file handling, and
UTF-8 across the shell.

Run these in Excel for Mac after importing the modules (or opening the built
`.xlam`).

## 0. In-process unit + functional tests (no network)
- [ ] In the VBA editor Immediate window (⌘+G), run: `?RunAllTests(True)`
- [ ] A dialog reports `FAIL: 0`. If not, read the Immediate window log.

_These are the same tests CI runs — they prove parsing/encoding/DI logic works on
this machine's VBA build._

## 1. Curl transport (real network, Ollama)
Prereq: `ollama serve` running with at least one model pulled.
- [ ] `=LLM_CONFIG()` shows a provider/model.
- [ ] `=LIST_MODELS("ollama")` spills your installed models.
- [ ] `=PROMPT("Say hello", "ollama", "<your-model>")` returns text, not `Error:`.
- [ ] Menu → `TestCurlConnection` succeeds.

## 2. UTF-8 across the shell (the old umlaut bug)
- [ ] `=PROMPT("Antworte nur mit dem Wort: Grüße")` returns `Grüße` correctly
      (no `√º`/mojibake). This exercises UTF-8 **out** (request body) and **in**
      (response) through curl.
- [ ] `=PROMPT("Reply with exactly: café 😀")` preserves the accent and emoji.

## 3. Injection / quoting safety
- [ ] `=PROMPT("Echo this verbatim: '; ls / #  and \"quotes\" and $HOME")`
      returns the text as data — no shell interpretation, no error. (Prompt goes
      through a body file + curl `--config`, never the command line.)

## 4. Concurrency / temp-file collisions
- [ ] Fill `A1:A10` with `=PROMPT("Return the number "&ROW())` and recalc.
      Each cell returns its own number (unique temp tokens; no cross-talk).

## 5. Cloud provider (if you have a key)
- [ ] Configure a key via `ShowSettings`, then `QuickTest` succeeds.
- [ ] `=LIST_MODELS("openai")` (or your provider) returns models.

## Record the run
Note macOS version, Excel version, and any failures in the PR description so we
track which Mac/Excel builds are verified.
