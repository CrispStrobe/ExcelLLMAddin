# Publishing to Microsoft AppSource

How to get the Excel LLM Add-in into the one marketplace that exists for Office
add-ins. Written in the spirit of the repo's `appstore.md` (the Apple pipeline) —
but this is a *different world*, so read the first section carefully.

## The platform reality (important)

An Office.js add-in is **not a native app**. It cannot go on the Apple App Store,
Google Play, or the Mac App Store. There are exactly two ways to distribute it:

| Channel | Who installs | Review | Notes |
|---|---|---|---|
| **Microsoft AppSource** | anyone (Insert ▸ Get Add-ins) | Microsoft certification (days) | the public "store" |
| **Org deployment** (M365 admin center → Integrated apps) | your tenant's users, pushed by an admin | none | fastest for a team; no store |
| Sideload the hosted manifest | one user | none | dev / personal |

Unlike Apple's App Store Connect API (which automates almost everything),
Microsoft's Office-add-in submission is **mostly a Partner Center web-UI process**.
There is no reliable public API to submit an Office-add-in marketplace offer. So
the split is inverted from `appstore.md`: the *technical prep* is automatable (and
done here); the *submission* is human.

## What's automatable vs human

| Step | Who |
|---|---|
| Production manifest (validated, real GUID, prod URLs) | **Done** (`npm run build` → `dist/manifest.xml`, `npm run validate`) |
| Hosting on HTTPS | **Done** (GitHub Pages, auto-deploy) |
| Privacy policy + Terms of Use URLs (mandatory) | **Done** (`/privacy.html`, `/terms.html`) |
| Icons + screenshots | **Done** (generated; screenshots in `store/screenshots/`) |
| Listing copy (name, descriptions, keywords, categories) | **Done** (below) |
| Enrol in Partner Center | **Human** (browser, one-time) |
| Create the offer + upload manifest + paste listing | **Human** (Partner Center UI) |
| Submit for certification | **Human** (a real decision) |
| Certification outcome | Microsoft (days) |

## Step 1 — enrol in Partner Center (human, one-time)

1. Go to **partner.microsoft.com** and sign in with a work/Microsoft account.
2. Enrol in the **Microsoft Cloud Partner Program** and open the **Marketplace**
   program (publishing to the commercial marketplace / AppSource is free).
3. Complete the account/publisher profile (publisher display name, etc.).

## Step 2 — produce the production artifacts (automatable — already done)

```bash
cd officejs
npm run build                 # dist/ with prod URLs (crispstrobe.github.io/ExcelLLMAddin)
npx office-addin-manifest validate dist/manifest.xml
```
The hosted, validated manifest to submit is
`https://crispstrobe.github.io/ExcelLLMAddin/manifest.xml`.

**Before a real submission, confirm two things in `manifest.xml`:**
- `<Id>` is a real unique GUID (it is — regenerated), and never changes across
  updates.
- The `<SupportUrl>` is a page you monitor (currently the GitHub repo).

## Step 3 — create the offer (human, Partner Center UI)

1. Partner Center → **Marketplace offers** → **New offer** → **Office add-in**
   (under the Microsoft 365 and Copilot program).
2. **Packages / technical config:** upload `dist/manifest.xml` (or point at the
   hosted URL).
3. **Listing:** paste the copy from Step 4; upload the logo (`assets/icon-*.png`)
   and screenshots (`store/screenshots/`).
4. **Properties:** set the categories (Step 4).
5. **Availability / pricing:** free; choose markets.
6. **Submit** for certification. Microsoft validates against the commercial
   marketplace policies *and tests the add-in live*, then publishes.

## Step 4 — listing content (ready to paste)

- **Name:** Excel LLM Add-in
- **Subtitle / summary (short):** Call OpenAI, Mistral, Nebius, OpenRouter, or
  local Ollama from a cell — with an agent that can edit your sheet.
- **Categories:** Productivity; Data & analytics
- **Search keywords:** LLM, AI, GPT, OpenAI, OpenRouter, Ollama, prompt,
  classify, translate, summarize, embeddings, agent
- **Support URL:** https://github.com/CrispStrobe/ExcelLLMAddin/issues
- **Privacy policy URL:** https://crispstrobe.github.io/ExcelLLMAddin/privacy.html
- **Terms of use URL:** https://crispstrobe.github.io/ExcelLLMAddin/terms.html
- **Description (long):**

  > Bring large language models into Excel. Use simple worksheet functions —
  > =LLM.PROMPT, CLASSIFY, EXTRACT, FIELDS, TRANSLATE, SUMMARIZE, SENTIMENT,
  > LIST, ASK, SIMILARITY, MAP, and a live-streaming STREAM — to run AI over your
  > data right in the grid. A built-in agent goes further: describe a change in
  > plain English and it reads and edits your workbook (ranges, formulas,
  > formatting, sheets) via tool-calling, with approve-before-apply safety.
  >
  > Bring your own provider: OpenAI, Mistral, Nebius, Scaleway, OpenRouter, or a
  > local Ollama server. Your API keys stay on your device (or on your own
  > serverless proxy) — nothing is sent to us. Runs identically on Excel for Mac,
  > Windows, the web, and iPad.
  >
  > Open source (AGPL-3.0).

## Gotchas (AppSource-specific)

- **Reviewers must be able to run it.** Like Apple's demo-account rule, Microsoft
  testers need a working path end to end. In the submission notes, give them a
  reviewer OpenRouter key (or a deployed proxy URL) to paste into the task pane,
  or step-by-step Ollama instructions. A cloud provider with no key/proxy fails
  review on CORS.
- **Privacy + Terms URLs are hard-required** for AppSource (they're optional for
  sideload/org deployment). Both are live at the URLs above.
- **The GUID is identity.** Submit with the real `<Id>`; every update bumps
  `<Version>` (≥ the live one); the GUID never changes.
- **Real logos, not the 1×1 stubs.** The icons are generated (fine to submit) but
  swap for branded art before a public listing if you want polish.
- **Contact info** for the publisher profile: confirm the support email/phone you
  want on file (Partner Center asks for them).

## Faster alternative: org deployment (no store, no review)

If you just want your team/tenant to have it now: **admin.microsoft.com →
Settings → Integrated apps → Upload custom apps** → provide the hosted manifest →
assign to users. It appears in their Excel automatically. See `../PUBLISHING.md`.
