# Publishing the Excel LLM Add-in

How the Office.js add-in gets from source to an installable product, and which
steps a machine/agent can do versus which need a human in a browser. Same spirit
as the repo's `appstore.md` (the iOS pipeline), but Office add-ins ship through
**GitHub Pages (hosting) + Microsoft Partner Center → AppSource** or **Microsoft
365 admin center (org deployment)** — there is no Apple/Xcode/signing here.

Unlike an `.xlam`, the whole build is plain text + `npm`, so everything up to the
Partner Center submission form is automatable and reproducible from any OS.

## The three distribution channels (pick by audience)

| Channel | End-user install effort | Who can set it up | Review? |
|---|---|---|---|
| **Sideload the hosted manifest** | ~4 steps (drop a file) | any user | none |
| **Org centralized deployment** | 0 clicks (auto-appears) | M365 admin | none |
| **AppSource store** | <5 clicks (Get Add-ins → Add) | Partner Center owner | Microsoft cert |

All three consume the **same hosted `dist/`** produced below.

## Step 1 — host `dist/` on HTTPS (automatable)

Production files must be served from a public HTTPS origin. This repo ships a
GitHub Pages deploy at `.github/workflows/deploy-pages.yml`, publishing to
`https://crispstrobe.github.io/ExcelLLMAddin/`.

- **One-time human step:** repo **Settings → Pages → Source: "GitHub Actions"**.
  (There is no API to flip this on for you; it's a single click.)
- After that, every push to `main` touching `officejs/**` rebuilds and redeploys.
- Hosting elsewhere? Change `prodOrigin`/`urlProd` in `webpack.config.js` and
  `npm run build`; `dist/manifest.xml` is rewritten to the new origin
  automatically.

**Gotcha — the dev certificate is dev-only.** The "not signed by a valid security
certificate" error only happens with the local `npm start` dev-server (fixed in
`webpack.config.js` by serving the `office-addin-dev-certs` cert). Pages serves a
real, browser-trusted certificate, so production never hits this.

## Step 2 — make the manifest production-ready (mostly automatable)

`npm run build` already rewrites localhost → the prod origin and validates clean
(`npm run validate`). Two things still need real values before any public release:

- **Regenerate the add-in GUID.** `manifest.xml` ships a fixed placeholder
  `<Id>`. Every published add-in needs its own stable, unique GUID — generate one
  with `uuidgen` (macOS/Linux) and paste it in. **Never** reuse the placeholder or
  change it after release (it's the add-in's identity).
- **Replace the placeholder icons.** `assets/icon-*.png` are 1×1 stubs. AppSource
  requires real icons at 16/32/64/80/128 px plus a 300×300 store logo. Drop real
  PNGs in `assets/` (same filenames) before submitting.

Validate any time:
```bash
npm run build && npx office-addin-manifest validate dist/manifest.xml
```

## Step 3a — sideload the hosted manifest (any single user, no review)

For an individual on **Excel for Mac** (no dev server, loads from Pages):
```bash
mkdir -p ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
curl -o ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/manifest.xml \
  https://crispstrobe.github.io/ExcelLLMAddin/manifest.xml
# fully quit + reopen Excel, then: Insert ▸ Add-ins ▸ My Add-ins ▸ (Developer/Shared) ▸ Excel LLM Add-in
```
On **Windows/Web**, host the manifest on a shared-folder catalog or use Step 3b.

## Step 3b — org centralized deployment (0-click for users, admin sets up once)

The fastest path to a real "it just appears" install for a team. A **Microsoft
365 admin** does this once:

1. **admin.microsoft.com → Settings → Integrated apps → Upload custom apps**
2. Provide the manifest (upload `dist/manifest.xml` or its Pages URL)
3. Assign to users/groups, accept permissions, Deploy

Users then see the add-in in Excel automatically (Mac/Windows/Web), no action
required. This step is **admin-web-UI only** — there is a Graph
`/deviceAppManagement` surface but Office add-in centralized deployment is not
reliably automatable, so treat it as human.

## Step 4 — AppSource (the true store, <5-click public install)

This is the equivalent of App Store submission — a human, browser-driven process
through **Partner Center**, with a Microsoft certification review. What's what:

**Automatable (done in this repo / by CI):**
- Build, host, and produce a validated production manifest.
- `npx office-addin-manifest validate` as a pre-flight.

**Human-only (no API for these):**
1. **Create a Partner Center account** and enrol in the *Microsoft 365 and
   Copilot* program (one-time, like the Apple Developer enrolment). partner.microsoft.com.
2. **Create the app submission**, upload `dist/manifest.xml`, and fill the store
   listing: description, screenshots, categories, **privacy policy URL**, **terms
   of use URL**, support URL.
3. **Submit for certification.** Microsoft validates against the commercial
   marketplace policies and *tests the add-in live*, then publishes to AppSource.

**Gotcha — reviewers must be able to actually run `=LLM.PROMPT`.** Like Apple's
demo-account requirement, Microsoft's testers need a working path end to end.
Provide, in the submission notes, either: (a) a reviewer OpenRouter key or a
deployed proxy URL to pre-fill, or (b) explicit steps to point at local Ollama.
A cloud provider with no key/proxy will fail review on CORS.

**Gotcha — privacy policy + terms are mandatory for AppSource** (not for
sideload/centralized). Have real URLs ready before starting the submission; the
form blocks without them.

**Gotcha — the manifest GUID and version are identity.** Submit with the real
regenerated GUID (Step 2). Each update bumps `<Version>` (must be ≥ the live one);
the GUID never changes.

## What a human must do, minimally, end to end

1. Flip on GitHub Pages once (Step 1). ← the only thing blocking a working
   sideload/centralized install
2. For AppSource only: Partner Center account, listing content (privacy/terms/
   screenshots), and hit Submit (Step 4).

Everything else — build, host, validate, redeploy on every push — is automated.
