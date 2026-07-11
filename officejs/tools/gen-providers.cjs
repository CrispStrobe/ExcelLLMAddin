#!/usr/bin/env node
// Generate the per-edition provider tables from the single source of truth
// (shared/providers.json), so the TS catalog, the proxy worker, and (via a guard
// test) the VBA add-in can't silently drift apart.
//
//   node tools/gen-providers.cjs         # rewrite generated files
//   node tools/gen-providers.cjs --check # exit 1 if they are out of date (CI)
//
// The pure builders are exported so providers.gen.test.ts can assert the
// committed files match without shelling out.

const fs = require("fs");
const path = require("path");

const ROOT = path.resolve(__dirname, "..", "..");
const PATHS = {
  source: path.join(ROOT, "shared", "providers.json"),
  ts: path.join(ROOT, "officejs", "src", "core", "providers.generated.ts"),
  worker: path.join(ROOT, "officejs", "proxy", "worker.js"),
};

const WORKER_BEGIN = "  // <providers:generated>";
const WORKER_END = "  // </providers:generated>";

function loadProviders(file = PATHS.source) {
  return JSON.parse(fs.readFileSync(file, "utf8")).providers;
}

/** The TS catalog consumed by providers.ts (types/functions stay hand-written there). */
function buildProvidersTs(providers) {
  const entries = providers
    .map(
      (p) =>
        `  ${p.id}: {\n` +
        `    id: ${q(p.id)}, label: ${q(p.label)},\n` +
        `    defaultBaseUrl: ${q(p.defaultBaseUrl)},\n` +
        `    requiresKey: ${p.requiresKey}, style: ${q(p.style)}, browserFriendly: ${p.browserFriendly},\n` +
        `  },`
    )
    .join("\n");
  return (
    "// AUTO-GENERATED from shared/providers.json by tools/gen-providers.cjs.\n" +
    "// Do not edit by hand — run `npm run gen:providers` to regenerate.\n" +
    'import type { ProviderSpec } from "./providers";\n' +
    "\n" +
    "export const PROVIDERS: Record<string, ProviderSpec> = {\n" +
    entries +
    "\n};\n"
  );
}

/** The lines that live between the markers in worker.js's PROVIDERS object. */
function buildWorkerBlock(providers) {
  return providers
    .map(
      (p) =>
        `  ${p.id}: { baseUrl: ${q(p.defaultBaseUrl)}, style: ${q(p.style)}, keyEnv: ${
          p.keyEnv == null ? "null" : q(p.keyEnv)
        } },`
    )
    .join("\n");
}

/** Splice a freshly-built block between the markers in the worker source. */
function applyWorkerBlock(workerSrc, block) {
  const begin = workerSrc.indexOf(WORKER_BEGIN);
  const end = workerSrc.indexOf(WORKER_END);
  if (begin === -1 || end === -1 || end < begin) {
    throw new Error("worker.js is missing the <providers:generated> markers");
  }
  const before = workerSrc.slice(0, begin + WORKER_BEGIN.length);
  const after = workerSrc.slice(end);
  return `${before}\n${block}\n${after}`;
}

function q(s) {
  return JSON.stringify(s);
}

function computeOutputs() {
  const providers = loadProviders();
  const ts = buildProvidersTs(providers);
  const workerSrc = fs.readFileSync(PATHS.worker, "utf8");
  const worker = applyWorkerBlock(workerSrc, buildWorkerBlock(providers));
  return { ts, worker };
}

function main() {
  const check = process.argv.includes("--check");
  const { ts, worker } = computeOutputs();
  const targets = [
    { file: PATHS.ts, next: ts },
    { file: PATHS.worker, next: worker },
  ];
  let stale = false;
  for (const t of targets) {
    const current = fs.existsSync(t.file) ? fs.readFileSync(t.file, "utf8") : "";
    if (current === t.next) continue;
    stale = true;
    if (check) {
      console.error(`Out of date: ${path.relative(ROOT, t.file)} — run \`npm run gen:providers\``);
    } else {
      fs.writeFileSync(t.file, t.next);
      console.log(`Wrote ${path.relative(ROOT, t.file)}`);
    }
  }
  if (check && stale) process.exit(1);
  if (!check && !stale) console.log("Provider tables already up to date.");
}

module.exports = { loadProviders, buildProvidersTs, buildWorkerBlock, applyWorkerBlock, computeOutputs, PATHS };

if (require.main === module) main();
