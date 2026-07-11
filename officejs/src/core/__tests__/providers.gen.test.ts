// Guard: keep every edition's provider table in lock-step with the single source
// of truth (shared/providers.json). This is the test that would have caught the
// Nebius base-URL drift (VBA had ...nebius.ai while TS/proxy had ...nebius.com).

import * as fs from "fs";
import * as path from "path";
import { PROVIDERS } from "../providers";

// tools/gen-providers.cjs is CommonJS; require it for its pure builders.
// eslint-disable-next-line @typescript-eslint/no-var-requires
const gen = require("../../../tools/gen-providers.cjs");
const providers = gen.loadProviders() as Array<Record<string, any>>;
const read = (p: string) => fs.readFileSync(p, "utf8");

describe("provider tables are generated from shared/providers.json", () => {
  test("TS PROVIDERS matches the source (id/url/style/flags)", () => {
    expect(Object.keys(PROVIDERS).sort()).toEqual(providers.map((p) => p.id).sort());
    for (const p of providers) {
      expect(PROVIDERS[p.id]).toEqual({
        id: p.id,
        label: p.label,
        defaultBaseUrl: p.defaultBaseUrl,
        requiresKey: p.requiresKey,
        style: p.style,
        browserFriendly: p.browserFriendly,
      });
    }
  });

  test("committed providers.generated.ts is up to date", () => {
    expect(read(gen.PATHS.ts)).toBe(gen.buildProvidersTs(providers));
  });

  test("committed worker.js provider block is up to date", () => {
    // computeOutputs() re-splices the block into the current worker source and
    // returns what the file *should* be; it must equal the committed file.
    expect(read(gen.PATHS.worker)).toBe(gen.computeOutputs().worker);
  });
});

describe("VBA edition agrees with the source", () => {
  const ROOT = path.resolve(__dirname, "..", "..", "..", "..");
  const modConfig = read(path.join(ROOT, "modConfig.bas"));
  const modMenu = read(path.join(ROOT, "modMenu.bas"));

  test("every provider's base URL appears verbatim in modConfig.bas GetBaseURL/InitializeDefaults", () => {
    // Ollama's URL is user-configurable in VBA, so only assert the fixed ones.
    for (const p of providers) {
      if (p.id === "ollama") continue;
      expect(modConfig.includes(p.defaultBaseUrl)).toBe(true);
    }
  });

  test("modMenu.bas suggests only canonical provider base URLs", () => {
    // The menu suggests default URLs. Match any URL that mentions a provider's
    // id token (e.g. "nebius") and require it to equal the canonical URL — so a
    // .ai↔.com host swap is caught even though the host itself changed.
    const urls = modMenu.match(/https?:\/\/[^"'\s]+/g) || [];
    for (const p of providers) {
      if (p.id === "ollama") continue;
      for (const u of urls) {
        if (u.includes(p.id)) expect(u).toBe(p.defaultBaseUrl);
      }
    }
  });
});
