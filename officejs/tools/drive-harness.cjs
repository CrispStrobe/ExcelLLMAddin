/**
 * Headless smoke test for the dev harness — drives the task pane in Chrome and
 * runs a REAL provider round-trip, no Excel. Proves the pane + core + fetch path.
 *
 * Prereqs:
 *   - a dev server running the harness:  npm start   (serves /harness.html on :3000)
 *   - Google Chrome installed (used via playwright-core channel: "chrome")
 *   - OPENROUTER_API_KEY in the environment
 *
 * Usage:
 *   OPENROUTER_API_KEY=sk-or-... npm run harness:smoke
 *   HARNESS_URL=https://localhost:3999/harness.html OPENROUTER_API_KEY=... node tools/drive-harness.cjs
 *
 * The key is only typed into the page; it is never printed.
 */
const { chromium } = require("playwright-core");

const URL = process.env.HARNESS_URL || "https://localhost:3000/harness.html";
const KEY = process.env.OPENROUTER_API_KEY || "";
const MODEL = process.env.HARNESS_MODEL || "openai/gpt-4o-mini";

(async () => {
  if (!KEY) throw new Error("Set OPENROUTER_API_KEY in the environment.");
  const browser = await chromium.launch({ channel: "chrome", headless: true });
  const ctx = await browser.newContext({ ignoreHTTPSErrors: true });
  const page = await ctx.newPage();
  const errs = [];
  page.on("pageerror", (e) => errs.push(e.message));

  await page.goto(URL, { waitUntil: "domcontentloaded" });
  await page.waitForFunction(() => document.querySelectorAll("#provider option").length > 0, { timeout: 15000 });

  await page.selectOption("#provider", "openrouter");
  await page.fill("#apiKey", KEY);
  await page.fill("#model", MODEL);
  await page.click("#save");

  await page.click("#test");
  await page.waitForFunction(
    () => { const s = document.getElementById("status"); return s && s.textContent && !/Testing/i.test(s.textContent); },
    { timeout: 45000 }
  );
  const testStatus = (await page.textContent("#status")) || "";

  await page.click("#refreshModels");
  await page.waitForFunction(
    () => { const s = document.getElementById("status"); return s && /(\d+\s*model|Error)/i.test(s.textContent || ""); },
    { timeout: 45000 }
  );
  const listStatus = (await page.textContent("#status")) || "";
  await page.fill("#modelFilter", "claude");
  const filtered = await page.$$eval("#modelSelect option", (os) => os.length);

  await browser.close();

  console.log("TEST  :", testStatus);
  console.log("LIST  :", listStatus);
  console.log("FILTER:", filtered, "models match 'claude'");
  if (errs.length) console.log("ERRORS:", errs.join(" | "));

  const ok = /^OK:/.test(testStatus) && /\d+\s*model/i.test(listStatus) && filtered > 0 && errs.length === 0;
  if (!ok) {
    console.error("SMOKE FAILED");
    process.exit(1);
  }
  console.log("SMOKE OK");
})().catch((e) => {
  console.error("DRIVER ERROR:", e.message);
  process.exit(1);
});
