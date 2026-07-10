/**
 * Generate AppSource store screenshots (1366x768) from the live dev server:
 * the landing page and the configured task pane (via the harness). Run with the
 * dev server up. Outputs to store/screenshots/.
 *   HARNESS_URL=https://localhost:3999 node tools/gen-screenshots.cjs
 */
const { chromium } = require("playwright-core");
const path = require("path");
const fs = require("fs");

const BASE = process.env.HARNESS_URL || "https://localhost:3999";
const OUT = path.resolve(__dirname, "..", "store", "screenshots");
const KEY = process.env.OPENROUTER_API_KEY || "";

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  const browser = await chromium.launch({ channel: "chrome", headless: true });
  const ctx = await browser.newContext({ ignoreHTTPSErrors: true, viewport: { width: 1366, height: 768 } });

  // 1) Landing page (what it does).
  const p1 = await ctx.newPage();
  await p1.goto(`${BASE}/index.html`, { waitUntil: "networkidle" });
  await p1.screenshot({ path: path.join(OUT, "01-overview.png") });

  // 2) Task pane, configured, on a neutral canvas.
  const p2 = await ctx.newPage();
  await p2.goto(`${BASE}/harness.html`, { waitUntil: "domcontentloaded" });
  await p2.waitForFunction(() => document.querySelectorAll("#provider option").length > 0, { timeout: 15000 });
  await p2.selectOption("#provider", "openrouter");
  await p2.fill("#model", "openai/gpt-4o-mini");
  if (KEY) await p2.fill("#apiKey", KEY);
  await p2.fill("#agentInput", "In D1 put the sum of B2:B10, then bold anything over 100");
  await p2.screenshot({ path: path.join(OUT, "02-taskpane.png") });

  await browser.close();
  console.log("wrote", fs.readdirSync(OUT).join(", "), "to store/screenshots/");
})().catch((e) => {
  console.error("SCREENSHOT ERROR:", e.message);
  process.exit(1);
});
