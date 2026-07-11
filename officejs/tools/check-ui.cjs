// Headless UI smoke check for the new task-pane bits (no LLM call). Loads the
// harness in Chrome and asserts: page loads without console errors, the prompt
// library renders chips, clicking one fills the agent input, and the usage meter
// line is present. Requires the dev server running + Google Chrome.
const { chromium } = require("playwright-core");

const URL = process.env.HARNESS_URL || "https://localhost:3000/harness.html";

(async () => {
  const browser = await chromium.launch({ channel: "chrome", headless: true });
  const ctx = await browser.newContext({ ignoreHTTPSErrors: true });
  const page = await ctx.newPage();
  const errors = [];
  // Ignore incidental resource 404s (favicon, source maps) — only real JS errors matter.
  page.on("console", (m) => {
    if (m.type() === "error" && !/Failed to load resource/.test(m.text())) errors.push(m.text());
  });
  page.on("pageerror", (e) => errors.push(String(e)));

  await page.goto(URL, { waitUntil: "networkidle" });
  await page.waitForTimeout(500);

  const chipCount = await page.locator("#agentPresets button.chip").count();
  const usageText = (await page.locator("#usage").textContent()) || "";

  // Click the first preset and confirm it fills the agent input.
  await page.locator("#agentPresets button.chip").first().click();
  const filled = await page.inputValue("#agentInput");

  await browser.close();

  const fails = [];
  if (errors.length) fails.push("console errors: " + errors.join(" | "));
  if (chipCount < 3) fails.push(`expected >=3 preset chips, got ${chipCount}`);
  if (!/reset/.test(usageText)) fails.push("usage meter missing");
  if (!filled || filled.length < 5) fails.push("clicking a preset did not fill the agent input");

  if (fails.length) {
    console.error("UI CHECK FAILED:\n - " + fails.join("\n - "));
    process.exit(1);
  }
  console.log(`UI CHECK OK: ${chipCount} chips, preset fills input, usage meter present.`);
})().catch((e) => {
  console.error("UI CHECK ERROR:", e.message);
  process.exit(1);
});
