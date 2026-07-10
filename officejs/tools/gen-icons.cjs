/**
 * Generate the add-in icons (assets/icon-*.png) from an inline SVG, rendered by
 * headless Chrome via playwright-core. Run: `node tools/gen-icons.cjs`.
 * Reproducible — edit the SVG here and re-run to refresh all sizes.
 */
const { chromium } = require("playwright-core");
const path = require("path");

const SIZES = [16, 32, 64, 80, 128];
const OUT = path.resolve(__dirname, "..", "assets");

// Excel-green rounded tile + white chat bubble + green spark = "AI in cells".
const SVG = `
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 128 128">
  <defs>
    <linearGradient id="g" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0" stop-color="#27AE60"/>
      <stop offset="1" stop-color="#0E7C42"/>
    </linearGradient>
  </defs>
  <rect x="6" y="6" width="116" height="116" rx="26" fill="url(#g)"/>
  <rect x="30" y="28" width="68" height="52" rx="14" fill="#ffffff"/>
  <path d="M46 78 L46 96 L64 80 Z" fill="#ffffff"/>
  <path d="M64 37 L69 49 L81 54 L69 59 L64 71 L59 59 L47 54 L59 49 Z" fill="#0E7C42"/>
</svg>`;

(async () => {
  const browser = await chromium.launch({ channel: "chrome", headless: true });
  const page = await browser.newPage();
  for (const s of SIZES) {
    await page.setViewportSize({ width: s, height: s });
    await page.setContent(
      `<html><body style="margin:0;padding:0">
         <svg width="${s}" height="${s}" viewBox="0 0 128 128">${SVG.replace(/<\/?svg[^>]*>/g, "")}</svg>
       </body></html>`
    );
    const el = await page.$("svg");
    await el.screenshot({ path: path.join(OUT, `icon-${s}.png`), omitBackground: true });
    console.log("wrote", `assets/icon-${s}.png`, `(${s}x${s})`);
  }
  await browser.close();
})().catch((e) => {
  console.error("ICON GEN ERROR:", e.message);
  process.exit(1);
});
