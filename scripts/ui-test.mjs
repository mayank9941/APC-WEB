// Interactive UI test of the search screen + report page.
import puppeteer from "puppeteer-core";

const browser = await puppeteer.launch({
  executablePath:
    "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe",
  headless: true,
  args: ["--window-size=1280,900"],
});
const page = await browser.newPage();
await page.setViewport({ width: 1280, height: 900 });

const results = [];
const check = (name, cond) => {
  results.push(`${cond ? "PASS" : "FAIL"}: ${name}`);
};

await page.goto("http://localhost:3000/", { waitUntil: "networkidle0" });

// 1. Focus with empty query -> full dropdown, alphabetical
await page.click(".search-box input");
await page.waitForSelector(".search-dropdown .option");
let names = await page.$$eval(".search-dropdown .option .name", (els) =>
  els.map((e) => e.textContent)
);
check("dropdown opens on focus with all plazas (>1000)", names.length > 1000);
const lower = names.map((n) => n.toLowerCase());
check(
  "dropdown sorted alphabetically (case-insensitive)",
  lower.every((n, i) => i === 0 || lower[i - 1] <= n)
);
check(
  "options rendered as Name (code)",
  /\(\d+\)$/.test(names[0].trim())
);

// 2. Case-insensitive name filter
await page.type(".search-box input", "hAtHi");
await new Promise((r) => setTimeout(r, 300));
names = await page.$$eval(".search-dropdown .option .name", (els) =>
  els.map((e) => e.textContent)
);
check(
  "case-insensitive name search 'hAtHi' finds Hathitala",
  names.some((n) => n.includes("Hathitala"))
);

// 3. Code search
await page.$eval(".search-box input", (el) => {
  const setter = Object.getOwnPropertyDescriptor(
    window.HTMLInputElement.prototype,
    "value"
  ).set;
  setter.call(el, "");
  el.dispatchEvent(new Event("input", { bubbles: true }));
});
await page.type(".search-box input", "330085");
await new Promise((r) => setTimeout(r, 300));
names = await page.$$eval(".search-dropdown .option .name", (els) =>
  els.map((e) => e.textContent)
);
check(
  "code search '330085' finds Hathitala",
  names.length >= 1 && names.some((n) => n.includes("(330085)"))
);

// 4. Enter selects and navigates to report
await page.keyboard.press("Enter");
await page.waitForFunction(
  () => document.querySelector(".report-header h1, .error-box"),
  { timeout: 30000 }
);
await page.waitForFunction(
  () => !document.body.textContent.includes("Loading plaza data"),
  { timeout: 30000 }
);
check(
  "Enter navigates to plaza report",
  page.url().includes("/plaza/330085")
);
const h1 = await page.$eval(".report-header h1", (e) => e.textContent);
check("report shows plaza name", h1.includes("Hathitala"));

// 5. Change growth to 5% -> Table 3 row 6 shows 5, APC changes
const apcBefore = await page.$eval(".apc-hero .value", (e) => e.textContent);
await page.select("#growth", "5");
await new Promise((r) => setTimeout(r, 400));
const apcAfter = await page.$eval(".apc-hero .value", (e) => e.textContent);
check(
  `growth 0->5 recalculates APC live (${apcBefore} -> ${apcAfter})`,
  apcBefore !== apcAfter
);

// 6. Add MF entry 2025-10-20 factor 1.05 -> Oct-25 multiplier 1.0306
await page.evaluate(() => {
  const btns = [...document.querySelectorAll(".btn")];
  btns.find((b) => b.textContent.includes("Add MF")).click();
});
await page.waitForSelector(".mf-row input[type=date]");
await page.focus(".mf-row input[type=date]");
await page.$eval(".mf-row input[type=date]", (el) => {
  const setter = Object.getOwnPropertyDescriptor(
    window.HTMLInputElement.prototype,
    "value"
  ).set;
  setter.call(el, "2025-10-20");
  el.dispatchEvent(new Event("input", { bubbles: true }));
});
await page.$eval(".mf-row input[type=number]", (el) => {
  const setter = Object.getOwnPropertyDescriptor(
    window.HTMLInputElement.prototype,
    "value"
  ).set;
  setter.call(el, "1.05");
  el.dispatchEvent(new Event("input", { bubbles: true }));
});
await new Promise((r) => setTimeout(r, 400));
const octMult = await page.evaluate(() => {
  const rows = [...document.querySelectorAll("table.apc tbody tr")];
  const oct = rows.find((r) => r.textContent.startsWith("Oct-2025"));
  return oct ? oct.children[4].textContent : null;
});
check(
  `MF 1.05 eff 2025-10-20 gives Oct-25 multiplier 1.0306 (got ${octMult})`,
  octMult === "1.0306"
);

// 7. Change "up to" month -> subtitle window updates
await page.select("#upto", "2025-12");
await new Promise((r) => setTimeout(r, 400));
const subtitle = await page.$eval(".report-subtitle", (e) => e.textContent);
check(
  `month selector changes window (${subtitle.trim().slice(0, 80)})`,
  subtitle.includes("Dec-2025") && subtitle.includes("Jan-2025")
);

await page.screenshot({ path: "../search-report-final.png", fullPage: false });
await browser.close();

console.log(results.join("\n"));
const failed = results.filter((r) => r.startsWith("FAIL"));
process.exit(failed.length ? 1 : 0);
