/* End-to-end check of the built page in a real browser.
   The CDN hosts are unreachable from this sandbox, so the two library tags are
   rewritten to local copies of the exact same versions. */
const fs = require("fs");
const path = require("path");
const http = require("http");
const { chromium } = require("playwright");

const ROOT = path.resolve("e2eroot");
fs.rmSync(ROOT, { recursive: true, force: true });
fs.mkdirSync(ROOT, { recursive: true });
fs.copyFileSync("node_modules/xlsx/dist/xlsx.full.min.js", path.join(ROOT, "xlsx.js"));
fs.copyFileSync("node_modules/exceljs/dist/exceljs.min.js", path.join(ROOT, "exceljs.js"));
fs.writeFileSync(path.join(ROOT, "index.html"),
  fs.readFileSync(path.resolve(__dirname, "../web/index.html"), "utf8")
    .replace(/https:\/\/cdn\.jsdelivr\.net\/npm\/xlsx@[^"]+/, "xlsx.js")
    .replace(/https:\/\/cdn\.jsdelivr\.net\/npm\/exceljs@[^"]+/, "exceljs.js"));

const TYPES = { ".html": "text/html", ".js": "text/javascript" };
const server = http.createServer((req, res) => {
  const f = path.join(ROOT, req.url === "/" ? "index.html" : req.url.split("?")[0]);
  if (!f.startsWith(ROOT) || !fs.existsSync(f)) { res.writeHead(404); return res.end(); }
  res.writeHead(200, { "content-type": TYPES[path.extname(f)] || "application/octet-stream" });
  res.end(fs.readFileSync(f));
});

const VIEWS = ["overview", "master", "people", "sources", "demand", "unassigned",
               "recon", "exceptions", "history"];
const fail = [];
const ok = m => console.log("  ok   " + m);
const bad = m => { fail.push(m); console.log("  FAIL " + m); };

(async () => {
  await new Promise(r => server.listen(0, r));
  const url = "http://127.0.0.1:" + server.address().port + "/";
  const browser = await chromium.launch({ executablePath: "/opt/pw-browsers/chromium" });

  for (const scheme of ["light", "dark"]) {
    const ctx = await browser.newContext({ colorScheme: scheme, viewport: { width: 1440, height: 950 },
                                           acceptDownloads: true });
    const page = await ctx.newPage();
    const errors = [];
    page.on("pageerror", e => errors.push("pageerror: " + e.message));
    page.on("console", m => { if (m.type() === "error") errors.push("console: " + m.text()); });

    console.log(`\n── ${scheme} ───────────────────────────────`);
    await page.goto(url, { waitUntil: "load" });
    await page.waitForTimeout(400);

    // 1. sample week
    await page.click("#demo");
    await page.waitForSelector("#v-overview:not([hidden])", { timeout: 15000 });
    const kpis = await page.locator("#ovKpis .kpi").count();
    kpis === 6 ? ok(`sample week renders ${kpis} KPI tiles`) : bad(`KPI tiles: ${kpis}`);
    const findings = await page.locator("#insights .find").count();
    findings > 0 ? ok(`${findings} findings generated`) : bad("no findings generated");
    (await page.locator("#sampleBanner").isVisible()) ? ok("sample is labelled as sample")
                                                      : bad("sample banner hidden");

    // 2. every view renders rows
    for (const v of VIEWS) {
      await page.click(`#nav button[data-view="${v}"]`);
      await page.waitForSelector(`#v-${v}:not([hidden])`);
      const rows = await page.locator(`#v-${v} table tbody tr`).count();
      const blocks = await page.locator(`#v-${v} .rep, #v-${v} .fn-row, #v-${v} .find`).count();
      if (v === "history" || rows + blocks > 0) ok(`${v}: ${rows} rows, ${blocks} blocks`);
      else bad(`${v} rendered nothing`);
    }

    // 3. source filter
    await page.click('#nav button[data-view="overview"]');
    const srcBtns = page.locator("#srcFilter button");
    if (await srcBtns.count() > 1) {
      await srcBtns.nth(1).click();
      await page.waitForTimeout(150);
      const scope = await page.locator("#fnScope").textContent();
      scope && scope !== "All sources" ? ok(`source filter -> ${scope}`) : bad("source filter did nothing");
      await srcBtns.nth(0).click();
    } else bad("no source filter buttons");

    // 4. exception severity filter
    await page.click('#nav button[data-view="exceptions"]');
    const sevBtns = page.locator("#exFilter button");
    if (await sevBtns.count() > 1) {
      await sevBtns.nth(1).click();
      await page.waitForTimeout(120);
      ok(`severity filter -> ${await page.locator("#tExceptions tbody tr").count()} rows`);
    } else bad("no severity filter buttons");

    // 5. Excel download from the sample
    await page.click('#nav button[data-view="master"]');
    const [dl] = await Promise.all([
      page.waitForEvent("download", { timeout: 30000 }),
      page.locator("#v-master .dl").click(),
    ]);
    const out = path.resolve(`e2e_${scheme}.xlsx`);
    await dl.saveAs(out);
    const kb = fs.statSync(out).size / 1024;
    /^Master_Lead_Report_\d{4}-\d{2}-\d{2}_to_\d{4}-\d{2}-\d{2}\.xlsx$/.test(dl.suggestedFilename())
      ? ok(`downloaded ${dl.suggestedFilename()} (${kb.toFixed(0)} KB)`)
      : bad("download filename: " + dl.suggestedFilename());

    // 6. a real upload, saved to history, then reopened
    await page.click('#nav button[data-view="upload"]');
    await page.setInputFiles("#file", ["sample_meta.xlsx", "sample_indiamart.xlsx",
                                       "sample_justdial.xlsx"]);
    await page.waitForTimeout(200);
    await page.click("#run");
    await page.waitForSelector("#v-overview:not([hidden])", { timeout: 30000 });
    const label = await page.locator("#runLabel").textContent();
    /2026-09-14 → 2026-09-20 · 367 leads/.test(label)
      ? ok(`upload reconciled: ${label}`) : bad(`run label: ${label}`);
    (await page.locator("#sampleBanner").isVisible()) ? bad("sample banner still showing")
                                                      : ok("sample banner cleared");

    await page.waitForTimeout(600);
    await page.click('#nav button[data-view="history"]');
    const saved = await page.locator("#runs .runrow").count();
    saved >= 1 ? ok(`${saved} run(s) in history`) : bad("run was not saved to history");

    // reopen it
    await page.locator("#runs [data-open]").first().click();
    await page.waitForSelector("#v-overview:not([hidden])", { timeout: 20000 });
    const label2 = await page.locator("#runLabel").textContent();
    label2 === label ? ok("saved run reopens identically") : bad(`reopened as ${label2}`);

    // 7. balance badge
    await page.click('#nav button[data-view="recon"]');
    const balanced = await page.locator("#balance .balance:not(.off)").count();
    balanced ? ok("control totals balance") : bad("reconciliation reports out of balance");

    // 8. phone width
    await page.setViewportSize({ width: 390, height: 844 });
    await page.click('#nav button[data-view="overview"]');
    await page.waitForTimeout(250);
    const over = await page.evaluate(() =>
      document.documentElement.scrollWidth - document.documentElement.clientWidth);
    over <= 1 ? ok("no horizontal page scroll at 390px") : bad(`page overflows by ${over}px at 390px`);
    await page.setViewportSize({ width: 1440, height: 950 });

    if (scheme === "light") {
      await page.click('#nav button[data-view="overview"]');
      await page.waitForTimeout(250);
      await page.screenshot({ path: "shot_overview.png", fullPage: false });
      await page.click('#nav button[data-view="people"]');
      await page.waitForTimeout(250);
      await page.screenshot({ path: "shot_people.png", fullPage: false });
    } else {
      await page.click('#nav button[data-view="overview"]');
      await page.waitForTimeout(250);
      await page.screenshot({ path: "shot_dark.png", fullPage: false });
    }

    errors.length ? bad(`${errors.length} console/page error(s): ` + errors.slice(0, 4).join(" | "))
                  : ok("no console or page errors");
    await ctx.close();
  }

  await browser.close();
  server.close();
  console.log("\n" + (fail.length ? `${fail.length} FAILURE(S)` : "ALL CHECKS PASSED"));
  process.exit(fail.length ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
