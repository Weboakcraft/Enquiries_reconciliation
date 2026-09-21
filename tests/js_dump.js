/* Runs the JS engine over the same workbooks and writes out_js.json. */
const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");
const E = require("../web-src/engine.js");
const R = require("../web-src/xlsxread.js");

const cfg = JSON.parse(fs.readFileSync(__dirname + "/../server/config/sources.json", "utf8"));
const files = process.argv.slice(2);

const parsed = files.map(p => {
  const buf = fs.readFileSync(p);
  const sheets = R.readWorkbook(XLSX, new Uint8Array(buf));
  return E.readFile(sheets, path.basename(p), cfg);
});

const rec = E.reconcile(parsed, cfg);
const m = E.compute(rec, cfg);
const ins = E.buildInsights(m, rec, cfg);

fs.writeFileSync("out_js.json",
  JSON.stringify({ reconciliation: rec, metrics: m, insights: ins }, null, 1));
console.log("rows", JSON.stringify(rec.control_totals));
console.log("period", JSON.stringify(rec.period));
console.log("leads", rec.leads.length);
