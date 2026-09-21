const fs=require("fs"),path=require("path"),XLSX=require("xlsx"),ExcelJS=require("exceljs");
const E=require("../web-src/engine.js"),R=require("../web-src/xlsxread.js"),Rep=require("../web-src/report.js");
const cfg=JSON.parse(fs.readFileSync(__dirname+"/../server/config/sources.json","utf8"));
const parsed=process.argv.slice(2).map(p=>E.readFile(R.readWorkbook(XLSX,new Uint8Array(fs.readFileSync(p))),path.basename(p),cfg));
const rec=E.reconcile(parsed,cfg), m=E.compute(rec,cfg), ins=E.buildInsights(m,rec,cfg);
Rep.build(ExcelJS,rec,m,ins,cfg).then(buf=>{fs.writeFileSync("out_js.xlsx",Buffer.from(buf));console.log("wrote out_js.xlsx");});
