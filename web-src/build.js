/* Assembles web/index.html from the sources in this folder.
   Run from anywhere:  node web-src/build.js                          */
const fs = require("fs");
const path = require("path");

const HERE = __dirname;
const ROOT = path.resolve(HERE, "..");
const read = p => fs.readFileSync(path.resolve(HERE, p), "utf8");
const guard = s => s.replace(/<\/script/gi, "<\\/script");

// The engine's source of truth is the Flask app's config file: one profile
// list, both builds.
const cfg = JSON.stringify(JSON.parse(
  fs.readFileSync(path.join(ROOT, "server", "config", "sources.json"), "utf8")));

const CDN_XLSX  = "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js";
const CDN_EXCEL = "https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js";

const head = `<title>Lead Reconciler</title>\n<style>\n${read("style.css")}</style>`;
const body = read("body.html");
const scripts = `
<script src="${CDN_XLSX}"></script>
<script src="${CDN_EXCEL}"></script>
<script>window.OAK_CONFIG=${guard(cfg)};</script>
<script>${guard(read("engine.js"))}</script>
<script>${guard(read("xlsxread.js"))}</script>
<script>${guard(read("report.js"))}</script>
<script>${guard(read("app.js"))}</script>`;

const full = `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<link rel="icon" href="data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'><rect width='32' height='32' rx='7' fill='%230e4a52'/><path d='M7 9h18l-7 8v7l-4 2v-9z' fill='%23b4762e'/></svg>">
${head}
</head>
<body>
${body}${scripts}
</body>
</html>
`;

fs.mkdirSync(path.join(ROOT, "web"), { recursive: true });
fs.writeFileSync(path.join(ROOT, "web", "index.html"), full);

// Same page without the document skeleton, for hosts that supply their own.
fs.mkdirSync(path.join(HERE, "build"), { recursive: true });
fs.writeFileSync(path.join(HERE, "build", "embed.html"), `${head}\n${body}${scripts}\n`);

const kb = f => (fs.statSync(f).size / 1024).toFixed(0) + " KB";
console.log("web/index.html            ", kb(path.join(ROOT, "web", "index.html")));
console.log("web-src/build/embed.html  ", kb(path.join(HERE, "build", "embed.html")));
