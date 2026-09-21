/* Turns a workbook into the plain cell matrices the engine expects.
   Dates come back as local-midnight Date objects with no timezone shift,
   which is what openpyxl handed the Python engine. */
(function (root) {
"use strict";

const MAX_COLS = 80;

function cellValue(XLSX, cell) {
  if (!cell || cell.t === "z") return null;
  if (cell.t === "d") {
    const d = cell.v instanceof Date ? cell.v : new Date(cell.v);
    return isNaN(d) ? null : d;
  }
  if (cell.t === "n") {
    const fmt = cell.z;
    if (fmt && XLSX.SSF.is_date(fmt)) {
      const p = XLSX.SSF.parse_date_code(cell.v);
      if (p && p.y) return new Date(p.y, p.m - 1, p.d, p.H || 0, p.M || 0, Math.floor(p.S || 0));
    }
    return cell.v;
  }
  if (cell.t === "e") return null;
  return cell.v === undefined ? null : cell.v;
}

/** ArrayBuffer/Uint8Array -> [{name, rows}] */
function readWorkbook(XLSX, data) {
  const wb = XLSX.read(data, { type: "array", cellNF: true, cellDates: false,
                               cellStyles: false, sheetStubs: false });
  return wb.SheetNames.map(name => {
    const ws = wb.Sheets[name];
    const rows = [];
    if (ws && ws["!ref"]) {
      const range = XLSX.utils.decode_range(ws["!ref"]);
      const lastCol = Math.min(range.e.c, range.s.c + MAX_COLS - 1);
      for (let R = range.s.r; R <= range.e.r; R++) {
        const row = [];
        for (let Cc = range.s.c; Cc <= lastCol; Cc++) {
          row.push(cellValue(XLSX, ws[XLSX.utils.encode_cell({ r: R, c: Cc })]));
        }
        // openpyxl indexes from column A; pad if the sheet starts further right
        for (let k = 0; k < range.s.c; k++) row.unshift(null);
        rows.push(row);
      }
    }
    return { name, rows };
  });
}

const API = { readWorkbook };
root.OakXlsx = API;
if (typeof module !== "undefined" && module.exports) module.exports = API;

})(typeof globalThis !== "undefined" ? globalThis : this);
