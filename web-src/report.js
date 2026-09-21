/* ===========================================================================
   Master workbook builder — a port of core/excel.py onto ExcelJS, so the
   "Download Excel report" button produces the same file without a server.
   =========================================================================== */
(function (root) {
"use strict";

const BAR = "2E75B6", NAVY = "1F4E79", PINK = "FFC7CE", PINKT = "9C0006";
const CREAM = "FFF2CC", CREAMT = "C55A11", GREY = "6B7684";
const GREEN = "1E7A4B", GREENL = "DCF0E4", RED = "B3261E", REDL = "FBE3E1";
const AMBER = "B26B00", AMBERL = "FDF0DC", A = "Arial";

const col = hex => ({ argb: "FF" + hex });
const font = (size, opts) => Object.assign({ name: A, size }, opts || {});

const F_T    = font(16, { bold: true, color: col(NAVY) });
const F_ST   = font(9,  { color: col(GREY), italic: true });
const F_BAR  = font(11, { bold: true, color: col("FFFFFF") });
const F_H    = font(10, { bold: true, color: col("FFFFFF") });
const F_B    = font(10);
const F_BB   = font(10, { bold: true });
const F_N    = font(9,  { color: col(GREY) });
const F_PINK = font(10, { bold: true, color: col(PINKT) });
const F_TOT  = font(10, { bold: true, color: col(CREAMT) });

const fill = hex => ({ type: "pattern", pattern: "solid", fgColor: col(hex) });
const FL_BAR = fill(BAR), FL_H = fill(NAVY), FL_PINK = fill(PINK), FL_CREAM = fill(CREAM);
const FL_G = fill(GREENL), FL_R = fill(REDL), FL_A = fill(AMBERL);

const _side = { style: "thin", color: col("9EB6CE") };
const BORD = { left: _side, right: _side, top: _side, bottom: _side };
const C = { horizontal: "center", vertical: "middle" };
const L = { horizontal: "left", vertical: "middle" };
const W = { horizontal: "center", vertical: "middle", wrapText: true };
const LW = { horizontal: "left", vertical: "middle", wrapText: true };

const INT = '#,##0;-#,##0;"-"', PCT = '0.0%;-0.0%;"-"', PLAIN = "#,##0", DMY = "DD-MMM-YY";

/** Excel serial dates are UTC-based, so build the Date in UTC to avoid a shift. */
const xlDate = isoStr => new Date(Date.UTC(+isoStr.slice(0, 4), +isoStr.slice(5, 7) - 1,
                                           +isoStr.slice(8, 10)));

function _title(ws, t, sub, width) {
  ws.getCell(1, 1).value = t;   ws.getCell(1, 1).font = F_T;
  ws.getCell(2, 1).value = sub; ws.getCell(2, 1).font = F_ST;
  ws.mergeCells(1, 1, 1, width);
  ws.mergeCells(2, 1, 2, width);
  ws.getRow(1).height = 24;
  ws.getRow(2).height = 14;
}

function _bar(ws, row, text, width) {
  for (let i = 1; i <= width; i++) ws.getCell(row, i).fill = FL_BAR;
  const c = ws.getCell(row, 1);
  c.value = text; c.font = F_BAR; c.alignment = L;
  ws.mergeCells(row, 1, row, width);
  ws.getRow(row).height = 20;
}

function _hdr(ws, row, vals, h) {
  vals.forEach((v, i) => {
    const c = ws.getCell(row, 1 + i);
    c.value = v; c.font = F_H; c.fill = FL_H; c.alignment = W; c.border = BORD;
  });
  ws.getRow(row).height = h || 30;
}

function _put(ws, r, c, v, fmt, f, align, fl) {
  const cell = ws.getCell(r, c);
  cell.value = (v === undefined || v === "" ? null : v);
  cell.font = f || F_B;
  cell.alignment = align || C;
  cell.border = BORD;
  if (fmt) cell.numFmt = fmt;
  if (fl) cell.fill = fl;
  return cell;
}

function _widths(ws, widths) { widths.forEach((w, i) => { ws.getColumn(i + 1).width = w; }); }

const today = () => {
  const d = new Date();
  const MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  return `${String(d.getDate()).padStart(2, "0")} ${MON[d.getMonth()]} ${d.getFullYear()}`;
};

/* ------------------------------------------------------------------ build */

async function build(ExcelJS, rec, m, insights, cfg) {
  const C_ = cfg.canonical;
  const stages = m.stages, quals = m.qualities;
  const reps = m.reps, sources = m.sources;
  const per = `${m.period.start} to ${m.period.end}`;
  const o = m.overall;
  const wb = new ExcelJS.Workbook();

  /* ------------------------------------------------------ MASTER REPORT */
  let ws = wb.addWorksheet("MASTER REPORT");
  ws.views = [{ showGridLines: false }];
  _title(ws, "Master Lead Report",
         `All sources · ${per} · ${o.total} leads · generated ${today()}`, 7 + stages.length);

  _bar(ws, 4, `  EMPLOYEE WISE SNAPSHOT  (ALL SOURCES, ${per})`,
       4 + sources.length + stages.length);
  _hdr(ws, 5, ["Salesperson", "Total Enquiry"].concat(sources, stages, ["Conversion %", "Loss %"]));
  let r = 6;
  for (const rep of reps) {
    const un = rep === C_.unassigned_label;
    const fnt = un ? F_PINK : F_B, fl = un ? FL_PINK : null;
    const b = m.by_rep[rep];
    _put(ws, r, 1, rep, null, un ? F_PINK : F_BB, L, fl);
    _put(ws, r, 2, b.total, INT, un ? F_PINK : F_BB, C, fl);
    sources.forEach((s, i) => _put(ws, r, 3 + i, m.cube[rep][s].total, INT, fnt, C, fl));
    stages.forEach((s, i) => _put(ws, r, 3 + sources.length + i, b[s], INT, fnt, C, fl));
    const c0 = 3 + sources.length + stages.length;
    _put(ws, r, c0, b.conversion / 100, PCT, un ? F_PINK : F_BB, C, fl || FL_G);
    _put(ws, r, c0 + 1, b.loss_rate / 100, PCT, fnt, C, fl);
    if (!un) ws.getCell(r, 3 + sources.length).fill = FL_A;
    r++;
  }
  _put(ws, r, 1, "TOTAL", null, F_TOT, L, FL_CREAM);
  _put(ws, r, 2, o.total, INT, F_TOT, C, FL_CREAM);
  sources.forEach((s, i) => _put(ws, r, 3 + i, m.by_source[s].total, INT, F_TOT, C, FL_CREAM));
  stages.forEach((s, i) => _put(ws, r, 3 + sources.length + i, o[s], INT, F_TOT, C, FL_CREAM));
  let c0 = 3 + sources.length + stages.length;
  _put(ws, r, c0, o.conversion / 100, PCT, F_TOT, C, FL_CREAM);
  _put(ws, r, c0 + 1, o.loss_rate / 100, PCT, F_TOT, C, FL_CREAM);
  const end1 = r;

  const b2 = end1 + 3;
  _bar(ws, b2, "  SOURCE WISE SNAPSHOT", 10);
  _hdr(ws, b2 + 1, ["Source", "Total Enquiry", "% Share", "Assigned", "Unassigned",
                    "Won", "Closed no order", "Qualified", "Hot Leads", "Qty (pcs)"]);
  r = b2 + 2;
  for (const s of sources) {
    const b = m.by_source[s];
    _put(ws, r, 1, s, null, F_BB, L);
    _put(ws, r, 2, b.total, INT, F_BB);
    _put(ws, r, 3, o.total ? b.total / o.total : 0, PCT);
    _put(ws, r, 4, b.total - b.unassigned, INT);
    _put(ws, r, 5, b.unassigned, INT, b.unassigned ? F_PINK : F_B, C, b.unassigned ? FL_PINK : null);
    _put(ws, r, 6, b.Won, INT, F_B, C, FL_G);
    _put(ws, r, 7, b.closed_no_order, INT, F_B, C, FL_R);
    _put(ws, r, 8, b.Qualified, INT);
    _put(ws, r, 9, b.hot, INT, F_BB);
    _put(ws, r, 10, b.qty, PLAIN, F_BB);
    r++;
  }
  _put(ws, r, 1, "TOTAL", null, F_TOT, L, FL_CREAM);
  [[2, o.total], [4, o.total - o.unassigned], [5, o.unassigned], [6, o.Won],
   [7, o.closed_no_order], [8, o.Qualified], [9, o.hot], [10, o.qty]]
    .forEach(([cc, v]) => _put(ws, r, cc, v, cc === 10 ? PLAIN : INT, F_TOT, C, FL_CREAM));
  _put(ws, r, 3, 1, PCT, F_TOT, C, FL_CREAM);
  const end2 = r;

  const b3 = end2 + 2;
  _bar(ws, b3, "  STAGE WISE SNAPSHOT  (UNIFIED)", 3 + sources.length);
  _hdr(ws, b3 + 1, ["Stage (Unified)", "No. of Leads", "% Share"].concat(sources));
  r = b3 + 2;
  for (const row of m.stage_matrix) {
    const s = row.stage;
    const fl = s === "Won" ? FL_G : C_.closed_no_order.indexOf(s) >= 0 ? FL_R
             : s === C_.untouched_stage ? FL_A : null;
    _put(ws, r, 1, s, null, F_BB, L, fl);
    _put(ws, r, 2, row.total, INT, F_BB, C, fl);
    _put(ws, r, 3, row.share / 100, PCT, F_B, C, fl);
    sources.forEach((src, i) => _put(ws, r, 4 + i, row[src], INT, F_B, C, fl));
    r++;
  }
  _put(ws, r, 1, "TOTAL", null, F_TOT, L, FL_CREAM);
  _put(ws, r, 2, o.total, INT, F_TOT, C, FL_CREAM);
  _put(ws, r, 3, 1, PCT, F_TOT, C, FL_CREAM);
  sources.forEach((src, i) => _put(ws, r, 4 + i, m.by_source[src].total, INT, F_TOT, C, FL_CREAM));
  const end3 = r;

  const b4 = end3 + 2;
  _bar(ws, b4, "  LEAD QUALITY SNAPSHOT  (HOT / WARM / COLD)", 5 + sources.length);
  _hdr(ws, b4 + 1, ["Lead Quality", "Total", "% Share"].concat(sources, ["Qualified", "Not Qualified"]));
  r = b4 + 2;
  for (const row of m.quality_matrix) {
    const q = row.quality;
    const fl = q === "Hot Lead" ? FL_G
             : (q === "Not Interested" || q === "Don't Pick Call") ? FL_R : null;
    _put(ws, r, 1, q, null, F_BB, L, fl);
    _put(ws, r, 2, row.total, INT, F_BB, C, fl);
    _put(ws, r, 3, row.share / 100, PCT, F_B, C, fl);
    sources.forEach((src, i) => _put(ws, r, 4 + i, row[src], INT, F_B, C, fl));
    _put(ws, r, 4 + sources.length, row.qualified, INT, F_B, C, fl);
    _put(ws, r, 5 + sources.length, row.not_qualified, INT, F_B, C, fl);
    r++;
  }
  _put(ws, r, 1, "TOTAL", null, F_TOT, L, FL_CREAM);
  _put(ws, r, 2, o.total, INT, F_TOT, C, FL_CREAM);
  _put(ws, r, 3, 1, PCT, F_TOT, C, FL_CREAM);
  _widths(ws, [22, 13].concat(sources.map(() => 11), stages.map(() => 11), [12, 10]));

  /* ----------------------------------------------------- RECONCILIATION */
  ws = wb.addWorksheet("Reconciliation");
  ws.views = [{ showGridLines: false }];
  _title(ws, "Reconciliation & Audit Trail",
         "Every row that entered the engine is accounted for below. " +
         "If the control total balances, nothing was silently dropped.", 9);
  const t = rec.control_totals;
  _bar(ws, 4, "  CONTROL TOTALS", 9);
  _hdr(ws, 5, ["File", "Sheet(s) detected", "Format", "Rows read", "Blank rows",
               "Parsed", "No date", "Out of period", "Duplicates", "Accepted"]);
  r = 6;
  const KEYS = ["rows_read", "rows_blank", "parsed", "no_date", "out_of_period",
                "duplicates", "accepted"];
  for (const p of rec.per_file) {
    _put(ws, r, 1, p.file, null, F_BB, L);
    _put(ws, r, 2, p.sheets.map(s => s.sheet).join(", "), null, F_B, L);
    _put(ws, r, 3, p.sheets.map(s => s.label).join(", "), null, F_B, L);
    KEYS.forEach((k, i) => _put(ws, r, 4 + i, p[k], INT, k === "accepted" ? F_BB : F_B));
    r++;
  }
  _put(ws, r, 1, "TOTAL", null, F_TOT, L, FL_CREAM);
  _put(ws, r, 2, null, null, F_TOT, L, FL_CREAM);
  _put(ws, r, 3, null, null, F_TOT, L, FL_CREAM);
  KEYS.forEach((k, i) => _put(ws, r, 4 + i, t[k], INT, F_TOT, C, FL_CREAM));
  r += 2;
  const bal = ws.getCell(r, 1);
  bal.value = t.balanced
    ? "BALANCED — parsed − no-date − out-of-period − duplicates = accepted"
    : "OUT OF BALANCE — investigate before using this report";
  bal.font = font(11, { bold: true, color: col(t.balanced ? GREEN : RED) });
  ws.mergeCells(r, 1, r, 10);

  r += 2;
  _bar(ws, r, "  EXCEPTION REGISTER", 6);
  _hdr(ws, r + 1, ["Severity", "Type", "What was found", "Detail", "Source row", "Suggested fix"]);
  let rr = r + 2;
  for (const e of rec.exceptions) {
    const fl = e.severity === "critical" ? FL_PINK : e.severity === "high" ? FL_R
             : e.severity === "medium" ? FL_A : null;
    _put(ws, rr, 1, e.severity.toUpperCase(), null, F_BB, C, fl);
    _put(ws, rr, 2, e.kind, null, F_B, L, fl);
    _put(ws, rr, 3, e.message, null, F_B, L, fl);
    _put(ws, rr, 4, e.detail.slice(0, 180), null, F_N, L, fl);
    _put(ws, rr, 5, e.ref, null, F_N, L, fl);
    _put(ws, rr, 6, e.fix, null, F_N, L, fl);
    rr++;
  }
  if (rr - 1 >= r + 1) {
    ws.autoFilter = { from: { row: r + 1, column: 1 }, to: { row: rr - 1, column: 6 } };
  }
  _widths(ws, [13, 18, 46, 52, 34, 44, 13, 13, 13, 12]);

  /* ------------------------------------------------------------ INSIGHTS */
  ws = wb.addWorksheet("Findings & Actions");
  ws.views = [{ showGridLines: false }];
  _title(ws, "Findings & Recommended Actions",
         `Generated from the reconciled data for ${per}`, 8);
  _hdr(ws, 4, ["#", "Severity", "Finding", "Why it matters", "Action"]);
  r = 5;
  insights.forEach((f, i) => {
    const fl = f.sev === "critical" ? FL_PINK : f.sev === "high" ? FL_R : FL_A;
    _put(ws, r, 1, i + 1, INT, F_BB, C, fl);
    _put(ws, r, 2, f.sev.toUpperCase(), null, F_BB, C, fl);
    _put(ws, r, 3, f.title, null, F_BB, LW, fl);
    _put(ws, r, 4, f.why, null, F_N, LW, fl);
    _put(ws, r, 5, f.action, null, F_B, LW, fl);
    ws.getRow(r).height = 34;
    r++;
  });
  _widths(ws, [5, 12, 52, 62, 46]);

  /* ---------------------------------------------------------- UNASSIGNED */
  if (m.unassigned.length) {
    ws = wb.addWorksheet("Unassigned Leads");
    ws.views = [{ showGridLines: false }];
    _title(ws, "Unassigned Leads — no owner",
           `${m.unassigned.length} leads reached the system with a blank salesperson and ` +
           `were never picked up. Largest quantity first.`, 8);
    _hdr(ws, 4, ["Qty (pcs)", "Age (days)", "Date", "Source", "Customer",
                 "Phone", "City / Location", "Requirement"]);
    r = 5;
    for (const u of m.unassigned) {
      _put(ws, r, 1, u.qty, PLAIN, F_BB);
      _put(ws, r, 2, u.age, INT, u.age >= 4 ? font(10, { bold: true, color: col(RED) }) : F_B);
      _put(ws, r, 3, xlDate(u.date), DMY);
      _put(ws, r, 4, u.source);
      _put(ws, r, 5, u.customer, null, F_B, L);
      _put(ws, r, 6, u.phone);
      _put(ws, r, 7, u.city, null, F_B, L);
      _put(ws, r, 8, u.product, null, F_B, L);
      if (u.qty >= 50) for (let cc = 1; cc <= 8; cc++) ws.getCell(r, cc).fill = FL_PINK;
      r++;
    }
    _widths(ws, [11, 11, 12, 12, 26, 14, 20, 46]);
  }

  /* ------------------------------------------------------- DEMAND / LOSS */
  ws = wb.addWorksheet("Demand & Loss");
  ws.views = [{ showGridLines: false }];
  _title(ws, "Demand Lost & Disposition Analysis",
         "Pieces at stake behind every lead closed without an order.", 6);
  const d = m.demand_totals;
  _bar(ws, 4, "  DEMAND AT A GLANCE (pieces)", 6);
  _hdr(ws, 5, ["Total demand", "Closed without order", "Still open", "Won",
               "Rejected on product range", "% of demand closed"]);
  _put(ws, 6, 1, d.total, PLAIN, F_BB);
  _put(ws, 6, 2, d.closed, PLAIN, F_BB, C, FL_R);
  _put(ws, 6, 3, d.open, PLAIN, F_BB, C, FL_A);
  _put(ws, 6, 4, d.won, PLAIN, F_BB, C, FL_G);
  _put(ws, 6, 5, d.product_gap, PLAIN, F_BB, C, FL_A);
  _put(ws, 6, 6, d.total ? d.closed / d.total : 0, PCT, F_BB);
  _bar(ws, 8, "  WHY THAT DEMAND WAS LOST", 4);
  _hdr(ws, 9, ["Reason", "Leads", "Pieces at stake", "% of lost demand"]);
  r = 10;
  const tq = m.demand.reduce((s, x) => s + x.qty, 0) || 1;
  for (const x of m.demand) {
    const fl = (x.severity === "product" || x.severity === "recoverable") ? FL_A : null;
    _put(ws, r, 1, x.label, null, F_BB, L, fl);
    _put(ws, r, 2, x.leads, INT, F_B, C, fl);
    _put(ws, r, 3, x.qty, PLAIN, F_BB, C, fl);
    _put(ws, r, 4, x.qty / tq, PCT, F_B, C, fl);
    r++;
  }
  r += 1;
  _bar(ws, r, "  EVERY DISPOSITION (all leads closed without an order)", 4);
  _hdr(ws, r + 1, ["Reason", "Leads", "% of closed", "Pieces"]);
  rr = r + 2;
  for (const x of m.dispositions) {
    _put(ws, rr, 1, x.label, null, F_BB, L);
    _put(ws, rr, 2, x.leads, INT);
    _put(ws, rr, 3, x.share / 100, PCT);
    _put(ws, rr, 4, x.qty, PLAIN);
    rr++;
  }
  rr += 1;
  _bar(ws, rr, "  WORTH CHASING AGAIN", 9);
  _hdr(ws, rr + 1, ["Qty", "Date", "Source", "Salesperson", "Customer",
                    "Requirement", "Quality", "Stage", "Note logged"]);
  let r2 = rr + 2;
  for (const x of m.recoverable) {
    _put(ws, r2, 1, x.qty, PLAIN, F_BB);
    _put(ws, r2, 2, xlDate(x.date), DMY);
    _put(ws, r2, 3, x.source);
    _put(ws, r2, 4, x.rep, null, F_B, L);
    _put(ws, r2, 5, x.customer, null, F_B, L);
    _put(ws, r2, 6, x.product, null, F_B, L);
    _put(ws, r2, 7, x.quality, null, F_B, C, x.quality === "Hot Lead" ? FL_G : null);
    _put(ws, r2, 8, x.stage, null, F_BB, C, FL_R);
    _put(ws, r2, 9, x.note || "— nothing recorded —", null, F_N, L);
    r2++;
  }
  _widths(ws, [13, 12, 14, 16, 26, 40, 16, 15, 34]);

  /* --------------------------------------------------------- MASTER DATA */
  ws = wb.addWorksheet("Master Data");
  _title(ws, "Master Data — reconciled lead register",
         `${rec.leads.length} accepted leads for ${per}, normalised across every source.`, 17);
  const cols = ["Lead ref", "Date", "Source", "Salesperson", "Assigned?", "Customer", "Phone",
                "City", "Product / Requirement", "Qty", "Stage (Unified)", "Original status",
                "Lead Quality", "Note", "Disposition", "Source file", "Row"];
  _hdr(ws, 4, cols);
  r = 5;
  const order = {};
  stages.forEach((s, i) => { order[s] = i; });
  const UN = C_.unassigned_label;
  const sorted = rec.leads.slice().sort((a, b) =>
    ((a.rep === UN) - (b.rep === UN)) ||
    (a.rep < b.rep ? -1 : a.rep > b.rep ? 1 : 0) ||
    (order[a.stage] - order[b.stage]) ||
    (a.date < b.date ? -1 : a.date > b.date ? 1 : 0));

  const LEFT_COLS = new Set([5, 7, 8, 13, 14, 15]);          // 0-based, as in the Python
  for (const l of sorted) {
    const vals = [`${l.source.slice(0, 3).toUpperCase()}-${String(l.src_row).padStart(5, "0")}`,
                  xlDate(l.date), l.source, l.rep, l.assigned ? "Yes" : "No", l.customer,
                  l.phone, l.city, l.product, l.qty, l.stage, l.raw_stage,
                  l.quality, l.note, l.disposition.label, l.src_file, l.src_row];
    vals.forEach((v, i) => {
      const c = ws.getCell(r, 1 + i);
      c.value = (v === undefined || v === "" ? null : v);
      c.font = F_B; c.border = BORD;
      c.alignment = LEFT_COLS.has(i) ? L : C;
      if (i === 1) c.numFmt = DMY;
      if (i === 9) c.numFmt = PLAIN;
    });
    const st = l.stage;
    const cell = ws.getCell(r, 11);
    const stageFill = st === "Won" ? FL_G : C_.closed_no_order.indexOf(st) >= 0 ? FL_R
                    : st === C_.untouched_stage ? FL_A : null;
    if (stageFill) cell.fill = stageFill;
    cell.font = font(10, { bold: true, color: col(
      st === "Won" ? GREEN : C_.closed_no_order.indexOf(st) >= 0 ? RED
      : st === C_.untouched_stage ? AMBER : "000000") });
    if (!l.assigned) {
      for (const cc of [4, 5]) {
        ws.getCell(r, cc).fill = FL_PINK;
        ws.getCell(r, cc).font = F_PINK;
      }
    }
    r++;
  }
  ws.views = [{ state: "frozen", xSplit: 2, ySplit: 4, showGridLines: false }];
  if (r - 1 >= 4) ws.autoFilter = { from: { row: 4, column: 1 }, to: { row: r - 1, column: 17 } };
  _widths(ws, [13, 11, 11, 16, 10, 24, 13, 18, 40, 8, 17, 15, 15, 30, 26, 26, 7]);

  return wb.xlsx.writeBuffer();
}

const API = { build };
root.OakReport = API;
if (typeof module !== "undefined" && module.exports) module.exports = API;

})(typeof globalThis !== "undefined" ? globalThis : this);
