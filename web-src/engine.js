/* ===========================================================================
   OakCraft Lead Reconciler — reconciliation engine, browser edition.

   A faithful port of core/ingest.py, core/reconcile.py and core/metrics.py.
   No DOM, no network: give it workbook matrices and a config, get the same
   JSON structure the Flask API used to return. Runs in the browser and in
   Node (used by the test harness that diffs it against the Python original).
   =========================================================================== */
(function (root) {
"use strict";

const MAX_HEADER_SCAN = 8;   // how many top rows to try as the header row
const MAX_PROBE_COLS  = 80;  // ignore the long tail of empty columns Excel leaves

/* ------------------------------------------------------------------ utils */

const isStr = v => typeof v === "string";
const isDate = v => v instanceof Date && !isNaN(v);

function normKey(s) {
  return (s === null || s === undefined) ? "" : String(s).toLowerCase().replace(/[^a-z0-9]/g, "");
}

/** Python str() for the value kinds a spreadsheet cell can hold. */
function pyStr(v) {
  if (v === null || v === undefined) return "None";
  if (isDate(v)) {
    const p = n => String(n).padStart(2, "0");
    const t = `${p(v.getHours())}:${p(v.getMinutes())}:${p(v.getSeconds())}`;
    return `${v.getFullYear()}-${p(v.getMonth() + 1)}-${p(v.getDate())} ${t}`;
  }
  if (typeof v === "boolean") return v ? "True" : "False";
  return String(v);
}

function clean(v, nullTokens) {
  if (v === null || v === undefined) return "";
  const s = pyStr(v).replace(/\s+/g, " ").trim();
  return nullTokens.indexOf(s.toLowerCase()) >= 0 ? "" : s;
}

function toInt(v) {
  if (v === null || v === undefined) return 0;
  const raw = pyStr(v).trim();
  if (raw === "") return 0;
  const f = Number(raw);
  if (!isFinite(f)) return 0;
  const n = Math.trunc(f);
  return (n > 0 && n < 1000000) ? n : 0;
}

/** Python str.title() for ASCII: a letter following a non-letter is upper. */
function pyTitle(s) {
  let out = "", prevCased = false;
  for (const ch of s) {
    const cased = /[A-Za-z]/.test(ch);
    out += cased ? (prevCased ? ch.toLowerCase() : ch.toUpperCase()) : ch;
    prevCased = cased;
  }
  return out;
}

/* Date handling ------------------------------------------------------------
   Dates are carried as {y, m, d} triples so nothing depends on a timezone. */

const DEFAULT_FORMATS = ["%d-%m-%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M:%S",
                         "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"];

const FMT_PARTS = { "%d": "(\\d{1,2})", "%m": "(\\d{1,2})", "%Y": "(\\d{4})",
                    "%H": "(\\d{1,2})", "%M": "(\\d{1,2})", "%S": "(\\d{1,2})" };

function strptime(s, fmt) {
  const order = [];
  let pattern = "";
  for (let i = 0; i < fmt.length; i++) {
    if (fmt[i] === "%" && FMT_PARTS[fmt.slice(i, i + 2)]) {
      order.push(fmt.slice(i, i + 2));
      pattern += FMT_PARTS[fmt.slice(i, i + 2)];
      i++;
    } else {
      pattern += fmt[i].replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    }
  }
  const m = new RegExp("^" + pattern + "$").exec(s);
  if (!m) return null;
  const got = {};
  order.forEach((k, i) => { got[k] = parseInt(m[i + 1], 10); });
  const y = got["%Y"], mo = got["%m"], d = got["%d"];
  if (y === undefined || mo === undefined || d === undefined) return null;
  if (mo < 1 || mo > 12 || d < 1) return null;
  if (d > new Date(y, mo, 0).getDate()) return null;            // real month length
  if ((got["%H"] ?? 0) > 23 || (got["%M"] ?? 0) > 59 || (got["%S"] ?? 0) > 59) return null;
  return { y: y, m: mo, d: d };
}

function parseDate(v, formats) {
  if (v === null || v === undefined) return null;
  if (isStr(v) && !v.trim()) return null;
  if (isDate(v)) return { y: v.getFullYear(), m: v.getMonth() + 1, d: v.getDate() };
  const s = pyStr(v).trim();
  for (const f of (formats || []).concat(DEFAULT_FORMATS)) {
    const hit = strptime(s, f);
    if (hit) return hit;
  }
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s);
  if (m) {
    const y = +m[1], mo = +m[2], d = +m[3];
    if (mo >= 1 && mo <= 12 && d >= 1 && d <= new Date(y, mo, 0).getDate()) return { y, m: mo, d };
  }
  return null;
}

const iso = dd => `${dd.y}-${String(dd.m).padStart(2, "0")}-${String(dd.d).padStart(2, "0")}`;
const fromISO = s => ({ y: +s.slice(0, 4), m: +s.slice(5, 7), d: +s.slice(8, 10) });
const dayNum = dd => Math.floor(Date.UTC(dd.y, dd.m - 1, dd.d) / 86400000);
const fromDayNum = n => { const t = new Date(n * 86400000);
  return { y: t.getUTCFullYear(), m: t.getUTCMonth() + 1, d: t.getUTCDate() }; };
const addDays = (dd, k) => fromDayNum(dayNum(dd) + k);
const isoAddDays = (s, k) => iso(addDays(fromISO(s), k));
const diffDays = (a, b) => dayNum(fromISO(a)) - dayNum(fromISO(b));
const WEEKDAY = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];
const weekdayOf = s => WEEKDAY[((dayNum(fromISO(s)) % 7) + 7 + 3) % 7];   // 1970-01-01 = Thu

/* Percentages ---------------------------------------------------------------
   Python's round() breaks ties to even. Computing on the exact integer
   ratio rather than a float keeps the JS numbers identical to the Python. */
function pct(a, b) {
  if (!b) return 0.0;
  const neg = (a < 0) !== (b < 0);
  const N = 10000 * Math.abs(a), D = Math.abs(b);
  let q = Math.floor(N / D);
  const r = N - q * D;
  if (2 * r > D || (2 * r === D && q % 2 !== 0)) q += 1;
  return (neg ? -q : q) / 100;
}

/* Number grouping for the generated findings text. The reports are read in
   India, so the whole app groups the Indian way (12,34,567). */
const grp = x => Number(x || 0).toLocaleString("en-IN", { maximumFractionDigits: 0 });

/* Python's repr of a float always shows a decimal point: 30.0, not 30. */
const pyf = v => Number.isInteger(v) ? v.toFixed(1) : String(v);

/* ------------------------------------------------------------ detection */

function headerCandidates(rows) {
  const out = [];
  const last = Math.min(MAX_HEADER_SCAN, rows.length);
  for (let r = 1; r <= last; r++) {
    const row = rows[r - 1].slice(0, MAX_PROBE_COLS);
    let n = 0;
    for (const v of row) if (isStr(v) && v.trim()) n++;
    if (n >= 3) out.push([r, row]);
  }
  return out;
}

function scoreProfile(profile, headers) {
  const keys = new Set();
  for (const h of headers) if (h !== null && h !== undefined) keys.add(normKey(h));
  const req = profile.match.required.map(normKey);
  const anyc = (profile.match.any || []).map(normKey);
  if (!req.length || !req.every(k => keys.has(k))) return 0.0;
  if (anyc.length && !anyc.some(k => keys.has(k))) return 0.5;
  return 1.0;
}

/** sheets: [{name, rows}] — rows are raw cell matrices. */
function detectSheets(sheets, cfg) {
  const results = [];
  for (const sh of sheets) {
    let best = null;
    const cands = headerCandidates(sh.rows);
    for (const [headerRow, headers] of cands) {
      for (const prof of cfg.profiles) {
        const s = scoreProfile(prof, headers);
        if (s >= 1.0 && (best === null || s > best.score)) {
          best = { sheet: sh.name, profile: prof, header_row: headerRow,
                   headers: headers.filter(h => h !== null && h !== undefined), score: s };
        }
      }
      if (best) break;
    }
    if (best) results.push(best);
    else results.push({ sheet: sh.name, profile: null, header_row: null,
                        headers: (cands.length ? cands[0][1] : []).filter(h => h !== null && h !== undefined),
                        score: 0.0 });
  }
  return results;
}

/* ------------------------------------------------- field resolution */

function RowReader(headers) {
  const index = new Map();
  headers.forEach((h, i) => { const k = normKey(h); if (k && !index.has(k)) index.set(k, i); });
  return {
    headers,
    get(row, names) {
      for (const n of (names || [])) {
        const i = index.get(normKey(n));
        if (i !== undefined && i < row.length) {
          const v = row[i];
          if (v !== null && v !== undefined && pyStr(v).trim() !== "") return v;
        }
      }
      return null;
    },
    has(name) { return index.has(normKey(name)); },
  };
}

function mapQuality(raw, cfg, blank) {
  const s = (raw || "").toLowerCase();
  if (!s) return blank;
  for (const rule of cfg.quality_rules) if (s.indexOf(rule.contains) >= 0) return rule.value;
  return blank;
}

function mapStage(raw, stageCfg, canonical) {
  const s = (raw || "").trim();
  if (!s) return [stageCfg.blank || "New", true];
  const m = {};
  for (const [k, v] of Object.entries(stageCfg.map || {})) m[k.toUpperCase()] = v;
  if (s.toUpperCase() in m) return [m[s.toUpperCase()], true];
  for (const c of canonical) if (s.toLowerCase() === c.toLowerCase()) return [c, true];
  return [stageCfg.blank || "New", false];
}

function resolveRep(raw, cfg) {
  const s = String(raw === null || raw === undefined ? "" : pyStr(raw)).replace(/\s+/g, " ").trim();
  if (!s) return cfg.canonical.unassigned_label;
  const hit = (cfg.rep_aliases || {})[s.toLowerCase()];
  return hit ? hit : pyTitle(s);
}

const reCache = new Map();
const rx = p => { if (!reCache.has(p)) reCache.set(p, new RegExp(p)); return reCache.get(p); };

function classifyDisposition(note, quality, stage, cfg) {
  const d = (note || "").replace(/\s+/g, " ").trim().toLowerCase();
  const q = (quality || "").toLowerCase();
  for (const rule of cfg.disposition_rules) {
    if (rule.all_regex) {
      if (rule.all_regex.every(p => rx(p).test(d)))
        return { id: rule.id, label: rule.label, severity: rule.severity || "normal" };
    } else if (rx(rule.any_regex).test(d)) {
      return { id: rule.id, label: rule.label, severity: rule.severity || "normal" };
    }
  }
  if (!d) {
    if (q.indexOf("interest") >= 0) return { id: "no_req", label: "No requirement", severity: "normal" };
    if (q.indexOf("pick") >= 0) return { id: "no_resp", label: "No response / call not picked", severity: "normal" };
  }
  if (cfg.canonical.closed_no_order.indexOf(stage) >= 0)
    return { id: "not_logged", label: "Reason not logged", severity: "gap" };
  return { id: "none", label: "Not applicable", severity: "normal" };
}

/* ------------------------------------------------------------ read a file */

function readFile(sheets, filename, cfg) {
  const nulls = cfg.null_tokens.map(t => t.toLowerCase());
  const canonical = cfg.canonical.stages;
  const detections = detectSheets(sheets, cfg);
  const byName = new Map(sheets.map(s => [s.name, s]));

  const leads = [], sheetInfo = [], unmapped = [];
  let rowsRead = 0, rowsBlank = 0;

  for (const det of detections) {
    if (det.profile === null) {
      sheetInfo.push({ file: filename, sheet: det.sheet, profile: null, label: "not recognised",
                       rows: 0, headers: det.headers.slice(0, 25).map(String) });
      continue;
    }
    const prof = det.profile, src = byName.get(det.sheet);
    const rr = RowReader(det.headers);
    const f = prof.fields, sn = prof.source_name || {};
    const width = det.headers.length;
    let nSheet = 0;

    for (let i = det.header_row; i < src.rows.length; i++) {
      const excelRow = i + 1;
      const row = [];
      for (let c = 0; c < width; c++) row.push(c < src.rows[i].length ? src.rows[i][c] : null);
      rowsRead++;
      if (!row.some(c => c !== null && c !== undefined && pyStr(c).trim() !== "")) { rowsBlank++; continue; }

      // ---- source name
      let source, sourceDetail;
      if ("const" in sn) {
        source = sn.const;
        if (sn.sub_column && rr.has(sn.sub_column)) {
          const sub = clean(rr.get(row, [sn.sub_column]), nulls).toLowerCase();
          sourceDetail = (sn.sub_map || {})[sub] !== undefined ? sn.sub_map[sub] : (pyTitle(sub) || source);
        } else sourceDetail = source;
      } else {
        const rawSrc = clean(rr.get(row, [sn.column || ""]), nulls);
        source = sn.default || "Unknown";
        for (const [frag, name] of Object.entries(sn.normalize || {})) {
          if (rawSrc.toLowerCase().indexOf(frag) >= 0) { source = name; break; }
        }
        sourceDetail = source;
      }

      const rawStage = clean(rr.get(row, prof.stage.from), nulls);
      const [stage, ok] = mapStage(rawStage, prof.stage, canonical);
      if (!ok) unmapped.push({ file: filename, sheet: det.sheet, row: excelRow,
                               value: rawStage, field: prof.stage.from[0] });

      const rawQuality = clean(rr.get(row, prof.quality.from), nulls);
      const quality = mapQuality(rawQuality, cfg, prof.quality.blank || "Not Rated");
      const note = clean(rr.get(row, f.note.from), nulls);
      const qtyRaw = rr.get(row, f.qty.from);
      const qty = toInt(qtyRaw);
      const qtyDefect = !!(f.qty.from.length && qty === 0 && qtyRaw !== null && qtyRaw !== undefined &&
                           ["null", "none", "n/a", "-"].indexOf(pyStr(qtyRaw).trim().toLowerCase()) >= 0);

      const date = parseDate(rr.get(row, f.date.from), f.date.formats || []);
      const rep = resolveRep(rr.get(row, f.rep.from), cfg);
      const phoneRaw = rr.get(row, f.phone.from);

      leads.push({
        src_file: filename, src_sheet: det.sheet, src_row: excelRow,
        profile: prof.id, source, source_detail: sourceDetail,
        date: date ? iso(date) : null,
        rep, assigned: rep !== cfg.canonical.unassigned_label,
        customer: clean(rr.get(row, f.customer.from), nulls),
        phone: pyStr(phoneRaw === null || phoneRaw === undefined || phoneRaw === "" ? "" : phoneRaw)
                 .replace(/\D/g, "").slice(-10),
        city: clean(rr.get(row, f.city.from), nulls),
        product: clean(rr.get(row, f.product.from), nulls).slice(0, 80),
        qty, qty_defect: qtyDefect,
        stage, raw_stage: rawStage,
        quality, raw_quality: rawQuality,
        note,
        disposition: classifyDisposition(note, quality, stage, cfg),
      });
      nSheet++;
    }

    sheetInfo.push({ file: filename, sheet: det.sheet, profile: prof.id, label: prof.label,
                     rows: nSheet, headers: det.headers.slice(0, 25).map(String) });
  }

  return { leads, sheets: sheetInfo, rows_read: rowsRead, rows_blank: rowsBlank,
           unmapped_stage_values: unmapped };
}

/* ========================================================= reconciliation */

const normName = s => (s || "").toLowerCase().replace(/[^a-z0-9]/g, "");

function inferPeriod(leads) {
  const ds = leads.map(l => l.date).filter(Boolean).sort();
  if (!ds.length) {
    const t = new Date();
    const today = { y: t.getFullYear(), m: t.getMonth() + 1, d: t.getDate() };
    return [iso(addDays(today, -6)), iso(today)];
  }
  const end = ds[ds.length - 1];
  return [isoAddDays(end, -6), end];
}

function reconcile(files, cfg, period) {
  const unassignedLabel = cfg.canonical.unassigned_label;
  const allLeads = [];
  for (const f of files) allLeads.push(...f.leads);

  const [pStart, pEnd] = period || inferPeriod(allLeads);
  const exceptions = [];

  const ex = (kind, severity, message, detail, lead, fix) => exceptions.push({
    kind, severity, message, detail: detail || "", fix: fix || "",
    ref: lead ? `${lead.src_file} · ${lead.src_sheet} · row ${lead.src_row}` : "",
    customer: lead ? (lead.customer || "") : "",
  });

  // ------------------------------------------------------ 1. period filter
  const inPeriod = [], outPeriod = [], noDate = [];
  for (const l of allLeads) {
    if (!l.date) {
      noDate.push(l);
      ex("no_date", "high", "Row has no usable date and was excluded",
         `customer "${l.customer || "(blank)"}"`, l, "Check the date column format in the export");
    } else if (pStart <= l.date && l.date <= pEnd) {
      inPeriod.push(l);
    } else {
      outPeriod.push(l);
      ex("out_of_period", "low",
         `Row dated ${l.date} is outside the ${pStart} to ${pEnd} window`,
         `${l.source} · ${l.customer || "(blank)"}`, l, "Filter the export to the reporting week");
    }
  }

  // ------------------------------------------------------ 2. duplicates
  const keep = [], dropped = [];
  const seenPhone = new Map(), seenName = new Map();
  const placeholders = new Set((cfg.placeholder_names?.values || []).map(p => p.toLowerCase()));

  const ordered = inPeriod.slice().sort((a, b) =>
    (a.date < b.date ? -1 : a.date > b.date ? 1 :
     a.src_file < b.src_file ? -1 : a.src_file > b.src_file ? 1 :
     a.src_row - b.src_row));

  for (const l of ordered) {
    const ph = l.phone;
    const dupOf = (ph && ph.length === 10) ? seenPhone.get(ph) : undefined;

    if (dupOf !== undefined) {
      l._dup_of = `${dupOf.src_file} row ${dupOf.src_row}`;
      l._dup_kind = "phone";
      dropped.push(l);
      const cross = dupOf.source !== l.source;
      ex("duplicate", cross ? "high" : "medium",
         cross ? "Same phone number reached us through two different sources"
               : "Same phone number appears twice in one source",
         `"${l.customer}" · ${ph} · ${dupOf.source} row ${dupOf.src_row} vs ${l.source} row ${l.src_row}`,
         l, "Counted once only. De-duplicate at source to keep totals clean");
      continue;
    }

    if (ph && ph.length === 10) seenPhone.set(ph, l);

    const name = (l.customer || "").trim().toLowerCase();
    if (name && !placeholders.has(name) && !ph) {
      const nk = normName(l.customer) + "\u0000" + l.source;
      const prev = seenName.get(nk);
      if (prev && Math.abs(diffDays(l.date, prev.date)) <= 3) {
        ex("possible_duplicate", "low",
           "Two rows share a customer name within 3 days — both kept, please confirm",
           `"${l.customer}" · ${l.source} row ${l.src_row} and row ${prev.src_row}. ` +
           `No phone number on either row to confirm with.`, l,
           "Capture a phone number so duplicates can be detected reliably");
      } else {
        seenName.set(nk, l);
      }
    }
    keep.push(l);
  }

  const accepted = keep;

  // ------------------------------------------- 3. data-quality exceptions
  const closedNoOrder = new Set(cfg.canonical.closed_no_order);
  for (const l of accepted) {
    if (!l.assigned)
      ex("unassigned", "critical", "Lead has no salesperson assigned",
         `${l.source} · ${l.customer || "(blank)"} · ${l.qty || 0} pcs`, l,
         "Make Salesperson mandatory at lead capture");
    if (l.qty_defect)
      ex("qty_null", "medium", "Quantity field contains text instead of a number",
         `${l.source} · ${l.customer}`, l, "Stop writing 'null' into a numeric field");
    if (closedNoOrder.has(l.stage) && l.disposition.id === "not_logged")
      ex("no_reason", "medium", `Closed as ${l.stage} with no reason recorded`,
         `${l.source} · ${l.customer}`, l, "Make Loss Reason mandatory before closing a lead");
    if (l.disposition.severity === "recoverable" && closedNoOrder.has(l.stage))
      ex("recoverable", "high", "Closed without an order although the note records live interest",
         `${l.qty} pcs · ${l.customer} · note: "${l.note}"`, l, "Re-open and chase");
    if ((l.quality === "Hot Lead" || l.quality === "Warm Lead") && closedNoOrder.has(l.stage))
      ex("hot_lost", "high", `Lead rated ${l.quality} was closed as ${l.stage}`,
         `${l.qty} pcs · ${l.customer}`, l, "Review with the rep before writing off");
    if (l.quality === "Not Interested" && /requir/.test((l.note || "").toLowerCase()) &&
        !/not\s*requ|no\s*requ/.test((l.note || "").toLowerCase()))
      ex("contradiction", "medium", 'Rated "Not Interested" but the note records a requirement',
         `${l.customer} · note: "${l.note}"`, l, 'Block "Not Interested" when a requirement is logged');
  }

  for (const f of files) {
    for (const u of f.unmapped_stage_values) {
      exceptions.push({
        kind: "unmapped_stage", severity: "high",
        message: `Stage value "${u.value}" is not in the mapping and defaulted to New`,
        detail: `field ${u.field}`, fix: "Add this value to config/sources.json",
        ref: `${u.file} · ${u.sheet} · row ${u.row}`, customer: "",
      });
    }
  }

  // ------------------------------------------------------ 4. control totals
  const perFile = files.map(f => {
    const fname = f.sheets.length ? f.sheets[0].file : "?";
    return {
      file: fname, sheets: f.sheets, rows_read: f.rows_read, rows_blank: f.rows_blank,
      parsed: f.leads.length,
      no_date: noDate.filter(l => l.src_file === fname).length,
      out_of_period: outPeriod.filter(l => l.src_file === fname).length,
      duplicates: dropped.filter(l => l.src_file === fname).length,
      accepted: accepted.filter(l => l.src_file === fname).length,
    };
  });

  const sum = k => perFile.reduce((s, p) => s + p[k], 0);
  const totals = {
    rows_read: sum("rows_read"), rows_blank: sum("rows_blank"),
    parsed: allLeads.length, no_date: noDate.length, out_of_period: outPeriod.length,
    duplicates: dropped.length, accepted: accepted.length,
  };
  totals.balanced = (totals.parsed - totals.no_date - totals.out_of_period -
                     totals.duplicates === totals.accepted);

  const sevRank = { critical: 0, high: 1, medium: 2, low: 3 };
  exceptions.sort((a, b) => {
    const ra = sevRank[a.severity] ?? 9, rb = sevRank[b.severity] ?? 9;
    return ra !== rb ? ra - rb : (a.kind < b.kind ? -1 : a.kind > b.kind ? 1 : 0);
  });

  const counts = new Map();
  for (const e of exceptions) counts.set(e.kind, (counts.get(e.kind) || 0) + 1);
  const exceptionSummary = [...counts.entries()]
    .map(([kind, count], i) => ({ kind, count, i }))
    .sort((a, b) => (b.count - a.count) || (a.i - b.i))
    .map(({ kind, count }) => ({ kind, count,
      severity: exceptions.find(e => e.kind === kind).severity }));

  const repSet = new Set(accepted.map(l => l.rep));
  repSet.delete(unassignedLabel);
  const reps = [...repSet].sort();
  if (accepted.some(l => !l.assigned)) reps.push(unassignedLabel);

  return {
    period: { start: pStart, end: pEnd },
    leads: accepted,
    excluded: { no_date: noDate, out_of_period: outPeriod, duplicates: dropped },
    control_totals: totals,
    per_file: perFile,
    exceptions,
    exception_summary: exceptionSummary,
    reps,
    sources: [...new Set(accepted.map(l => l.source))].sort(),
  };
}

/* ================================================================ metrics */

function block(rows, cfg) {
  const C = cfg.canonical, stages = C.stages;
  const c = {};
  for (const s of stages) c[s] = 0;
  for (const r of rows) c[r.stage] = (c[r.stage] || 0) + 1;
  const total = rows.length;
  const nw = c[C.untouched_stage] || 0;
  const won = c[C.won_stage] || 0;
  const closedNo = C.closed_no_order.reduce((s, k) => s + (c[k] || 0), 0);
  const qplus = (c["Qualified"] || 0) + (c["Quoted / Proposal"] || 0) + won;
  const worked = total - nw;
  const out = {};
  for (const s of stages) out[s] = c[s] || 0;
  out.total = total;
  out.worked = worked;
  out.qplus = qplus;
  out.quoted_plus = (c["Quoted / Proposal"] || 0) + won;
  out.closed_no_order = closedNo;
  out.open = total - won - closedNo;
  out.qty = rows.reduce((s, r) => s + r.qty, 0);
  out.hot = rows.filter(r => r.quality === "Hot Lead").length;
  out.unassigned = rows.filter(r => !r.assigned).length;
  out.touch_rate = pct(worked, total);
  out.qual_rate = pct(qplus, total);
  out.quoted_rate = pct((c["Quoted / Proposal"] || 0) + won, total);
  out.conversion = pct(won, total);
  out.win_on_worked = pct(won, worked);
  out.loss_rate = pct(closedNo, total);
  return out;
}

function compute(rec, cfg) {
  const C = cfg.canonical;
  const leads = rec.leads;
  const stages = C.stages, quals = C.qualities;
  const reps = rec.reps, sources = rec.sources;
  const end = rec.period.end;
  const asof = isoAddDays(end, 1);

  const overall = block(leads, cfg);
  const byRep = {}; for (const r of reps) byRep[r] = block(leads.filter(l => l.rep === r), cfg);
  const bySource = {}; for (const s of sources) bySource[s] = block(leads.filter(l => l.source === s), cfg);
  const byDetail = {};
  for (const d of [...new Set(leads.map(l => l.source_detail))].sort())
    byDetail[d] = block(leads.filter(l => l.source_detail === d), cfg);

  const cube = {};
  for (const r of reps) {
    cube[r] = { ALL: byRep[r] };
    for (const s of sources) cube[r][s] = block(leads.filter(l => l.rep === r && l.source === s), cfg);
  }
  const totCube = { ALL: overall };
  for (const s of sources) totCube[s] = bySource[s];

  const stageMatrix = stages.map(s => {
    const row = { stage: s,
      total: leads.filter(l => l.stage === s).length,
      share: pct(leads.filter(l => l.stage === s).length, leads.length) };
    for (const src of sources) row[src] = leads.filter(l => l.stage === s && l.source === src).length;
    return row;
  });

  const qualityMatrix = quals.map(q => {
    const row = { quality: q,
      total: leads.filter(l => l.quality === q).length,
      share: pct(leads.filter(l => l.quality === q).length, leads.length),
      qualified: leads.filter(l => l.quality === q && l.stage === "Qualified").length,
      not_qualified: leads.filter(l => l.quality === q && l.stage === "Not Qualified").length };
    for (const src of sources) row[src] = leads.filter(l => l.quality === q && l.source === src).length;
    return row;
  });

  // ---- untouched ageing
  const newLeads = leads.filter(l => l.stage === C.untouched_stage);
  const ageOf = l => diffDays(asof, l.date);
  const tally = arr => {
    const m = new Map();
    for (const l of arr) m.set(ageOf(l), (m.get(ageOf(l)) || 0) + 1);
    const o = {};
    [...m.keys()].sort((a, b) => a - b).forEach(k => { o[String(k)] = m.get(k); });
    return o;
  };
  const ageing = tally(newLeads);
  const ageingByRep = {};
  for (const r of reps) ageingByRep[r] = tally(newLeads.filter(l => l.rep === r));
  const aged = newLeads.filter(l => ageOf(l) >= 2).length;

  // ---- unassigned register
  const unassigned = leads.filter(l => !l.assigned).map(l => ({
    qty: l.qty, age: ageOf(l), date: l.date, source: l.source, customer: l.customer,
    phone: l.phone, city: l.city, product: l.product,
  })).sort((a, b) => (b.qty - a.qty) || (a.date < b.date ? -1 : a.date > b.date ? 1 : 0));

  // ---- dispositions and demand lost
  const closed = leads.filter(l => C.closed_no_order.indexOf(l.stage) >= 0);
  const disp = new Map();
  for (const l of closed) {
    const k = l.disposition.label;
    if (!disp.has(k)) disp.set(k, { leads: 0, qty: 0, severity: "normal" });
    const d = disp.get(k);
    d.leads += 1; d.qty += l.qty; d.severity = l.disposition.severity;
  }
  const dispositions = [...disp.entries()]
    .map(([label, v], i) => ({ label, leads: v.leads, qty: v.qty, severity: v.severity,
                               share: pct(v.leads, closed.length), i }))
    .sort((a, b) => (b.leads - a.leads) || (a.i - b.i))
    .map(({ i, ...r }) => r);
  const demand = [...disp.entries()]
    .map(([label, v], i) => ({ label, leads: v.leads, qty: v.qty, severity: v.severity, i }))
    .filter(x => x.qty > 0)
    .sort((a, b) => (b.qty - a.qty) || (a.i - b.i))
    .map(({ i, ...r }) => r);

  const qsum = arr => arr.reduce((s, l) => s + l.qty, 0);
  const demandTotals = {
    total: qsum(leads),
    closed: qsum(closed),
    open: qsum(leads.filter(l => C.closed_no_order.concat([C.won_stage]).indexOf(l.stage) < 0)),
    won: qsum(leads.filter(l => l.stage === C.won_stage)),
    product_gap: qsum(closed.filter(l => l.disposition.severity === "product")),
  };

  // ---- leads worth chasing again
  const recoverable = closed
    .filter(l => l.qty >= 10 || l.quality === "Hot Lead" || l.quality === "Warm Lead" ||
                 l.disposition.severity === "recoverable" || l.disposition.severity === "product")
    .map((l, i) => ({ qty: l.qty, date: l.date, source: l.source, rep: l.rep,
                      customer: l.customer, product: l.product, quality: l.quality,
                      stage: l.stage, note: l.note, why: l.disposition.label, i }))
    .sort((a, b) => (b.qty - a.qty) || (a.i - b.i))
    .slice(0, 40)
    .map(({ i, ...r }) => r);

  // ---- daily
  const daily = [];
  for (let d = rec.period.start; d <= end; d = isoAddDays(d, 1)) {
    const rows = leads.filter(l => l.date === d);
    const row = { date: d, day: weekdayOf(d), total: rows.length };
    for (const s of stages) row[s] = rows.filter(l => l.stage === s).length;
    for (const s of sources) row["src_" + s] = rows.filter(l => l.source === s).length;
    daily.push(row);
  }

  return {
    period: rec.period, asof,
    stages, qualities: quals, reps, sources,
    overall, by_rep: byRep, by_source: bySource, by_detail: byDetail,
    cube, tot_cube: totCube,
    stage_matrix: stageMatrix, quality_matrix: qualityMatrix,
    ageing, ageing_by_rep: ageingByRep, aged_untouched: aged,
    unassigned, dispositions, demand, demand_totals: demandTotals,
    recoverable, daily,
  };
}

/* =============================================================== insights */

function buildInsights(m, rec, cfg) {
  const C = cfg.canonical;
  const out = [], o = m.overall, n = grp;

  if (m.unassigned.length) {
    const q = m.unassigned.reduce((s, u) => s + u.qty, 0);
    const top = m.unassigned[0];
    out.push({ sev: "critical",
      title: `${m.unassigned.length} leads have no salesperson assigned at all`,
      why: `They carry ${n(q)} pieces of demand` +
           (top.qty > 0 ? `, including a single ${n(top.qty)}-piece enquiry from ${top.customer}` : "") +
           `. Oldest is ${Math.max(...m.unassigned.map(u => u.age))} days old.`,
      action: "Assign every one of them today, largest quantity first." });
  }

  const d = m.demand_totals;
  if (d.total && d.closed) {
    out.push({ sev: "critical",
      title: `${n(d.closed)} of ${n(d.total)} pieces of demand were closed without an order`,
      why: `${pyf(pct(d.closed, d.total))}% of everything asked for this week. ` +
           `${n(d.open)} pieces are still open and winnable.`,
      action: "Work the open pieces before they age out." });
  }
  if (d.product_gap) {
    out.push({ sev: "high",
      title: `${n(d.product_gap)} pieces were rejected on product range, not on buyer intent`,
      why: "These buyers wanted something the range does not cover. " +
           "That is a sourcing decision sitting inside a sales report.",
      action: "Put the range gap in front of management with the piece count attached." });
  }

  if (o.New) {
    out.push({ sev: "high",
      title: `${o.New} of ${o.total} leads (${pyf(pct(o.New, o.total))}%) were never worked`,
      why: `${m.aged_untouched} of them are already 2+ days old. ` +
           `Online lead intent decays within hours.`,
      action: "Clear the aged backlog before taking on new spend." });
  }

  const worst = Object.entries(m.by_rep)
    .filter(([r, b]) => r !== C.unassigned_label && b.total >= 10)
    .sort((a, b) => a[1].touch_rate - b[1].touch_rate);
  if (worst.length && worst[0][1].touch_rate < 70) {
    const [r, b] = worst[0];
    out.push({ sev: "high",
      title: `${r} worked only ${pyf(b.touch_rate)}% of ${b.total} assigned leads`,
      why: `${b.New} leads never touched. ` +
           `Best in team is ${pyf(Math.max(...worst.map(x => x[1].touch_rate)))}%.`,
      action: "Redistribute the aged untouched leads and review capacity." });
    const split = m.sources.map(s => [s, m.cube[r][s]]).filter(([, b2]) => b2.total >= 8);
    const good = split.filter(([, b2]) => b2.qual_rate >= 20).map(([s]) => s);
    const bad = split.filter(([, b2]) => b2.qual_rate === 0).map(([s]) => s);
    if (good.length && bad.length) {
      out.push({ sev: "medium",
        title: `${r}’s gap is channel-specific, not a blanket skill problem`,
        why: `Qualifies well on ${good.join(", ")} but produces nothing from ${bad.join(", ")}.`,
        action: `Weight their allocation toward ${good[0]} and coach the rest.` });
    }
  }

  const ranked = Object.entries(m.by_source).filter(([, b]) => b.total >= 10)
    .sort((a, b) => b[1].qual_rate - a[1].qual_rate);
  if (ranked.length >= 2) {
    const best = ranked[0], low = ranked[ranked.length - 1];
    out.push({ sev: "high",
      title: `${best[0]} qualifies ${pyf(best[1].qual_rate)}% of its leads; ` +
             `${low[0]} manages ${pyf(low[1].qual_rate)}%`,
      why: `${best[0]}: ${best[1].total} leads, ${n(best[1].qty)} pcs of demand. ` +
           `${low[0]}: ${low[1].total} leads, ${n(low[1].qty)} pcs.`,
      action: `Shift budget and attention from ${low[0]} toward ${best[0]}.` });
  }

  const overload = Object.entries(m.by_rep).filter(([r]) => r !== C.unassigned_label)
    .sort((a, b) => b[1].total - a[1].total);
  if (overload.length >= 2 && overload[0][1].total > 1.4 * overload[overload.length - 1][1].total) {
    const [r, b] = overload[0];
    out.push({ sev: "medium",
      title: `${r} carries ${b.total} leads — ${pyf(pct(b.total, o.total))}% of everything`,
      why: `Loss rate ${pyf(b.loss_rate)}%. The lightest load in the team is ` +
           `${overload[overload.length - 1][1].total}.`,
      action: "Rebalance the allocation before quality of follow-up slips further." });
  }

  const gaps = {};
  for (const e of rec.exception_summary) gaps[e.kind] = e.count;
  if (gaps.no_reason) {
    out.push({ sev: "medium",
      title: `${gaps.no_reason} leads were closed with no reason recorded`,
      why: "Those losses cannot be diagnosed or coached against.",
      action: "Make Loss Reason mandatory before a lead can be closed." });
  }
  if (gaps.recoverable || gaps.hot_lost) {
    const k = (gaps.recoverable || 0) + (gaps.hot_lost || 0);
    out.push({ sev: "high",
      title: `${k} leads were written off despite recorded interest`,
      why: "Either rated Hot/Warm or carrying a note that records a live requirement.",
      action: "Re-open these first — cheapest pipeline available." });
  }
  if (!rec.leads.some(l => l.source === "Meta" && l.qty)) {
    out.push({ sev: "medium", title: "No deal value is captured in any feed",
      why: "Revenue conversion, average order value and ROI per source cannot be " +
           "calculated. Only lead counts and piece quantities are measurable.",
      action: "Add an Order Value field at Qualified stage." });
  }

  const rank = { critical: 0, high: 1, medium: 2 };
  return out.map((x, i) => ({ x, i }))
    .sort((a, b) => ((rank[a.x.sev] ?? 3) - (rank[b.x.sev] ?? 3)) || (a.i - b.i))
    .map(({ x }) => x);
}

/* =============================================================== exports */

const API = { readFile, reconcile, compute, buildInsights, pct, grp,
              iso, fromISO, isoAddDays, diffDays, parseDate, normKey, pyTitle, toInt, clean };
root.OakEngine = API;
if (typeof module !== "undefined" && module.exports) module.exports = API;

})(typeof globalThis !== "undefined" ? globalThis : this);
