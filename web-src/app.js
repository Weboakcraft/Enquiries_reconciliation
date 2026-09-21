/* OakCraft Lead Reconciler — front end. Everything runs in this tab: the
   spreadsheets are parsed here, reconciled here, and the Excel report is
   built here. Only the saved run summaries travel, to the team's store. */
(() => {
"use strict";

const E = window.OakEngine, RX = window.OakXlsx, REP = window.OakReport, CFG = window.OAK_CONFIG;

const $  = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];
const esc = s => String(s ?? "").replace(/[&<>"]/g, c =>
  ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const nf  = n => (Number(n) || 0).toLocaleString("en-IN");
const pcts = v => (v === 0 || v == null) ? "—" : `${Number(v).toFixed(1)}%`;
const pc2 = v => `${Number(v || 0).toFixed(2)}%`;
const dash = v => (v === 0 || v == null) ? "—" : nf(v);
const MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const fdate = s => {
  if (!s || s.length < 10) return s || "";
  return `${s.slice(8, 10)} ${MON[+s.slice(5, 7) - 1]}`;
};
const stamp = () => {
  const d = new Date(), p = n => String(n).padStart(2, "0");
  return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}-` +
         `${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}-` +
         Math.random().toString(36).slice(2, 6);
};

const STAGE_VAR = { "New": "--s-new", "Contacted": "--s-con", "Qualified": "--s-qual",
  "Quoted / Proposal": "--s-prop", "Won": "--s-won", "Lost": "--s-lost", "Not Qualified": "--s-nq" };
const STAGE_SHORT = { "New": "Never worked", "Quoted / Proposal": "Quoted" };
const short = s => STAGE_SHORT[s] || s;
const sv = s => `var(${STAGE_VAR[s] || "--muted"})`;
const QVAR = q => q === "Hot Lead" ? "--s-won" : q === "Warm Lead" ? "--s-new"
  : q === "Cold Lead" ? "--s-con" : q === "Not Rated" ? "--muted" : "--s-lost";

let DATA = null, FILES = [], SRC = "ALL", EXSEV = "ALL", IS_SAMPLE = false;

/* ═══════════════════════════════════════════════════ capabilities */

const useCap = name => {
  try {
    if (window.claude && typeof window.claude.use === "function")
      return window.claude.use(name).catch(() => null);
  } catch (_) { /* not in a capable host */ }
  return Promise.resolve(null);
};

/* ═══════════════════════════════════════════════════ run storage
   Shared team store when the page can reach one, this browser otherwise.
   A run's headline figures live in one small summary document; the full
   reconciled detail is gzipped and split across part documents beneath it. */

const Store = {
  mode: "local", db: null, user: null, myId: null, canWrite: true,
  runs: [], listeners: [],

  onChange(fn) { this.listeners.push(fn); },
  emit() { this.listeners.forEach(f => { try { f(); } catch (e) { console.error(e); } }); },

  async init() {
    this.readLocal();
    this.emit();
    const [db, user] = await Promise.all([useCap("db"), useCap("user")]);
    this.user = user;
    if (user) {
      try { this.myId = await user.id(); } catch (_) { this.myId = null; }
      try { const w = await user.can("data.write"); this.canWrite = (w !== false); } catch (_) {}
    }
    if (!db) { this.setPill(); return; }
    this.db = db;
    this.mode = "shared";
    this.setPill();
    try {
      db.collection("runs").limit(300).onSnapshot(snap => {
        this.runs = snap.docs.map(d => Object.assign({ run_id: d.id }, d.data()))
                             .filter(r => r && r.period_end)
                             .sort((a, b) => (a.run_id < b.run_id ? 1 : -1));
        this.emit();
      }, err => {
        console.warn("shared history unavailable:", err && err.code);
        this.mode = "local"; this.db = null; this.readLocal(); this.setPill(); this.emit();
      });
    } catch (err) {
      this.mode = "local"; this.db = null; this.readLocal(); this.setPill(); this.emit();
    }
  },

  setPill() {
    const el = $("#storePill");
    if (this.mode === "shared") {
      el.className = "pill live";
      el.textContent = this.canWrite ? "Shared with your team" : "Shared history · read only";
    } else {
      el.className = "pill local";
      el.textContent = "Saved in this browser";
    }
    const hint = $("#histHint");
    hint.textContent = this.mode === "shared"
      ? "Runs are saved to the team's shared history, so everyone here sees the same week-on-week movement. The spreadsheets themselves never leave the browser they were opened in."
      : "Runs are saved in this browser only. Open this page while signed in to your organisation to share the history with your team.";
  },

  /* ---- local fallback ---- */
  readLocal() {
    if (this.mode === "shared") return;
    try {
      this.runs = JSON.parse(localStorage.getItem("oak.runs") || "[]");
    } catch (_) { this.runs = []; }
  },
  writeLocal() {
    try { localStorage.setItem("oak.runs", JSON.stringify(this.runs)); } catch (_) {}
  },

  /* ---- payload packing ---- */
  async pack(obj) {
    const json = JSON.stringify(obj);
    if (typeof CompressionStream === "function") {
      try {
        const stream = new Blob([json]).stream().pipeThrough(new CompressionStream("gzip"));
        const buf = new Uint8Array(await new Response(stream).arrayBuffer());
        let bin = "";
        for (let i = 0; i < buf.length; i += 0x8000)
          bin += String.fromCharCode.apply(null, buf.subarray(i, i + 0x8000));
        return { enc: "gzip-b64", text: btoa(bin) };
      } catch (_) { /* fall through to plain JSON */ }
    }
    return { enc: "json", text: json };
  },
  async unpack(enc, text) {
    if (enc !== "gzip-b64") return JSON.parse(text);
    const bin = atob(text);
    const buf = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) buf[i] = bin.charCodeAt(i);
    const stream = new Blob([buf]).stream().pipeThrough(new DecompressionStream("gzip"));
    return JSON.parse(await new Response(stream).text());
  },

  summarise(payload) {
    const m = payload.metrics, o = m.overall;
    return {
      run_id: payload.run_id, created: payload.created, by: this.myId || null,
      period_start: m.period.start, period_end: m.period.end,
      leads: o.total, new: o.New, conversion: o.conversion, touch_rate: o.touch_rate,
      qual_rate: o.qual_rate, unassigned: o.unassigned, qty: o.qty,
      files: payload.reconciliation.per_file.map(f => f.file),
    };
  },

  async save(payload) {
    const summary = this.summarise(payload);
    const { enc, text } = await this.pack(payload);
    const SIZE = 180000;
    const parts = [];
    for (let i = 0; i < text.length; i += SIZE) parts.push(text.slice(i, i + SIZE));
    if (parts.length > 40) throw { code: "too_big", message: "This run is too large to save." };
    summary.chunks = parts.length;
    summary.enc = enc;

    if (this.mode === "shared") {
      const col = this.db.collection("runs");
      for (let i = 0; i < parts.length; i++) {
        await col.doc(summary.run_id).collection("parts").doc(String(i)).set({ d: parts[i] });
      }
      await col.doc(summary.run_id).set(summary);    // summary last: it is the commit
    } else {
      try { localStorage.setItem("oak.run." + summary.run_id, JSON.stringify({ enc, text })); }
      catch (_) { throw { code: "full", message: "This browser has no room left to save runs." }; }
      this.runs = [summary, ...this.runs.filter(r => r.run_id !== summary.run_id)];
      this.writeLocal();
      this.emit();
    }
    return summary;
  },

  async open(runId) {
    if (this.mode === "shared") {
      const doc = await this.db.collection("runs").doc(runId).get();
      if (!doc.exists) throw new Error("not found");
      const s = doc.data();
      let text = "";
      for (let i = 0; i < (s.chunks || 0); i++) {
        const p = await this.db.collection("runs").doc(runId).collection("parts").doc(String(i)).get();
        if (!p.exists) throw new Error("incomplete");
        text += p.data().d;
      }
      return this.unpack(s.enc, text);
    }
    const raw = localStorage.getItem("oak.run." + runId);
    if (!raw) throw new Error("not found");
    const { enc, text } = JSON.parse(raw);
    return this.unpack(enc, text);
  },

  async remove(runId) {
    if (this.mode === "shared") {
      const s = this.runs.find(r => r.run_id === runId);
      for (let i = 0; i < ((s && s.chunks) || 0); i++)
        await this.db.collection("runs").doc(runId).collection("parts").doc(String(i)).delete();
      await this.db.collection("runs").doc(runId).delete();
    } else {
      try { localStorage.removeItem("oak.run." + runId); } catch (_) {}
      this.runs = this.runs.filter(r => r.run_id !== runId);
      this.writeLocal();
      this.emit();
    }
  },

  /* Oldest first, one row per reporting week — the shape the trend table wants. */
  trend() {
    const rows = this.runs.slice().sort((a, b) =>
      a.period_end < b.period_end ? -1 : a.period_end > b.period_end ? 1 : 0);
    const out = [];
    for (const r of rows) {
      const k = r.period_start + "|" + r.period_end;
      const at = out.findIndex(x => x.period_start + "|" + x.period_end === k);
      if (at >= 0) out.splice(at, 1);
      out.push(r);
    }
    return out;
  },
};

/* ═══════════════════════════════════════════════════ navigation */

function show(view) {
  $$("main > section").forEach(s => { s.hidden = s.id !== "v-" + view; });
  $$("#nav button").forEach(b => b.setAttribute("aria-current", String(b.dataset.view === view)));
  window.scrollTo(0, 0);
}
$("#nav").addEventListener("click", e => {
  const b = e.target.closest("button[data-view]");
  if (b && !b.disabled) show(b.dataset.view);
});
function unlock(on) {
  $$("#nav button[data-view]").forEach(b => {
    if (!["upload", "history"].includes(b.dataset.view)) b.disabled = !on;
  });
}
let toastEl = null;
function toast(msg, ok) {
  if (toastEl) toastEl.remove();
  toastEl = document.createElement("div");
  toastEl.className = "toast" + (ok ? " ok" : "");
  toastEl.setAttribute("role", "status");
  toastEl.textContent = msg;
  document.body.appendChild(toastEl);
  const mine = toastEl;
  setTimeout(() => { if (mine === toastEl) { mine.remove(); toastEl = null; } }, 6500);
}

/* ═══════════════════════════════════════════════════ upload */

const drop = $("#drop"), fileInput = $("#file");
$("#browse").onclick = () => fileInput.click();
drop.onclick = e => { if (e.target === drop || e.target.closest(".icon,h3,p")) fileInput.click(); };
fileInput.onchange = () => { addFiles([...fileInput.files]); fileInput.value = ""; };
["dragenter", "dragover"].forEach(ev => drop.addEventListener(ev, e => {
  e.preventDefault(); drop.classList.add("hot");
}));
["dragleave", "drop"].forEach(ev => drop.addEventListener(ev, e => {
  e.preventDefault();
  if (ev === "dragleave" && drop.contains(e.relatedTarget)) return;
  drop.classList.remove("hot");
}));
drop.addEventListener("drop", e => addFiles([...(e.dataTransfer?.files || [])]));

function addFiles(list) {
  const ok = list.filter(f => /\.(xlsx|xlsm|xls)$/i.test(f.name));
  const bad = list.length - ok.length;
  ok.forEach(f => { if (!FILES.some(x => x.name === f.name && x.size === f.size)) FILES.push(f); });
  if (bad) toast(`${bad} file${bad > 1 ? "s" : ""} skipped — only Excel files are accepted.`);
  renderFiles();
}
function renderFiles() {
  $("#flist").innerHTML = FILES.map((f, i) => `
    <div class="frow">
      <span class="nm">${esc(f.name)}</span>
      <span class="sz">${(f.size / 1024).toFixed(0)} KB</span>
      <button class="x" data-i="${i}" title="Remove" aria-label="Remove ${esc(f.name)}">&times;</button>
    </div>`).join("");
  $("#run").disabled = !FILES.length;
  $("#clear").hidden = !FILES.length;
  $("#status").textContent = FILES.length ? `${FILES.length} file${FILES.length > 1 ? "s" : ""} ready` : "";
}
$("#flist").onclick = e => {
  const b = e.target.closest("button[data-i]");
  if (b) { FILES.splice(+b.dataset.i, 1); renderFiles(); }
};
$("#clear").onclick = () => { FILES = []; renderFiles(); };

const yieldToPaint = () => new Promise(r => setTimeout(r, 0));

$("#run").onclick = async () => {
  if (!FILES.length) return;
  $("#run").disabled = true; $("#demo").disabled = true; $("#prog").hidden = false;
  $("#status").textContent = "Reading files, detecting formats, reconciling…";
  await yieldToPaint();
  try {
    const parsed = [], skipped = [];
    for (const f of FILES) {
      $("#status").textContent = `Reading ${f.name}…`;
      await yieldToPaint();
      let res;
      try {
        const sheets = RX.readWorkbook(window.XLSX, new Uint8Array(await f.arrayBuffer()));
        res = E.readFile(sheets, f.name, CFG);
      } catch (err) {
        skipped.push({ file: f.name, reason: `could not be read: ${err.message || err}` });
        continue;
      }
      if (!res.leads.length) {
        const detected = res.sheets.map(s => s.label).join(", ") || "nothing";
        skipped.push({ file: f.name, reason: `no recognisable lead rows (detected: ${detected})` });
        continue;
      }
      parsed.push(res);
    }
    if (!parsed.length) {
      toast(skipped.length
        ? `${skipped[0].file} — ${skipped[0].reason}`
        : "None of the uploaded files matched a known format.");
      return;
    }
    $("#status").textContent = "Reconciling…";
    await yieldToPaint();
    const payload = buildRun(parsed, skipped);
    IS_SAMPLE = false;
    load(payload);
    show("overview");
    save(payload);
  } catch (err) {
    console.error(err);
    toast(`Analysis failed: ${err.message || err}`);
  } finally {
    $("#run").disabled = !FILES.length; $("#demo").disabled = false;
    $("#prog").hidden = true; $("#status").textContent = "";
  }
};

function buildRun(parsed, skipped) {
  const rec = E.reconcile(parsed, CFG);
  const m = E.compute(rec, CFG);
  const insights = E.buildInsights(m, rec, CFG);
  const light = Object.assign({}, rec);
  delete light.excluded;
  return {
    run_id: stamp(),
    created: new Date().toISOString().slice(0, 19),
    skipped: skipped || [],
    reconciliation: light, metrics: m, insights,
  };
}

async function save(payload) {
  if (IS_SAMPLE) return;
  if (Store.mode === "shared" && !Store.canWrite) {
    toast("Run complete. You have read-only access, so it was not added to the shared history.");
    return;
  }
  try {
    await Store.save(payload);
    toast(Store.mode === "shared"
      ? "Saved to the team's shared history."
      : "Saved in this browser.", true);
  } catch (err) {
    console.warn(err);
    toast(err && err.message ? `Not saved — ${err.message}`
                             : "The run could not be saved to history.");
  }
}

/* ═══════════════════════════════════════════════════ sample week */

$("#demo").onclick = async () => {
  $("#demo").disabled = true;
  await yieldToPaint();
  try {
    const sheets = sampleSheets();
    const parsed = sheets.map(s => E.readFile([s.sheet], s.file, CFG));
    IS_SAMPLE = true;
    load(buildRun(parsed, []));
    show("overview");
  } catch (err) {
    console.error(err); toast("The sample could not be built.");
  } finally { $("#demo").disabled = false; }
};
$("#dropSample").onclick = () => {
  IS_SAMPLE = false; DATA = null; unlock(false);
  $("#sampleBanner").hidden = true;
  $("#runLabel").textContent = "No run loaded";
  $("#navUn").textContent = ""; $("#navEx").textContent = "";
  show("upload");
};

/* A believable but entirely invented week, built from a fixed seed so the
   demo is the same for everyone, and run through the real pipeline. */
function sampleSheets() {
  let seed = 20260914;
  const rnd = () => (seed = (seed * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff;
  const pick = a => a[Math.floor(rnd() * a.length)];
  const end = new Date(); end.setHours(0, 0, 0, 0);
  const p = n => String(n).padStart(2, "0");
  const day = k => { const d = new Date(end.getTime() - k * 86400000);
    return { d, s: `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}` }; };

  const REPS = ["Ujala Rajput", "Amira Saikh", "Mayank Mittal", "Niti Kumari", ""];
  const CITY = ["Delhi", "Noida", "Gurugram", "Jaipur", "Mumbai", "Pune", "Lucknow"];
  const PROD = ["Office chair", "Workstation 6 seater", "Conference table", "Executive desk",
                "Visitor chair", "Storage unit", "Reception counter"];
  const NOTE = ["", "call back next week", "not picked, catalogue shared", "no requirement",
                "low price expectation", "requirement of 40 chairs", "wants plastic chairs",
                "delhi base dealer only", "call cut", "catalogue shared, awaiting reply"];
  const QUAL = ["Hot Lead", "Warm Lead", "Cold Lead", "Not Interested", "Don't Pick Call", ""];

  const phone = () => "9" + String(Math.floor(rnd() * 900000000) + 100000000);

  // 1. Meta export — text timestamps, no quality rating, no quantity
  const meta = [["created_time_ist", "platform", "full_name", "phone_number",
                 "city_state", "lead_status", "assigned_to"]];
  for (let i = 0; i < 74; i++) {
    const { d } = day(Math.floor(rnd() * 7));
    meta.push([`${p(d.getDate())}-${p(d.getMonth() + 1)}-${d.getFullYear()} ` +
               `${p(9 + Math.floor(rnd() * 9))}:${p(Math.floor(rnd() * 60))}:00`,
               pick(["ig", "fb"]),
               pick(["Rahul Sharma", "Priya Nair", "Imran Qureshi", "Sneha Rao", "Arjun Desai"]),
               phone(), pick(CITY),
               pick(["CREATED", "CREATED", "CREATED", "CONTACTED", "QUALIFIED", "LOST"]),
               pick(REPS)]);
  }

  // 2. IndiaMART workflow — real dates, quantities, quality ratings
  const im = [["Timestamp", "Buyer Name", "Mobile Number", "City/ Location", "Subject",
               "Quantity", "Assigned Salesperson", "Qualification Status",
               "Lead's Quality", "DISCUSSION"]];
  for (let i = 0; i < 96; i++) {
    const { d } = day(Math.floor(rnd() * 7));
    im.push([d, pick(["Sunrise Interiors", "Vertex Furnishings", "Nova Office", "Anand Traders",
                      "India Mart Buyer", "Crest Modular", "Lakeview Hotels"]),
             phone(), pick(CITY), pick(PROD),
             pick([0, 1, 4, 10, 24, 40, 75, 150, "null"]),
             pick(REPS),
             pick(["QUALIFIED", "NOT QUALIFIED", "LOST", "WON", ""]),
             pick(QUAL), pick(NOTE)]);
  }

  // 3. JustDial workflow — its own stage words, mixed sources
  const jd = [["Enquiry Id", "Date", "Customer", "Mobile", "City", "Product", "Qty",
               "Salesperson", "Stage", "Lead's Quality", "Final Remark's", "Source"]];
  for (let i = 0; i < 88; i++) {
    const { d } = day(Math.floor(rnd() * 7));
    jd.push([`E${String(1000 + i)}`, d,
             pick(["Jd Buyer", "Meridian Offices", "Shree Enterprises", "Kalpana Singh",
                   "Bluestone Realty", "Jd Buyer"]),
             phone(), pick(CITY), pick(PROD),
             pick([0, 2, 6, 12, 30, 60, 120]),
             pick(REPS),
             pick(["NEW", "NEW", "CONTACTED", "QUOTED", "WON", "LOST", "NOT QUALIFIED"]),
             pick(QUAL), pick(NOTE), pick(["JustDial", "JustDial", "India Mart"])]);
  }

  return [
    { file: "sample_meta_leads.xlsx",      sheet: { name: "Leads",     rows: meta } },
    { file: "sample_indiamart_leads.xlsx", sheet: { name: "Sheet1",    rows: im } },
    { file: "sample_justdial_leads.xlsx",  sheet: { name: "Enquiries", rows: jd } },
  ];
}

/* ═══════════════════════════════════════════════════ load a run */

function load(j) {
  DATA = j;
  unlock(true);
  const m = j.metrics;
  $("#sampleBanner").hidden = !IS_SAMPLE;
  $("#runLabel").textContent =
    `${m.period.start} → ${m.period.end} · ${m.overall.total} leads` + (IS_SAMPLE ? " · sample" : "");
  $("#navUn").textContent = m.overall.unassigned || "";
  $("#navEx").textContent = j.reconciliation.exceptions.length || "";
  if (j.skipped?.length) j.skipped.forEach(s => toast(`${s.file} skipped — ${s.reason}`));
  renderOverview(); renderMaster(); renderPeople(); renderSources();
  renderDemand(); renderUnassigned(); renderRecon(); renderExceptions();
}

/* ═══════════════════════════════════════════════════ overview */

function renderOverview() {
  const m = DATA.metrics, o = m.overall, rec = DATA.reconciliation;
  $("#ovPeriod").textContent =
    `${m.period.start} to ${m.period.end} · ${rec.per_file.length} file(s) · ` +
    `${m.sources.join(", ")} · reconciled ${new Date(DATA.created).toLocaleString("en-GB")}`;
  $("#asof").textContent = "as on " + fdate(m.asof);

  const crit = rec.exceptions.filter(e => e.severity === "critical").length;
  $("#ovKpis").innerHTML = [
    ["", "Leads reconciled", nf(o.total), `${rec.control_totals.rows_read} rows read in`],
    ["warn", "Never worked", nf(o.New), `${pcts(100 * o.New / (o.total || 1))} · ${m.aged_untouched} aged 2+ days`],
    [o.unassigned ? "crit" : "", "Unassigned", nf(o.unassigned),
      o.unassigned ? `${nf(m.unassigned.reduce((s, u) => s + u.qty, 0))} pcs with no owner` : "all leads owned"],
    ["bad", "Closed, no order", nf(o.closed_no_order), `${pcts(o.loss_rate)} of all leads`],
    ["good", "Conversion", pc2(o.conversion), `${o.Won} won out of ${o.total}`],
    [crit ? "crit" : "", "Exceptions", nf(rec.exceptions.length), `${crit} critical`],
  ].map(([cls, lab, num, foot]) => `
    <div class="kpi ${cls}"><div class="lab">${lab}</div>
      <div class="num">${num}</div><div class="foot">${esc(foot)}</div></div>`).join("");

  $("#srcFilter").innerHTML = $("#srcFilter2").innerHTML =
    ["ALL", ...m.sources].map(s => `<button data-src="${esc(s)}"
      aria-pressed="${s === SRC}">${s === "ALL" ? "All sources" : esc(s)}</button>`).join("");

  drawFunnel(); drawAgeing();

  $("#insights").innerHTML = DATA.insights.map((f, i) => `
    <div class="find"><div class="n">${String(i + 1).padStart(2, "0")}</div>
      <div><span class="sev ${f.sev}">${f.sev}</span>
        <div class="t" style="margin-top:6px">${esc(f.title)}</div>
        <div class="w">${esc(f.why)}</div>
        <div class="a">${esc(f.action)}</div></div></div>`).join("")
    || `<div class="empty">No findings — the data looks clean.</div>`;
}

function drawFunnel() {
  const m = DATA.metrics, t = m.tot_cube[SRC] || m.overall;
  const max = Math.max(...m.stages.map(s => t[s] || 0), 1);
  $("#funnel").innerHTML = m.stages.map(s => `
    <div class="fn-row">
      <div class="fn-name"><span class="dot" style="background:${sv(s)}"></span>${esc(short(s))}</div>
      <div class="trk"><i style="width:${100 * (t[s] || 0) / max}%;background:${sv(s)}"></i></div>
      <div class="fn-val"><b>${t[s] || 0}</b><span>${pcts(100 * (t[s] || 0) / (t.total || 1))}</span></div>
    </div>`).join("");
  $("#fnNote").innerHTML =
    `<b>${t.total}</b> leads · <b>${t.worked}</b> worked (${pcts(t.touch_rate)} touch rate) ·
     <b>${t.qplus}</b> reached Qualified or better (${pcts(t.qual_rate)}) ·
     <b>${nf(t.qty)}</b> pieces of demand · <b>${t.Won}</b> won (${pc2(t.conversion)} conversion)`;
  const lbl = SRC === "ALL" ? "All sources" : SRC;
  $("#fnScope").textContent = lbl; $("#peopleScope").textContent = lbl;
}

function drawAgeing() {
  const a = DATA.metrics.ageing, ks = Object.keys(a);
  if (!ks.length) { $("#ageing").innerHTML = `<div class="empty">Nothing untouched.</div>`;
                    $("#ageNote").textContent = ""; return; }
  const max = Math.max(...ks.map(k => a[k]));
  $("#ageing").innerHTML = ks.map(k => `
    <div class="col"><i class="${+k >= 2 ? "crit" : ""}" style="height:${20 + 100 * a[k] / max}px">${a[k]}</i>
      <span>${k}d</span></div>`).join("");
  const aged = DATA.metrics.aged_untouched;
  $("#ageNote").innerHTML = aged
    ? `<b style="color:var(--s-lost)">${aged} leads are 2 days or older and still untouched.</b>
       Online lead intent decays within hours — this is the cheapest recovery available.`
    : "Nothing older than a day is sitting untouched.";
}

document.addEventListener("click", e => {
  const b = e.target.closest("[data-src]");
  if (!b) return;
  SRC = b.dataset.src;
  $$("[data-src]").forEach(x => x.setAttribute("aria-pressed", String(x.dataset.src === SRC)));
  drawFunnel(); drawReps();
});

/* ═══════════════════════════════════════════════════ tables */

function table(el, head, rows) {
  $(el).innerHTML = `<thead><tr>${head.map(h => `<th>${h}</th>`).join("")}</tr></thead>
    <tbody>${rows.join("")}</tbody>`;
}
const td = (v, cls = "") => `<td class="${cls}">${v}</td>`;

function renderMaster() {
  const m = DATA.metrics, o = m.overall;
  $("#mrPeriod").textContent = `${m.period.start} to ${m.period.end}`;
  table("#tEmp",
    ["Salesperson", "Total", ...m.sources, ...m.stages.map(short), "Conversion %", "Loss %"],
    m.reps.map(r => {
      const b = m.by_rep[r], un = b.unassigned === b.total && b.total > 0;
      return `<tr class="${un ? "unass" : ""}">${td(esc(r), "nm")}${td(`<b>${b.total}</b>`)}
        ${m.sources.map(s => td(dash(m.cube[r][s].total))).join("")}
        ${m.stages.map(s => td(dash(b[s]), s === "Won" && b[s] ? "hi-g" : s === "New" && b[s] > 20 ? "hi-a" : "")).join("")}
        ${td(b.conversion ? pc2(b.conversion) : "—", b.conversion ? "hi-g" : "")}
        ${td(pcts(b.loss_rate), b.loss_rate > 50 ? "hi" : "")}</tr>`;
    }).concat(`<tr class="tot">${td("TOTAL")}${td(o.total)}
      ${m.sources.map(s => td(m.by_source[s].total)).join("")}
      ${m.stages.map(s => td(dash(o[s]))).join("")}
      ${td(pc2(o.conversion))}${td(pcts(o.loss_rate))}</tr>`));

  table("#tSrc", ["Source", "Leads", "Share", "Assigned", "Unassigned", "Won", "Closed", "Qualified", "Hot", "Qty"],
    m.sources.map(s => { const b = m.by_source[s];
      return `<tr>${td(esc(s), "nm")}${td(b.total)}${td(pcts(100 * b.total / (o.total || 1)))}
        ${td(b.total - b.unassigned)}${td(dash(b.unassigned), b.unassigned ? "hi" : "")}
        ${td(dash(b.Won), b.Won ? "hi-g" : "")}${td(b.closed_no_order)}${td(b.Qualified)}
        ${td(dash(b.hot))}${td(`<b>${nf(b.qty)}</b>`)}</tr>`; })
      .concat(`<tr class="tot">${td("TOTAL")}${td(o.total)}${td("100.0%")}
        ${td(o.total - o.unassigned)}${td(dash(o.unassigned))}${td(dash(o.Won))}
        ${td(o.closed_no_order)}${td(o.Qualified)}${td(dash(o.hot))}${td(nf(o.qty))}</tr>`));

  table("#tStage", ["Stage (unified)", "Leads", "Share", ...m.sources],
    m.stage_matrix.map(r => `<tr>${td(
      `<span class="dot" style="display:inline-block;background:${sv(r.stage)};margin-right:7px"></span>${esc(r.stage)}`, "nm")}
      ${td(`<b>${r.total}</b>`)}${td(pcts(r.share))}
      ${m.sources.map(s => td(dash(r[s]))).join("")}</tr>`)
      .concat(`<tr class="tot">${td("TOTAL")}${td(o.total)}${td("100.0%")}
        ${m.sources.map(s => td(m.by_source[s].total)).join("")}</tr>`));

  const qrows = m.quality_matrix.map(r => `<tr>${td(
    `<span class="dot" style="display:inline-block;background:var(${QVAR(r.quality)});margin-right:7px"></span>${esc(r.quality)}`, "nm")}
    ${td(`<b>${r.total}</b>`)}${td(pcts(r.share))}
    ${m.sources.map(s => td(dash(r[s]))).join("")}${td(dash(r.qualified))}${td(dash(r.not_qualified))}</tr>`);
  table("#tQual", ["Lead quality", "Leads", "Share", ...m.sources, "Qualified", "Not qualified"], qrows);
  table("#tQual2", ["Lead quality", "Leads", "Share", ...m.sources, "Qualified", "Not qualified"], qrows);
}

function renderPeople() {
  drawReps();
  const m = DATA.metrics, o = m.overall;
  table("#tScore",
    ["Salesperson", "Leads", "Never worked", "Touch rate", "Qualified+", "Qualification rate",
     "Quoted+", "Won", "Conversion %", "Closed no order", "Loss rate", "Open", "Hot", "Qty"],
    m.reps.map(r => { const b = m.by_rep[r], un = b.unassigned === b.total && b.total > 0;
      return `<tr class="${un ? "unass" : ""}">${td(esc(r), "nm")}${td(b.total)}
        ${td(dash(b.New), b.New > 20 ? "hi-a" : "")}${td(pcts(b.touch_rate), b.touch_rate < 60 ? "hi" : "")}
        ${td(dash(b.qplus))}${td(pcts(b.qual_rate), b.qual_rate === 0 ? "hi" : "")}
        ${td(dash(b.quoted_plus))}${td(dash(b.Won), b.Won ? "hi-g" : "")}
        ${td(b.conversion ? pc2(b.conversion) : "—", b.conversion ? "hi-g" : "")}
        ${td(b.closed_no_order)}${td(pcts(b.loss_rate), b.loss_rate > 50 ? "hi" : "")}
        ${td(b.open)}${td(dash(b.hot))}${td(nf(b.qty))}</tr>`; })
      .concat(`<tr class="tot">${td("TEAM TOTAL")}${td(o.total)}${td(o.New)}${td(pcts(o.touch_rate))}
        ${td(o.qplus)}${td(pcts(o.qual_rate))}${td(o.quoted_plus)}${td(o.Won)}${td(pc2(o.conversion))}
        ${td(o.closed_no_order)}${td(pcts(o.loss_rate))}${td(o.open)}${td(o.hot)}${td(nf(o.qty))}</tr>`));
}

function drawReps() {
  const m = DATA.metrics;
  $("#reps").innerHTML = m.reps.map(r => {
    const c = SRC === "ALL" ? m.by_rep[r] : m.cube[r][SRC], T = c.total;
    const un = m.by_rep[r].unassigned === m.by_rep[r].total && m.by_rep[r].total > 0;
    if (!T) return `<div class="rep"><div class="rhead"><span class="rname">${esc(r)}</span>
      <span class="rtot">no leads from this source</span></div></div>`;
    const bars = m.stages.filter(s => c[s] > 0).map(s => { const w = 100 * c[s] / T;
      return `<i style="width:${w}%;background:${sv(s)}" title="${esc(short(s))}: ${c[s]}">${w > 7 ? c[s] : ""}</i>`;
    }).join("");
    const chips = un
      ? `<span class="chip bad">No owner — <b>${T}</b> leads never opened</span>
         <span class="chip bad">Demand at risk <b>${nf(c.qty)}</b> pcs</span>`
      : [`<span class="chip ${c.touch_rate < 60 ? "bad" : ""}">Touch <b>${pcts(c.touch_rate)}</b></span>`,
         `<span class="chip ${c.New > 20 ? "warn" : ""}">Never worked <b>${c.New}</b></span>`,
         `<span class="chip ${c.qplus === 0 ? "bad" : ""}">Qualified+ <b>${c.qplus}</b> · ${pcts(c.qual_rate)}</span>`,
         `<span class="chip ${c.Won ? "good" : ""}">Conversion <b>${pc2(c.conversion)}</b></span>`,
         `<span class="chip">Hot <b>${c.hot}</b></span>`,
         `<span class="chip">Qty <b>${nf(c.qty)}</b> pcs</span>`].join("");
    return `<div class="rep"><div class="rhead">
        <span class="rname" ${un ? 'style="color:var(--pink)"' : ""}>${esc(r)}</span>
        <span class="rtot"><b>${T}</b> leads assigned</span></div>
      <div class="stack">${bars}</div><div class="chips">${chips}</div></div>`;
  }).join("");
  $("#legend").innerHTML = m.stages.map(s =>
    `<span class="lg"><span class="dot" style="background:${sv(s)}"></span>${esc(short(s))}</span>`).join("");
}

function renderSources() {
  const m = DATA.metrics, o = m.overall;
  table("#tSrcPerf",
    ["Source", "Leads", "Share", ...m.stages.map(short), "Qualification rate", "Conversion %", "Loss rate", "Hot", "Qty (pcs)"],
    m.sources.map(s => { const b = m.by_source[s];
      return `<tr>${td(esc(s), "nm")}${td(`<b>${b.total}</b>`)}${td(pcts(100 * b.total / (o.total || 1)))}
        ${m.stages.map(x => td(dash(b[x]))).join("")}
        ${td(pcts(b.qual_rate), b.qual_rate > 20 ? "hi-g" : "")}
        ${td(b.conversion ? pc2(b.conversion) : "—", b.conversion ? "hi-g" : "")}
        ${td(pcts(b.loss_rate), b.loss_rate > 80 ? "hi" : "")}${td(dash(b.hot))}${td(`<b>${nf(b.qty)}</b>`)}</tr>`; })
      .concat(`<tr class="tot">${td("ALL SOURCES")}${td(o.total)}${td("100.0%")}
        ${m.stages.map(x => td(dash(o[x]))).join("")}${td(pcts(o.qual_rate))}${td(pc2(o.conversion))}
        ${td(pcts(o.loss_rate))}${td(o.hot)}${td(nf(o.qty))}</tr>`));

  table("#tDetail", ["Platform", "Leads", "Never worked", "Qualified+", "Won", "Closed"],
    Object.entries(m.by_detail).map(([k, b]) =>
      `<tr>${td(esc(k), "nm")}${td(b.total)}${td(dash(b.New))}${td(dash(b.qplus))}
        ${td(dash(b.Won), b.Won ? "hi-g" : "")}${td(b.closed_no_order)}</tr>`));
}

function renderDemand() {
  const m = DATA.metrics, d = m.demand_totals;
  $("#demKpis").innerHTML = [
    ["", "Total demand", nf(d.total) + " pcs", "everything asked for this week"],
    ["bad", "Closed without order", nf(d.closed) + " pcs", pcts(100 * d.closed / (d.total || 1)) + " of demand"],
    ["warn", "Still open", nf(d.open) + " pcs", "winnable if worked now"],
    ["warn", "Rejected on range", nf(d.product_gap) + " pcs", "a sourcing decision, not a sales one"],
  ].map(([c, l, n, f]) => `<div class="kpi ${c}"><div class="lab">${l}</div>
      <div class="num">${n}</div><div class="foot">${f}</div></div>`).join("");

  const tq = m.demand.reduce((s, x) => s + x.qty, 0) || 1;
  table("#tDemand", ["Reason", "Leads", "Pieces at stake", "% of lost demand"],
    m.demand.map(x => `<tr class="${["product", "recoverable"].includes(x.severity) ? "flag" : ""}">
      ${td(esc(x.label), "nm")}${td(x.leads)}${td(`<b>${nf(x.qty)}</b>`)}${td(pcts(100 * x.qty / tq))}</tr>`)
      .concat(`<tr class="tot">${td("TOTAL")}${td(m.demand.reduce((s, x) => s + x.leads, 0))}
        ${td(nf(tq))}${td("100.0%")}</tr>`));

  table("#tDisp", ["Reason", "Leads", "% of closed", "Pieces"],
    m.dispositions.map(x => `<tr>${td(esc(x.label), "nm")}${td(x.leads)}
      ${td(pcts(x.share))}${td(dash(x.qty))}</tr>`));

  table("#tRecover", ["Qty", "Date", "Source", "Salesperson", "Customer", "Requirement", "Quality", "Stage", "Note"],
    m.recoverable.map(x => `<tr class="${x.quality === "Hot Lead" ? "flag" : ""}">
      ${td(`<b>${nf(x.qty)}</b>`)}${td(fdate(x.date))}${td(esc(x.source))}${td(esc(x.rep), "nm")}
      ${td(esc(x.customer), "nm")}${td(esc(x.product), "l")}
      ${td(esc(x.quality), x.quality === "Hot Lead" ? "hi-g" : "")}${td(esc(x.stage), "hi")}
      ${td(x.note ? esc(x.note) : "<i>nothing recorded</i>", "l")}</tr>`)
      .concat(m.recoverable.length ? [] : [`<tr><td class="l">Nothing to recover.</td></tr>`]));
}

function renderUnassigned() {
  const u = DATA.metrics.unassigned;
  if (!u.length) {
    $("#unKpis").innerHTML = "";
    $("#tUnassigned").innerHTML = `<tbody><tr><td class="l">
      <div class="empty"><h3>Every lead has an owner</h3>
      Nothing arrived with a blank salesperson this week.</div></td></tr></tbody>`;
    return;
  }
  const qty = u.reduce((s, x) => s + x.qty, 0);
  $("#unKpis").innerHTML = [
    ["crit", "Leads with no owner", nf(u.length), "nobody has opened these"],
    ["crit", "Demand at risk", nf(qty) + " pcs", "sitting with no salesperson"],
    ["bad", "Oldest", Math.max(...u.map(x => x.age)) + " days", "since it arrived"],
    ["warn", "Biggest single enquiry", nf(u[0].qty) + " pcs", esc(u[0].customer || "—")],
  ].map(([c, l, n, f]) => `<div class="kpi ${c}"><div class="lab">${l}</div>
      <div class="num">${n}</div><div class="foot">${f}</div></div>`).join("");
  table("#tUnassigned", ["Qty (pcs)", "Age", "Received", "Source", "Customer", "Phone", "City", "Requirement"],
    u.map(x => `<tr class="${x.qty >= 50 ? "unass" : ""}">
      ${td(`<b>${nf(x.qty)}</b>`)}${td(x.age + "d", x.age >= 4 ? "hi" : "")}${td(fdate(x.date))}
      ${td(esc(x.source))}${td(esc(x.customer), "nm")}${td(esc(x.phone))}
      ${td(esc(x.city), "l")}${td(esc(x.product), "l")}</tr>`));
}

/* ═══════════════════════════════════════════════════ reconciliation */

function renderRecon() {
  const rec = DATA.reconciliation, t = rec.control_totals;
  $("#balance").innerHTML = `<div class="balance ${t.balanced ? "" : "off"}">
    <span class="tick">${t.balanced ? "&#10003;" : "!"}</span>
    <span>${t.balanced
      ? "Balanced — every row is accounted for. Parsed − no-date − out-of-period − duplicates = accepted."
      : "Out of balance — investigate before relying on these numbers."}</span></div>`;
  $("#flow").innerHTML = [
    ["parsed", "Rows parsed", ""], ["no_date", "No date", "minus"],
    ["out_of_period", "Out of period", "minus"], ["duplicates", "Duplicates", "minus"],
    ["accepted", "Accepted", "end"],
  ].map(([k, lab, cls], i) => `${i ? `<span class="op">${cls === "end" ? "=" : "−"}</span>` : ""}
      <div class="step ${cls}"><div class="v">${nf(t[k])}</div><div class="k">${lab}</div></div>`).join("");

  table("#tFiles", ["File", "Rows read", "Blank", "Parsed", "No date", "Out of period", "Duplicates", "Accepted"],
    rec.per_file.map(p => `<tr>${td(esc(p.file), "nm")}${td(nf(p.rows_read))}${td(dash(p.rows_blank))}
      ${td(nf(p.parsed))}${td(dash(p.no_date))}${td(dash(p.out_of_period))}
      ${td(dash(p.duplicates), p.duplicates ? "hi" : "")}${td(`<b>${nf(p.accepted)}</b>`)}</tr>`)
      .concat(`<tr class="tot">${td("TOTAL")}${td(nf(t.rows_read))}${td(nf(t.rows_blank))}
        ${td(nf(t.parsed))}${td(dash(t.no_date))}${td(dash(t.out_of_period))}
        ${td(dash(t.duplicates))}${td(nf(t.accepted))}</tr>`));

  const sheets = rec.per_file.flatMap(p => p.sheets);
  table("#tSheets", ["File", "Sheet", "Format detected", "Rows", "Columns seen"],
    sheets.map(s => `<tr>${td(esc(s.file), "nm")}${td(esc(s.sheet))}
      ${td(s.profile ? esc(s.label) : "<i>not recognised — skipped</i>", s.profile ? "hi-g" : "")}
      ${td(dash(s.rows))}${td(`<span class="mono" style="font-size:10.5px">${esc((s.headers || []).slice(0, 10).join(" · "))}</span>`, "l")}</tr>`));
}

function renderExceptions() {
  const rec = DATA.reconciliation;
  const sevs = ["ALL", "critical", "high", "medium", "low"].filter(s =>
    s === "ALL" || rec.exceptions.some(e => e.severity === s));
  if (!sevs.includes(EXSEV)) EXSEV = "ALL";
  $("#exFilter").innerHTML = sevs.map(s => {
    const n = s === "ALL" ? rec.exceptions.length : rec.exceptions.filter(e => e.severity === s).length;
    return `<button data-sev="${s}" aria-pressed="${s === EXSEV}">${s === "ALL" ? "All" : s} (${n})</button>`;
  }).join("");
  drawExceptions();
}
function drawExceptions() {
  const rows = DATA.reconciliation.exceptions.filter(e => EXSEV === "ALL" || e.severity === EXSEV);
  table("#tExceptions", ["Severity", "Type", "What was found", "Detail", "Source row", "Suggested fix"],
    rows.map(e => `<tr>${td(`<span class="sev ${e.severity}">${e.severity}</span>`)}
      ${td(`<span class="mono" style="font-size:11px">${esc(e.kind)}</span>`, "l")}
      ${td(esc(e.message), "l")}${td(`<span style="color:var(--muted)">${esc(e.detail)}</span>`, "l")}
      ${td(`<span class="mono" style="font-size:10.5px;color:var(--muted)">${esc(e.ref)}</span>`, "l")}
      ${td(`<span style="color:var(--muted)">${esc(e.fix)}</span>`, "l")}</tr>`)
      .concat(rows.length ? [] : [`<tr><td class="l">Nothing at this severity.</td></tr>`]));
}
$("#exFilter").addEventListener("click", e => {
  const b = e.target.closest("[data-sev]"); if (!b) return;
  EXSEV = b.dataset.sev;
  $$("#exFilter button").forEach(x => x.setAttribute("aria-pressed", String(x.dataset.sev === EXSEV)));
  drawExceptions();
});

/* ═══════════════════════════════════════════════════ history */

const initials = n => (n || "?").trim().split(/\s+/).slice(0, 2).map(w => w[0] || "").join("").toUpperCase();

async function paintNames() {
  const user = Store.user;
  const nodes = $$("[data-who]");
  if (!nodes.length) return;
  if (!user || typeof user.profiles !== "function") {
    nodes.forEach(n => n.closest(".who")?.remove());
    return;
  }
  const ids = [...new Set(nodes.map(n => n.dataset.who).filter(Boolean))];
  let ps = {};
  try { ps = (await user.profiles(ids)) || {}; } catch (_) {}
  nodes.forEach(n => {
    const name = (ps[n.dataset.who] && ps[n.dataset.who].name) || "";
    const who = n.closest(".who");
    if (!name) { who?.remove(); return; }
    n.textContent = name;
    const av = who.querySelector(".av");
    if (av) av.textContent = initials(name);
  });
}

function renderHistory() {
  const trend = Store.trend();
  const canWrite = Store.mode !== "shared" || Store.canWrite;

  if (!trend.length) {
    $("#tTrend").innerHTML = `<tbody><tr><td class="l"><div class="empty">
      <h3>No history yet</h3>Run a week, then run the next one, and the comparison appears here.
      </div></td></tr></tbody>`;
    $("#runs").innerHTML = `<div class="empty">Nothing saved yet.</div>`;
    $("#recentWrap").hidden = true;
    return;
  }
  const delta = (cur, prev, better = "up") => {
    if (prev == null) return "";
    const d = +(cur - prev).toFixed(2);
    if (!d) return `<span class="delta">·</span>`;
    const good = better === "up" ? d > 0 : d < 0;
    return `<span class="delta ${good ? "up" : "down"}">${d > 0 ? "▲" : "▼"} ${Math.abs(d).toFixed(2)}</span>`;
  };
  table("#tTrend", ["Week ending", "Leads", "Never worked", "Touch rate", "Qualification rate",
                    "Conversion %", "Unassigned", "Demand (pcs)"],
    trend.map((r, i) => { const p = i ? trend[i - 1] : null;
      return `<tr>${td(`${fdate(r.period_start)} – ${fdate(r.period_end)}`, "nm")}
        ${td(nf(r.leads))}${td(dash(r.new))}
        ${td(`${pcts(r.touch_rate)} ${delta(r.touch_rate, p?.touch_rate, "up")}`)}
        ${td(`${pcts(r.qual_rate)} ${delta(r.qual_rate, p?.qual_rate, "up")}`)}
        ${td(`${pc2(r.conversion)} ${delta(r.conversion, p?.conversion, "up")}`)}
        ${td(`${dash(r.unassigned)} ${delta(r.unassigned, p?.unassigned, "down")}`)}
        ${td(nf(r.qty))}</tr>`; }));

  const row = r => `
    <div class="runrow">
      <div><div class="p">${fdate(r.period_start)} – ${fdate(r.period_end)}</div>
        <div class="m">${nf(r.leads)} leads · ${pc2(r.conversion)} conversion · ${esc((r.files || []).join(", ")).slice(0, 90)}</div>
        ${r.by ? `<span class="who"><span class="av">·</span><span data-who="${esc(r.by)}">…</span></span>` : ""}</div>
      <div class="sp"><button class="btn ghost sm" data-open="${esc(r.run_id)}">Open</button>
        ${canWrite ? `<button class="btn ghost sm" data-del="${esc(r.run_id)}">Delete</button>` : ""}</div>
    </div>`;

  $("#runs").innerHTML = [...trend].reverse().map(row).join("");
  const recent = Store.runs.slice(0, 5);
  $("#recentWrap").hidden = !recent.length;
  $("#recent").innerHTML = recent.map(row).join("");
  paintNames();
}

async function openRun(id, btn) {
  const was = btn.textContent;
  btn.disabled = true; btn.textContent = "Opening…";
  try {
    const payload = await Store.open(id);
    IS_SAMPLE = false;
    load(payload);
    show("overview");
  } catch (err) {
    console.warn(err);
    toast("That run could not be opened — its saved detail is missing.");
  } finally { btn.disabled = false; btn.textContent = was; }
}

document.addEventListener("click", async e => {
  const o = e.target.closest("[data-open]"), d = e.target.closest("[data-del]");
  if (o) return openRun(o.dataset.open, o);
  if (d) {
    const r = Store.runs.find(x => x.run_id === d.dataset.del);
    const label = r ? `${fdate(r.period_start)} – ${fdate(r.period_end)}` : "this run";
    if (!window.confirm(`Delete the saved run for ${label}?` +
        (Store.mode === "shared" ? " It will disappear for everyone on the team." : ""))) return;
    d.disabled = true;
    try { await Store.remove(d.dataset.del); toast("Run deleted.", true); }
    catch (err) { console.warn(err); toast("That run could not be deleted."); d.disabled = false; }
  }
});

/* ═══════════════════════════════════════════════════ Excel download */

let excelBusy = false;
async function download(btn) {
  if (!DATA || excelBusy) return;
  excelBusy = true;
  const was = btn.textContent;
  btn.disabled = true; btn.textContent = "Building…";
  try {
    if (!window.ExcelJS) throw new Error("the spreadsheet library did not load");
    const per = DATA.metrics.period;
    const name = `Master_Lead_Report_${per.start}_to_${per.end}.xlsx`;
    const buf = await REP.build(window.ExcelJS, DATA.reconciliation, DATA.metrics,
                                DATA.insights, CFG);
    const blob = new Blob([buf], { type:
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const dl = await useCap("downloads");
    if (dl) {
      await dl.save({ filename: name, data: blob });
    } else {
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url; a.download = name; document.body.appendChild(a); a.click();
      a.remove(); setTimeout(() => URL.revokeObjectURL(url), 4000);
    }
  } catch (err) {
    if (err && err.code === "declined") { /* the viewer said no; nothing to report */ }
    else { console.error(err); toast(`The Excel report could not be built — ${err.message || err}`); }
  } finally {
    excelBusy = false; btn.disabled = false; btn.textContent = was;
  }
}
document.addEventListener("click", e => {
  const b = e.target.closest(".dl");
  if (b) download(b);
});

/* ═══════════════════════════════════════════════════ start */

Store.onChange(renderHistory);
renderHistory();
Store.init();
})();
