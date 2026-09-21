/* OakCraft Lead Reconciler — front end. Plain ES2020, no build step, no CDN. */
(() => {
"use strict";

const $  = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];
const esc = s => String(s ?? "").replace(/[&<>"]/g, c =>
  ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const nf  = n => (Number(n) || 0).toLocaleString("en-IN");
const pcts = v => (v === 0 || v == null) ? "—" : `${Number(v).toFixed(1)}%`;
const pc2 = v => `${Number(v || 0).toFixed(2)}%`;
const dash = v => (v === 0 || v == null) ? "—" : nf(v);
const fdate = s => { const d = new Date(s + "T00:00:00");
  return isNaN(d) ? s : d.toLocaleDateString("en-GB", { day: "2-digit", month: "short" }); };

const STAGE_VAR = { "New": "--s-new", "Contacted": "--s-con", "Qualified": "--s-qual",
  "Quoted / Proposal": "--s-prop", "Won": "--s-won", "Lost": "--s-lost", "Not Qualified": "--s-nq" };
const STAGE_SHORT = { "New": "Never worked", "Quoted / Proposal": "Quoted" };
const short = s => STAGE_SHORT[s] || s;
const sv = s => `var(${STAGE_VAR[s] || "--muted"})`;
const QVAR = q => q === "Hot Lead" ? "--s-won" : q === "Warm Lead" ? "--s-new"
  : q === "Cold Lead" ? "--s-con" : q === "Not Rated" ? "--muted" : "--s-lost";

let DATA = null, FILES = [], SRC = "ALL", EXSEV = "ALL";

/* ───────────────────────────── navigation ───────────────────────────── */
function show(view) {
  $$("main > section").forEach(s => s.hidden = s.id !== "v-" + view);
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
function toast(msg) {
  const t = document.createElement("div");
  t.className = "toast"; t.textContent = msg;
  document.body.appendChild(t);
  setTimeout(() => t.remove(), 6500);
}

/* ───────────────────────────── upload ───────────────────────────── */
const drop = $("#drop"), fileInput = $("#file");
$("#browse").onclick = () => fileInput.click();
drop.onclick = e => { if (e.target === drop || e.target.closest(".icon,h3,p")) fileInput.click(); };
fileInput.onchange = () => addFiles([...fileInput.files]);
["dragenter", "dragover"].forEach(ev => drop.addEventListener(ev, e => {
  e.preventDefault(); drop.classList.add("hot");
}));
["dragleave", "drop"].forEach(ev => drop.addEventListener(ev, e => {
  e.preventDefault(); if (ev === "dragleave" && drop.contains(e.relatedTarget)) return;
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
$("#clear").onclick = () => { FILES = []; fileInput.value = ""; renderFiles(); };

$("#run").onclick = async () => {
  if (!FILES.length) return;
  const fd = new FormData();
  FILES.forEach(f => fd.append("files", f));
  $("#run").disabled = true; $("#prog").hidden = false;
  $("#status").textContent = "Reading files, detecting formats, reconciling…";
  try {
    const r = await fetch("/api/analyze", { method: "POST", body: fd });
    const j = await r.json();
    if (!r.ok) {
      toast(j.error || "Analysis failed.");
      if (j.skipped?.length) toast(j.skipped.map(s => `${s.file}: ${s.reason}`).join(" · "));
      return;
    }
    load(j);
    show("overview");
  } catch (err) {
    toast("Could not reach the local server. Is app.py still running?");
  } finally {
    $("#run").disabled = false; $("#prog").hidden = true; $("#status").textContent = "";
  }
};

/* ───────────────────────────── load a run ───────────────────────────── */
function load(j) {
  DATA = j;
  unlock(true);
  const m = j.metrics;
  $("#runLabel").textContent = `${m.period.start} → ${m.period.end} · ${m.overall.total} leads`;
  $("#navUn").textContent = m.overall.unassigned || "";
  $("#navEx").textContent = j.reconciliation.exceptions.length || "";
  if (j.skipped?.length) j.skipped.forEach(s => toast(`${s.file} skipped — ${s.reason}`));
  renderOverview(); renderMaster(); renderPeople(); renderSources();
  renderDemand(); renderUnassigned(); renderRecon(); renderExceptions();
  renderHistory(j.trend || []);
}

/* ───────────────────────────── overview ───────────────────────────── */
function renderOverview() {
  const m = DATA.metrics, o = m.overall, rec = DATA.reconciliation;
  $("#ovPeriod").textContent =
    `${m.period.start} to ${m.period.end} · ${rec.per_file.length} file(s) · ` +
    `${m.sources.join(", ")} · reconciled ${new Date(DATA.created).toLocaleString("en-GB")}`;
  $("#asof").textContent = "as on " + fdate(m.asof);

  const crit = rec.exceptions.filter(e => e.severity === "critical").length;
  $("#ovKpis").innerHTML = [
    ["", "Leads reconciled", nf(o.total), `${rec.control_totals.rows_read} rows read in`],
    ["warn", "Never worked", nf(o.New), `${pcts(100 * o.New / o.total)} · ${m.aged_untouched} aged 2+ days`],
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
  if (!ks.length) { $("#ageing").innerHTML = `<div class="empty">Nothing untouched.</div>`; return; }
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

/* ───────────────────────────── tables ───────────────────────────── */
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
      return `<tr>${td(esc(s), "nm")}${td(b.total)}${td(pcts(100 * b.total / o.total))}
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

function renderPeople() { drawReps();
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
      return `<tr>${td(esc(s), "nm")}${td(`<b>${b.total}</b>`)}${td(pcts(100 * b.total / o.total))}
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
      || [`<tr><td class="l">Nothing to recover.</td></tr>`]);
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

/* ───────────────────────────── reconciliation ───────────────────────────── */
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
      || [`<tr><td class="l">Nothing at this severity.</td></tr>`]);
}
$("#exFilter").addEventListener("click", e => {
  const b = e.target.closest("[data-sev]"); if (!b) return;
  EXSEV = b.dataset.sev;
  $$("#exFilter button").forEach(x => x.setAttribute("aria-pressed", String(x.dataset.sev === EXSEV)));
  drawExceptions();
});

/* ───────────────────────────── history ───────────────────────────── */
function renderHistory(trend) {
  if (!trend.length) {
    $("#tTrend").innerHTML = `<tbody><tr><td class="l"><div class="empty">
      <h3>No history yet</h3>Run a second week and the comparison appears here.</div></td></tr></tbody>`;
    $("#runs").innerHTML = "";
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
      return `<tr>${td(`${fdate(r.period.start)} – ${fdate(r.period.end)}`, "nm")}
        ${td(nf(r.leads))}${td(dash(r.new))}
        ${td(`${pcts(r.touch_rate)} ${delta(r.touch_rate, p?.touch_rate, "up")}`)}
        ${td(`${pcts(r.qual_rate)} ${delta(r.qual_rate, p?.qual_rate, "up")}`)}
        ${td(`${pc2(r.conversion)} ${delta(r.conversion, p?.conversion, "up")}`)}
        ${td(`${dash(r.unassigned)} ${delta(r.unassigned, p?.unassigned, "down")}`)}
        ${td(nf(r.qty))}</tr>`; }));

  $("#runs").innerHTML = [...trend].reverse().map(r => `
    <div class="runrow">
      <div><div class="p">${fdate(r.period.start)} – ${fdate(r.period.end)}</div>
        <div class="m">${r.leads} leads · ${pc2(r.conversion)} conversion · ${esc(r.files.join(", ")).slice(0, 90)}</div></div>
      <div class="sp"><button class="btn ghost sm" data-open="${r.run_id}">Open</button>
        <button class="btn ghost sm" data-del="${r.run_id}">Delete</button></div>
    </div>`).join("");
}
$("#runs").addEventListener("click", async e => {
  const o = e.target.closest("[data-open]"), d = e.target.closest("[data-del]");
  if (o) {
    const r = await fetch("/api/run/" + o.dataset.open);
    if (r.ok) { load(await r.json()); show("overview"); } else toast("Could not open that run.");
  }
  if (d) {
    await fetch("/api/run/" + d.dataset.del, { method: "DELETE" });
    refreshRuns();
  }
});

async function refreshRuns() {
  try {
    const r = await fetch("/api/runs");
    if (!r.ok) return;
    const j = await r.json();
    renderHistory(j.trend || []);
    $("#recentWrap").hidden = !j.runs.length;
    $("#recent").innerHTML = j.runs.slice(0, 5).map(x => `
      <div class="runrow"><div><div class="p">${fdate(x.period.start)} – ${fdate(x.period.end)}</div>
        <div class="m">${x.leads} leads · ${pc2(x.conversion)} conversion</div></div>
        <div class="sp"><button class="btn ghost sm" data-openr="${x.run_id}">Open</button></div></div>`).join("");
  } catch { /* server not up yet */ }
}
$("#recent").addEventListener("click", async e => {
  const b = e.target.closest("[data-openr]"); if (!b) return;
  const r = await fetch("/api/run/" + b.dataset.openr);
  if (r.ok) { load(await r.json()); show("overview"); }
});

const download = () => { if (DATA) window.location = "/api/export/" + DATA.run_id; };
$("#dl").onclick = download; $("#dl2").onclick = download;

refreshRuns();
})();
