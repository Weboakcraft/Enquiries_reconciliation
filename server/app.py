"""
OakCraft Lead Reconciler — local web app.

Run:   python app.py      then open http://127.0.0.1:5000

Nothing leaves this machine. Uploads are parsed in memory, results are written
to data/runs/ as JSON so week-on-week trends survive a restart.
"""
from __future__ import annotations

import datetime as dt
import json
import os
import tempfile
import traceback
import uuid
import webbrowser
from pathlib import Path
from threading import Timer

from flask import Flask, jsonify, render_template, request, send_file

from core import excel, ingest, metrics, reconcile

BASE = Path(__file__).resolve().parent
CONFIG = BASE / "config" / "sources.json"
RUNS = BASE / "data" / "runs"
EXPORTS = BASE / "data" / "exports"
RUNS.mkdir(parents=True, exist_ok=True)
EXPORTS.mkdir(parents=True, exist_ok=True)

ALLOWED = {".xlsx", ".xlsm", ".xls"}
MAX_MB = 40
PORT = int(os.environ.get("PORT", "5000"))

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_MB * 1024 * 1024


def load_cfg() -> dict:
    with open(CONFIG, encoding="utf-8") as fh:
        return json.load(fh)


def run_path(run_id: str) -> Path:
    return RUNS / f"{run_id}.json"


def list_runs() -> list[dict]:
    out = []
    for p in sorted(RUNS.glob("*.json"), reverse=True):
        try:
            with open(p, encoding="utf-8") as fh:
                d = json.load(fh)
            out.append({
                "run_id": d["run_id"], "created": d["created"],
                "period": d["metrics"]["period"],
                "leads": d["metrics"]["overall"]["total"],
                "new": d["metrics"]["overall"]["New"],
                "conversion": d["metrics"]["overall"]["conversion"],
                "touch_rate": d["metrics"]["overall"]["touch_rate"],
                "qual_rate": d["metrics"]["overall"]["qual_rate"],
                "unassigned": d["metrics"]["overall"]["unassigned"],
                "qty": d["metrics"]["overall"]["qty"],
                "files": [f["file"] for f in d["reconciliation"]["per_file"]],
            })
        except (KeyError, json.JSONDecodeError):
            continue
    return out


def build_trend(current: dict | None = None) -> list[dict]:
    """History oldest-first, so the UI can draw a week-on-week line."""
    rows = list_runs()
    if current:
        rows = [r for r in rows if r["run_id"] != current["run_id"]] + [current]
    seen, uniq = set(), []
    for r in sorted(rows, key=lambda x: x["period"]["end"]):
        k = (r["period"]["start"], r["period"]["end"])
        if k in seen:
            uniq = [u for u in uniq if (u["period"]["start"], u["period"]["end"]) != k]
        seen.add(k)
        uniq.append(r)
    return uniq


# ------------------------------------------------------------------ routes
@app.route("/")
def index():
    return render_template("index.html", max_mb=MAX_MB)


@app.route("/api/analyze", methods=["POST"])
def analyze():
    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "No files received."}), 400

    cfg = load_cfg()
    parsed, skipped = [], []
    tmpdir = tempfile.mkdtemp(prefix="oakrec_")
    try:
        for fs in files:
            name = os.path.basename(fs.filename or "unnamed")
            if Path(name).suffix.lower() not in ALLOWED:
                skipped.append({"file": name, "reason": "not an Excel file"})
                continue
            tmp = os.path.join(tmpdir, name)
            fs.save(tmp)
            try:
                res = ingest.read_file(tmp, name, cfg)
            except Exception as exc:                       # a bad file must not kill the run
                skipped.append({"file": name, "reason": f"could not be read: {exc}"})
                continue
            if not res["leads"]:
                detected = ", ".join(s["label"] for s in res["sheets"]) or "nothing"
                skipped.append({"file": name,
                                "reason": f"no recognisable lead rows (detected: {detected})"})
                continue
            parsed.append(res)

        if not parsed:
            return jsonify({"error": "None of the uploaded files matched a known format.",
                            "skipped": skipped}), 422

        period = None
        if request.form.get("start") and request.form.get("end"):
            period = (request.form["start"], request.form["end"])

        rec = reconcile.reconcile(parsed, cfg, period)
        m = metrics.compute(rec, cfg)
        insights = metrics.build_insights(m, rec, cfg)

        run_id = dt.datetime.now().strftime("%Y%m%d-%H%M%S-") + uuid.uuid4().hex[:4]
        payload = {
            "run_id": run_id,
            "created": dt.datetime.now().isoformat(timespec="seconds"),
            "skipped": skipped,
            "reconciliation": rec,
            "metrics": m,
            "insights": insights,
        }
        with open(run_path(run_id), "w", encoding="utf-8") as fh:
            json.dump(payload, fh, separators=(",", ":"))

        summary = {"run_id": run_id, "created": payload["created"],
                   "period": m["period"], "leads": m["overall"]["total"],
                   "new": m["overall"]["New"],
                   "conversion": m["overall"]["conversion"],
                   "touch_rate": m["overall"]["touch_rate"],
                   "qual_rate": m["overall"]["qual_rate"],
                   "unassigned": m["overall"]["unassigned"], "qty": m["overall"]["qty"],
                   "files": [f["file"] for f in rec["per_file"]]}

        light = dict(payload)
        light["reconciliation"] = {k: v for k, v in rec.items() if k != "excluded"}
        light["trend"] = build_trend(summary)
        return jsonify(light)

    except Exception as exc:
        traceback.print_exc()
        return jsonify({"error": f"Analysis failed: {exc}"}), 500


@app.route("/api/runs")
def runs():
    return jsonify({"runs": list_runs(), "trend": build_trend()})


@app.route("/api/run/<run_id>")
def get_run(run_id):
    p = run_path(run_id)
    if not p.exists():
        return jsonify({"error": "Run not found."}), 404
    with open(p, encoding="utf-8") as fh:
        d = json.load(fh)
    d["reconciliation"].pop("excluded", None)
    d["trend"] = build_trend()
    return jsonify(d)


@app.route("/api/run/<run_id>", methods=["DELETE"])
def delete_run(run_id):
    p = run_path(run_id)
    if p.exists():
        p.unlink()
        return jsonify({"ok": True})
    return jsonify({"error": "Run not found."}), 404


@app.route("/api/export/<run_id>")
def export(run_id):
    p = run_path(run_id)
    if not p.exists():
        return jsonify({"error": "Run not found."}), 404
    with open(p, encoding="utf-8") as fh:
        d = json.load(fh)
    cfg = load_cfg()
    per = d["metrics"]["period"]
    out = EXPORTS / f'Master_Lead_Report_{per["start"]}_to_{per["end"]}.xlsx'
    excel.build(d["reconciliation"], d["metrics"], d["insights"], cfg, str(out))
    return send_file(out, as_attachment=True, download_name=out.name)


@app.route("/api/config")
def get_config():
    return jsonify(load_cfg())


@app.errorhandler(413)
def too_large(_):
    return jsonify({"error": f"File too large. Limit is {MAX_MB} MB per upload."}), 413


def open_browser():
    webbrowser.open_new(f"http://127.0.0.1:{PORT}/")


if __name__ == "__main__":
    print("\n  OakCraft Lead Reconciler")
    print("  ---------------------------------------------")
    print(f"  Running at http://127.0.0.1:{PORT}")
    print("  Data stays on this machine. Ctrl+C to stop.\n")
    if os.environ.get("WERKZEUG_RUN_MAIN") != "true":
        Timer(1.2, open_browser).start()
    app.run(debug=False, port=PORT)
