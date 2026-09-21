"""CLI self-test: runs the whole pipeline on files given as arguments."""
import json, sys
from pathlib import Path
from core import ingest, reconcile, metrics, excel

cfg = json.load(open("config/sources.json", encoding="utf-8"))
paths = sys.argv[1:]
parsed = []
for p in paths:
    r = ingest.read_file(p, Path(p).name, cfg)
    parsed.append(r)
    for s in r["sheets"]:
        print(f'  detect: {Path(p).name:<52} sheet "{s["sheet"]}" -> {s["label"]:<28} rows={s["rows"]}')
rec = reconcile.reconcile(parsed, cfg)
m = metrics.compute(rec, cfg)
ins = metrics.build_insights(m, rec, cfg)
t = rec["control_totals"]
print("\nCONTROL TOTALS")
for k in ["rows_read","rows_blank","parsed","no_date","out_of_period","duplicates","accepted"]:
    print(f'   {k:<16} {t[k]:>5}')
print(f'   balanced         {t["balanced"]}')
print(f'\nPERIOD  {m["period"]["start"]} -> {m["period"]["end"]}')
o = m["overall"]
print(f'TOTAL {o["total"]} | new {o["New"]} | contacted {o["Contacted"]} | qual {o["Qualified"]} '
      f'| quoted {o["Quoted / Proposal"]} | won {o["Won"]} | lost {o["Lost"]} | notqual {o["Not Qualified"]}')
print(f'conversion {o["conversion"]}% | touch {o["touch_rate"]}% | qual-rate {o["qual_rate"]}% | qty {o["qty"]:,}')
print("\nBY REP")
for r_, b in m["by_rep"].items():
    print(f'   {r_:<16} {b["total"]:>4} | new {b["New"]:>3} | qual+ {b["qplus"]:>3} | won {b["Won"]} '
          f'| touch {b["touch_rate"]:>5}% | conv {b["conversion"]:>5}%')
print("\nBY SOURCE")
for s, b in m["by_source"].items():
    print(f'   {s:<12} {b["total"]:>4} | unass {b["unassigned"]:>3} | qual-rate {b["qual_rate"]:>5}% '
          f'| qty {b["qty"]:>7,} | hot {b["hot"]:>3}')
print("\nEXCEPTIONS")
for e in rec["exception_summary"]:
    print(f'   {e["severity"]:<9} {e["kind"]:<18} {e["count"]:>4}')
print(f'\nDEMAND {m["demand_totals"]}')
print(f'\nINSIGHTS ({len(ins)})')
for i in ins: print(f'   [{i["sev"]:<8}] {i["title"]}')
out = excel.build(rec, m, ins, cfg, "data/exports/_selftest.xlsx")
print(f'\nExcel written: {out}')
