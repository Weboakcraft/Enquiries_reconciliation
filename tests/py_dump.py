"""Run the original Python pipeline and dump the full result as JSON."""
import json
import sys
from pathlib import Path
SERVER = Path(__file__).resolve().parent.parent / "server"
sys.path.insert(0, str(SERVER))
from core import ingest, reconcile, metrics
cfg = json.load(open(SERVER / "config/sources.json", encoding="utf-8"))
parsed = [ingest.read_file(p, Path(p).name, cfg) for p in sys.argv[1:]]
rec = reconcile.reconcile(parsed, cfg)
m = metrics.compute(rec, cfg)
ins = metrics.build_insights(m, rec, cfg)
json.dump({"reconciliation": rec, "metrics": m, "insights": ins},
          open("out_py.json", "w"), indent=1, sort_keys=True, default=str)
print("rows", rec["control_totals"], "\nperiod", rec["period"], "\nleads", len(rec["leads"]))
