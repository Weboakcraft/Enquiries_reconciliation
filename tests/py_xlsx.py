import json
import sys
from pathlib import Path
SERVER = Path(__file__).resolve().parent.parent / "server"
sys.path.insert(0, str(SERVER))
from core import ingest, reconcile, metrics, excel
cfg = json.load(open(SERVER / "config/sources.json", encoding="utf-8"))
parsed = [ingest.read_file(p, Path(p).name, cfg) for p in sys.argv[1:]]
rec = reconcile.reconcile(parsed, cfg)
m = metrics.compute(rec, cfg)
ins = metrics.build_insights(m, rec, cfg)
excel.build(rec, m, ins, cfg, "out_py.xlsx")
print("wrote out_py.xlsx")
