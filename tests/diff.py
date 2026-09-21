"""Deep-compares the Python and JavaScript engine outputs."""
import json
import re
import sys

GROUP = re.compile(r"(?<=\d)[,](?=\d)")


def norm(x, path=""):
    """Strip digit-grouping commas from generated prose, and internal keys."""
    if isinstance(x, dict):
        return {k: norm(v, f"{path}.{k}") for k, v in x.items() if not k.startswith("_")}
    if isinstance(x, list):
        return [norm(v, f"{path}[]") for v in x]
    if isinstance(x, str):
        return GROUP.sub("", x)
    if isinstance(x, float) and x.is_integer():
        return int(x)
    return x


def walk(a, b, path, out):
    if type(a) is not type(b) and not (isinstance(a, (int, float)) and isinstance(b, (int, float))):
        out.append(f"{path}: type {type(a).__name__} vs {type(b).__name__}")
        return
    if isinstance(a, dict):
        for k in sorted(set(a) | set(b)):
            if k not in a:
                out.append(f"{path}.{k}: missing in PY")
            elif k not in b:
                out.append(f"{path}.{k}: missing in JS")
            else:
                walk(a[k], b[k], f"{path}.{k}", out)
    elif isinstance(a, list):
        if len(a) != len(b):
            out.append(f"{path}: length {len(a)} vs {len(b)}")
        for i in range(min(len(a), len(b))):
            walk(a[i], b[i], f"{path}[{i}]", out)
    else:
        if isinstance(a, float) or isinstance(b, float):
            if abs(float(a) - float(b)) > 1e-9:
                out.append(f"{path}: {a!r} vs {b!r}")
        elif a != b:
            out.append(f"{path}: {a!r} vs {b!r}")


py = norm(json.load(open("out_py.json")))
js = norm(json.load(open("out_js.json")))

diffs = []
walk(py, js, "", diffs)
if not diffs:
    print("IDENTICAL — the JavaScript engine reproduces the Python engine exactly.")
    sys.exit(0)
print(f"{len(diffs)} difference(s):")
for d in diffs[:60]:
    print("  ", d)
sys.exit(1)
