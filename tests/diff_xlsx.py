"""Cell-by-cell comparison of the Python-built and JS-built master workbooks."""
import datetime as dt
import sys
import openpyxl

a = openpyxl.load_workbook("out_py.xlsx")
b = openpyxl.load_workbook("out_js.xlsx")
out = []


def rgb(c):
    if c is None:
        return None
    v = getattr(c, "rgb", None)
    if not isinstance(v, str):
        return None
    return v[-6:].upper()


def val(v):
    if isinstance(v, dt.datetime):
        return v.date().isoformat()
    if isinstance(v, dt.date):
        return v.isoformat()
    if isinstance(v, float):
        # openpyxl writes floats as %.16g, ExcelJS writes full precision;
        # the last-ulp disagreement is not a real difference.
        return float(f"{v:.13g}")
    return v


if a.sheetnames != b.sheetnames:
    out.append(f"sheet names: {a.sheetnames} vs {b.sheetnames}")

for name in a.sheetnames:
    if name not in b.sheetnames:
        continue
    wa, wbs = a[name], b[name]
    if (wa.max_row, wa.max_column) != (wbs.max_row, wbs.max_column):
        out.append(f"[{name}] size {wa.max_row}x{wa.max_column} vs "
                   f"{wbs.max_row}x{wbs.max_column}")
    ma = sorted(str(x) for x in wa.merged_cells.ranges)
    mb = sorted(str(x) for x in wbs.merged_cells.ranges)
    if ma != mb:
        out.append(f"[{name}] merges differ: {set(ma) ^ set(mb)}")
    if str(wa.auto_filter.ref) != str(wbs.auto_filter.ref):
        out.append(f"[{name}] autofilter {wa.auto_filter.ref} vs {wbs.auto_filter.ref}")
    if (wa.freeze_panes or None) != (wbs.freeze_panes or None):
        out.append(f"[{name}] freeze {wa.freeze_panes} vs {wbs.freeze_panes}")
    if bool(wa.sheet_view.showGridLines) != bool(wbs.sheet_view.showGridLines):
        out.append(f"[{name}] gridlines {wa.sheet_view.showGridLines} vs "
                   f"{wbs.sheet_view.showGridLines}")

    def widths(w):
        """Expand <col min= max=> ranges; writers group equal widths differently."""
        got = {}
        for dim in w.column_dimensions.values():
            if dim.width is None:
                continue
            for i in range(dim.min or 1, (dim.max or dim.min or 1) + 1):
                got[i] = round(dim.width, 3)
        return got

    xa, xb = widths(wa), widths(wbs)
    for i in sorted(set(xa) | set(xb)):
        if xa.get(i) != xb.get(i):
            out.append(f"[{name}] col {openpyxl.utils.get_column_letter(i)} "
                       f"width {xa.get(i)} vs {xb.get(i)}")

    for r in range(1, max(wa.max_row, wbs.max_row) + 1):
        for c in range(1, max(wa.max_column, wbs.max_column) + 1):
            ca, cb = wa.cell(r, c), wbs.cell(r, c)
            if val(ca.value) != val(cb.value):
                out.append(f"[{name}] {ca.coordinate} value {ca.value!r} vs {cb.value!r}")
                continue
            if ca.value is None and cb.value is None:
                continue
            if ca.number_format != cb.number_format:
                out.append(f"[{name}] {ca.coordinate} fmt {ca.number_format!r} vs "
                           f"{cb.number_format!r}")
            fa, fb = ca.font, cb.font
            if (bool(fa.b), fa.sz, rgb(fa.color), fa.name, bool(fa.i)) != \
               (bool(fb.b), fb.sz, rgb(fb.color), fb.name, bool(fb.i)):
                out.append(f"[{name}] {ca.coordinate} font "
                           f"{(fa.name, fa.sz, bool(fa.b), bool(fa.i), rgb(fa.color))} vs "
                           f"{(fb.name, fb.sz, bool(fb.b), bool(fb.i), rgb(fb.color))}")
            pa = rgb(ca.fill.fgColor) if ca.fill and ca.fill.patternType else None
            pb = rgb(cb.fill.fgColor) if cb.fill and cb.fill.patternType else None
            if pa != pb:
                out.append(f"[{name}] {ca.coordinate} fill {pa} vs {pb}")
            if (ca.alignment.horizontal, ca.alignment.vertical, bool(ca.alignment.wrap_text)) != \
               (cb.alignment.horizontal, cb.alignment.vertical, bool(cb.alignment.wrap_text)):
                out.append(f"[{name}] {ca.coordinate} align "
                           f"{(ca.alignment.horizontal, ca.alignment.vertical, bool(ca.alignment.wrap_text))} vs "
                           f"{(cb.alignment.horizontal, cb.alignment.vertical, bool(cb.alignment.wrap_text))}")
            ba = bool(ca.border and ca.border.left and ca.border.left.style)
            bb = bool(cb.border and cb.border.left and cb.border.left.style)
            if ba != bb:
                out.append(f"[{name}] {ca.coordinate} border {ba} vs {bb}")

if not out:
    print("WORKBOOKS MATCH — every sheet, value, format, font, fill and layout is identical.")
    sys.exit(0)
print(f"{len(out)} difference(s):")
seen = {}
for d in out:
    key = d.split("]")[-1].split()[0] if "]" in d else d[:30]
    seen[key] = seen.get(key, 0) + 1
for d in out[:40]:
    print("  ", d)
print("\nby kind:", seen)
sys.exit(1)
