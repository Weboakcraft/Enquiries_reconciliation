# Parity tests

The browser engine in `web-src/` is a port of the Python engine in `server/core/`.
These tests exist to prove the port did not change any number. They run both
engines over the same workbooks and diff the results field by field.

```bash
cd tests
npm install                       # xlsx, exceljs, playwright
pip install openpyxl flask
python3 gen_testdata.py           # three synthetic exports, fixed seed
```

`gen_testdata.py` deliberately includes the awkward cases: rows with no date,
rows from outside the week, the same phone number arriving through two sources,
blank salespeople, placeholder buyer names, `"null"` written into a numeric
quantity column, stage values that are not in the mapping, and the same person
spelled two ways.

### 1. Engine parity

```bash
F="sample_meta.xlsx sample_indiamart.xlsx sample_justdial.xlsx"
python3 py_dump.py $F        # -> out_py.json
node    js_dump.js $F        # -> out_js.json
python3 diff.py
```

Expected: `IDENTICAL — the JavaScript engine reproduces the Python engine exactly.`

This compares everything: control totals, the exception register, every block of
metrics, the ageing histogram, dispositions, demand, and the generated findings
text.

### 2. Workbook parity

```bash
python3 py_xlsx.py $F        # -> out_py.xlsx  (openpyxl)
node    js_xlsx.js $F        # -> out_js.xlsx  (exceljs)
python3 diff_xlsx.py
```

Expected: `WORKBOOKS MATCH — every sheet, value, format, font, fill and layout is
identical.` It checks cell values, number formats, fonts, fills, alignment,
borders, merged ranges, column widths, freeze panes and autofilters.

Two deliberate tolerances: `openpyxl` writes floats with `%.16g` while `exceljs`
writes full precision, and the two writers group equal column widths
differently. Neither changes what Excel shows.

### 3. The page itself

```bash
node e2e.js
```

Drives `web/index.html` in headless Chromium, in both light and dark mode:
loads the sample week, visits all nine screens, exercises the source and
severity filters, downloads the Excel report and checks its filename, uploads
the three generated workbooks and checks the reconciled totals, confirms the run
was saved to history and reopens identically, checks the control totals balance,
checks there is no horizontal scroll at 390px, and fails on any console error.

The CDN hosts are rewritten to local copies of the same library versions so the
test does not depend on the network.

### One intentional difference from the Python

Digit grouping in the generated findings text. The Python used `{:,}`
(`1,234,567`); the browser build groups the Indian way (`12,34,567`) to match
every other number on screen. `diff.py` strips grouping commas before comparing,
so the underlying figures are still compared exactly.
