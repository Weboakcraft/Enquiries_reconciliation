# OakCraft Lead Reconciler

Upload your IndiaMART, JustDial and Meta exports. The software works out which format each
file is, reconciles all three into one lead register, and shows the complete analysis — master
report, funnel, employee performance, demand lost, unassigned leads, and an audit trail.

Runs entirely on your own computer. Nothing is uploaded anywhere.

---

## Chalane ka tareeka (how to run)

**Windows** — double-click **`run.bat`**
**Mac / Linux** — double-click **`run.sh`** (or in Terminal: `./run.sh`)

Pehli baar chalane par ye khud hi zaroori packages install kar lega, phir browser me
`http://127.0.0.1:5000` khul jayega. Band karne ke liye us black window me `Ctrl + C` dabayein.

**Zaroorat sirf ek cheez ki hai: Python 3.9 ya usse naya.**
Agar install nahi hai, [python.org/downloads](https://www.python.org/downloads/) se le lijiye —
Windows par install karte waqt **"Add Python to PATH"** wala box zaroor tick karein.

Manually chalana ho to:

```bash
pip install -r requirements.txt
python app.py
```

Port badalna ho (agar 5000 busy hai): `PORT=5055 python app.py`

---

## Istemaal kaise karein

1. **Upload & run** par teeno Excel files drag kar dein (ek saath ya alag-alag — dono chalega).
2. **Run reconciliation** dabayein. 400 leads ~2 second me ho jaati hain.
3. Left side ke menu se result dekhein. **Download Excel report** se poora workbook mil jayega.

Har run apne aap save ho jaata hai, isliye **History & trend** me aap pichhle hafton se compare
kar sakte hain — conversion badha ya ghata, touch rate sudhra ya nahi.

---

## What each screen is for

| Screen | What it answers |
|---|---|
| **Overview** | Headline numbers, the funnel, and the findings ranked by what they cost you |
| **Master report** | Your standard weekly format — employee wise, source wise, stage wise, lead quality |
| **Employees** | Kis user ne kitni leads uthai, unka stage kya hai, conversion kitna hua |
| **Sources & quality** | Which feed is worth the money — qualification rate, demand, hot leads |
| **Demand & loss** | Counting leads hides the size of what was lost. This counts **pieces**. |
| **Unassigned** | Leads that arrived with a blank salesperson. Nobody owns them, nobody works them. |
| **Reconciliation** | Proof that every row is accounted for. Control totals must balance. |
| **Exceptions** | Every data problem found, with the source row and a suggested fix |
| **History & trend** | Week-on-week movement of every headline metric |

---

## How the reconciliation works

This is the part that makes it reconciliation software rather than a dashboard: **every row that
goes in is accounted for on the way out.**

```
rows parsed  −  no date  −  out of period  −  duplicates  =  accepted
```

If that equation balances, the Reconciliation screen shows a green **Balanced** badge. If it does
not, something is wrong and you should not trust the numbers until it is fixed.

**Duplicates.** Only a matching 10-digit phone number merges two rows automatically — that is the
only signal strong enough to trust. A matching *name* does **not** merge them, because these
exports are full of placeholder names like `Jd Buyer` that dozens of different buyers share.
Name matches are reported for review and still counted. (On your 14–20 Sep data this matters:
merging on name would have destroyed 16 real leads.)

**Identity.** `UJALA RAJPUT` and `Ujala Rajput` are the same person. Blank salesperson becomes
`(UNASSIGNED)` and is counted and shown separately, never quietly dropped.

**Reporting period.** Taken as the 7 days ending on the newest row found. Anything outside that
window is excluded and listed, so a stray old row cannot distort the week.

---

## Adding a new export format

You do **not** need to change any Python. Open `config/sources.json` and add a profile:

```json
{
  "id": "my_new_crm",
  "label": "My New CRM",
  "match": { "required": ["Owner", "Status"], "any": ["Enquiry Date"] },
  "source_name": { "const": "MyCRM" },
  "fields": {
    "date":     { "from": ["Enquiry Date"] },
    "rep":      { "from": ["Owner"] },
    "customer": { "from": ["Client Name"] },
    "phone":    { "from": ["Mobile"] },
    "city":     { "from": ["Location"] },
    "product":  { "from": ["Item"] },
    "qty":      { "from": ["Pieces"] },
    "note":     { "from": ["Remarks"] }
  },
  "stage":   { "from": ["Status"],
               "map": { "OPEN": "New", "IN PROGRESS": "Contacted", "ORDER": "Won" },
               "blank": "New" },
  "quality": { "from": ["Rating"], "blank": "Not Rated" }
}
```

`match.required` is the fingerprint — a sheet is recognised when every one of those columns is
present (case, spaces and punctuation are ignored). Column names are matched loosely, so
`Assigned Salesperson`, `assigned_salesperson` and `ASSIGNED SALESPERSON` all work.

The same file also controls rep-name aliases, placeholder names, the words that count as blank,
and the rules that classify why a lead was lost. Save the file and re-run — no restart needed.

---

## Files

```
app.py                  Flask server and API
config/sources.json     Everything the engine knows about your formats — edit this, not the code
core/ingest.py          Format detection and normalisation to one lead model
core/reconcile.py       Control totals, identity resolution, duplicates, exception register
core/metrics.py         All the analytics, plus the findings generator
core/excel.py           Builds the downloadable master workbook
templates/ static/      The web interface (plain HTML, CSS and JavaScript — no build step)
data/runs/              One JSON file per run. This is your history — back it up.
data/exports/           Generated Excel reports
sample-data/            A synthetic previous week, so you can try the trend view immediately
selftest.py             Run the whole pipeline from the command line
```

Command-line check without the browser:

```bash
python selftest.py file1.xlsx file2.xlsx file3.xlsx
```

---

## Known limits

- **Deal value is not captured in any of your three feeds**, so the software can measure lead
  counts and piece quantities but *not* revenue, average order value or ROI per source. Add an
  Order Value field at Qualified stage and this becomes available automatically.
- Meta exports carry no lead-quality rating, so those leads show as *Not Rated*.
- The reporting period is a 7-day window. For a different range, the API accepts explicit
  `start` and `end` dates.
- `data/runs/` is plain JSON on this machine. Back that folder up — it is your history.
