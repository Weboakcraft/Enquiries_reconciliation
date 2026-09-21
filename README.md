# OakCraft Lead Reconciler

Upload the IndiaMART, JustDial and Meta exports. The software works out which
format each file is, reconciles all three into one lead register, and shows the
complete analysis — master report, funnel, employee performance, demand lost,
unassigned leads, and an audit trail that proves every row is accounted for.

There are two builds in this repository. They produce **identical numbers** —
the browser engine is a line-by-line port of the Python one, and `tests/` exists
to prove it on every change.

| | `web/` — browser build | `server/` — Flask build |
|---|---|---|
| Needs a server | no | yes |
| Where files are parsed | in the tab you opened | on the server |
| Run history | shared across the team, or this browser | JSON files on the server's disk |
| Hosting cost | none (static) | a small always-on instance |
| Access control | whoever the page is shared with | **none built in** — see `server/DEPLOY.md` |

## The hosted site

The live site is published from this repository's `web/` build and is private to
the OakCraft workspace:

<https://claude.ai/artifact/NANATp1BsLmd5WJFHcsbeA>

It opens for anyone the page is shared with — no install, no Python, works on a
phone. Because everything runs in the browser, **the spreadsheets are never
uploaded anywhere**; only the saved run summaries and the reconciled figures
travel, to the team's shared history, so everyone sees the same week-on-week
movement.

There is a **See it with a sample week** button on the first screen. It builds an
invented week from a fixed seed and runs it through the real pipeline, so you can
see every screen working before you upload anything. It is labelled as a sample
throughout and is never saved to history.

## Using it

1. Drop the three exports on **Upload & run** — together or one at a time.
2. Press **Run reconciliation**. Four hundred leads take about two seconds.
3. Read the results from the left-hand menu. **Download Excel report** gives you
   the full workbook: master report, reconciliation, findings, unassigned leads,
   demand and loss, and the complete reconciled register.

Each run is saved, so **History & trend** compares this week against the last
ones — whether conversion moved, whether the touch rate improved.

## What each screen is for

| Screen | What it answers |
|---|---|
| **Overview** | Headline numbers, the funnel, and the findings ranked by what they cost you |
| **Master report** | Your standard weekly format — employee wise, source wise, stage wise, lead quality |
| **Employees** | Who picked up how many leads, what stage they are at, what converted |
| **Sources & quality** | Which feed is worth the money — qualification rate, demand, hot leads |
| **Demand & loss** | Counting leads hides the size of what was lost. This counts **pieces**. |
| **Unassigned** | Leads that arrived with a blank salesperson. Nobody owns them, nobody works them. |
| **Reconciliation** | Proof that every row is accounted for. Control totals must balance. |
| **Exceptions** | Every data problem found, with the source row and a suggested fix |
| **History & trend** | Week-on-week movement of every headline metric |

## How the reconciliation works

This is what makes it reconciliation software rather than a dashboard: **every
row that goes in is accounted for on the way out.**

```
rows parsed  −  no date  −  out of period  −  duplicates  =  accepted
```

If that equation balances, the Reconciliation screen shows a green **Balanced**
badge. If it does not, something is wrong and the numbers should not be trusted
until it is fixed.

**Duplicates.** Only a matching 10-digit phone number merges two rows
automatically — that is the only signal strong enough to trust. A matching *name*
does **not** merge them, because these exports are full of placeholder names like
`Jd Buyer` that dozens of different buyers share. Name matches are reported for
review and still counted.

**Identity.** `UJALA RAJPUT` and `Ujala Rajput` are the same person. A blank
salesperson becomes `(UNASSIGNED)` and is counted and shown separately, never
quietly dropped.

**Reporting period.** The 7 days ending on the newest row found. Anything outside
that window is excluded and listed, so a stray old row cannot distort the week.

## Adding a new export format

You do not need to change any code. Open `server/config/sources.json` and add a
profile — `match.required` is the fingerprint, and a sheet is recognised when
every one of those columns is present. Column names are matched loosely, so
`Assigned Salesperson`, `assigned_salesperson` and `ASSIGNED SALESPERSON` all
work. The same file also controls rep-name aliases, placeholder names, the words
that count as blank, and the rules that classify why a lead was lost.

Both builds read that one file. The Flask app picks it up on the next run; for
the browser build, rebuild the page:

```bash
node web-src/build.js
```

## Repository layout

```
web/index.html        The built site. This is what GitHub Pages serves.
web-src/              What it is built from — edit these, then run build.js
  engine.js             format detection, reconciliation, metrics, findings
  report.js             the Excel workbook builder
  xlsxread.js           workbook -> cell matrices
  style.css  body.html  app.js
  build.js              assembles web/index.html
server/               The Flask build, plus Dockerfile, Procfile and render.yaml
  config/sources.json   everything the engine knows about your formats
  core/                 the Python engine the browser engine is ported from
  DEPLOY.md             how to host it, and the access-control warning
tests/                Proves the two engines agree. See tests/README.md.
```

## Publishing the site to GitHub Pages

`.github/workflows/pages.yml` deploys `web/` on every push to `main`. It needs
one setting turned on first: **Settings → Pages → Build and deployment →
Source: GitHub Actions**. The site then appears at
`https://<owner>.github.io/<repo>/`.

Note that GitHub Pages sites are public. The page holds no data of its own — the
spreadsheets stay in whoever's browser opened them — but the shared team history
needs a signed-in workspace, so a Pages copy falls back to saving runs in each
person's own browser.

## Known limits

- **No deal value is captured in any of the three feeds**, so the software
  measures lead counts and piece quantities but *not* revenue, average order
  value or ROI per source. Add an Order Value field at Qualified stage and this
  becomes available automatically.
- Meta exports carry no lead-quality rating, so those leads show as *Not Rated*.
- The reporting period is a 7-day window.
