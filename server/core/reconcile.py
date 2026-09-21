"""
Reconciliation.

This is what separates the tool from a dashboard: every row that enters is
accounted for on the way out. Nothing is silently dropped, and every judgement
call the engine makes is written down as an exception the user can inspect.

Three jobs:
  1. Control totals  - rows read -> parsed -> in period -> deduped -> accepted.
  2. Identity        - one person, one name; blank owner becomes (UNASSIGNED).
  3. Exceptions      - typed, severity-ranked register of everything suspect.
"""
from __future__ import annotations

import datetime as dt
import re
from collections import Counter, defaultdict


def _norm_name(s: str) -> str:
    return re.sub(r"[^a-z0-9]", "", (s or "").lower())


def _period_of(leads: list[dict]) -> tuple[str | None, str | None]:
    ds = sorted(d for d in (l["date"] for l in leads) if d)
    return (ds[0], ds[-1]) if ds else (None, None)


def infer_period(leads: list[dict]) -> tuple[str, str]:
    """
    Pick the reporting window. Most rows cluster in one week; stragglers from an
    earlier export are excluded rather than allowed to distort the week.
    Rule: take the latest date, walk back 6 days -> a 7-day window.
    """
    ds = sorted(d for d in (l["date"] for l in leads) if d)
    if not ds:
        today = dt.date.today()
        return (today - dt.timedelta(days=6)).isoformat(), today.isoformat()
    end = dt.date.fromisoformat(ds[-1])
    start = end - dt.timedelta(days=6)
    return start.isoformat(), end.isoformat()


def reconcile(files: list[dict], cfg: dict,
              period: tuple[str, str] | None = None) -> dict:
    """files = list of read_file() results. Returns the full reconciliation record."""
    unassigned = cfg["canonical"]["unassigned_label"]
    all_leads: list[dict] = []
    for f in files:
        all_leads.extend(f["leads"])

    p_start, p_end = period or infer_period(all_leads)
    exceptions: list[dict] = []

    def ex(kind, severity, message, detail=None, lead=None, fix=None):
        exceptions.append({
            "kind": kind, "severity": severity, "message": message,
            "detail": detail or "", "fix": fix or "",
            "ref": (f'{lead["src_file"]} · {lead["src_sheet"]} · row {lead["src_row"]}'
                    if lead else ""),
            "customer": lead.get("customer", "") if lead else "",
        })

    # ------------------------------------------------------------ 1. period filter
    in_period, out_period, no_date = [], [], []
    for l in all_leads:
        if not l["date"]:
            no_date.append(l)
            ex("no_date", "high", "Row has no usable date and was excluded",
               f'customer "{l["customer"] or "(blank)"}"', l,
               "Check the date column format in the export")
        elif p_start <= l["date"] <= p_end:
            in_period.append(l)
        else:
            out_period.append(l)
            ex("out_of_period", "low",
               f'Row dated {l["date"]} is outside the {p_start} to {p_end} window',
               f'{l["source"]} · {l["customer"] or "(blank)"}', l,
               "Filter the export to the reporting week")

    # ------------------------------------------------------------ 2. duplicates
    #
    # Only a matching 10-digit phone number is strong enough to merge two rows
    # automatically. A matching *name* is not: these exports are full of
    # placeholder names ("Jd Buyer" and the like) that dozens of different
    # buyers share, so merging on name would quietly destroy real leads.
    # Name matches are therefore reported for review and still counted.
    keep, dropped = [], []
    seen_phone: dict[str, dict] = {}
    seen_name: dict[tuple, dict] = {}
    placeholders = {p.lower() for p in cfg.get("placeholder_names", {}).get("values", [])}

    for l in sorted(in_period, key=lambda x: (x["date"], x["src_file"], x["src_row"])):
        ph = l["phone"]
        dup_of = seen_phone.get(ph) if ph and len(ph) == 10 else None

        if dup_of is not None:
            l["_dup_of"] = f'{dup_of["src_file"]} row {dup_of["src_row"]}'
            l["_dup_kind"] = "phone"
            dropped.append(l)
            cross = dup_of["source"] != l["source"]
            ex("duplicate", "high" if cross else "medium",
               ("Same phone number reached us through two different sources" if cross
                else "Same phone number appears twice in one source"),
               f'"{l["customer"]}" · {ph} · {dup_of["source"]} row {dup_of["src_row"]} '
               f'vs {l["source"]} row {l["src_row"]}', l,
               "Counted once only. De-duplicate at source to keep totals clean")
            continue

        if ph and len(ph) == 10:
            seen_phone[ph] = l

        # name similarity -> flag only, never drop
        name = (l["customer"] or "").strip().lower()
        if name and name not in placeholders and not ph:
            nk = (_norm_name(l["customer"]), l["source"])
            prev = seen_name.get(nk)
            if prev and abs((dt.date.fromisoformat(l["date"])
                             - dt.date.fromisoformat(prev["date"])).days) <= 3:
                ex("possible_duplicate", "low",
                   "Two rows share a customer name within 3 days — both kept, please confirm",
                   f'"{l["customer"]}" · {l["source"]} row {l["src_row"]} and row '
                   f'{prev["src_row"]}. No phone number on either row to confirm with.', l,
                   "Capture a phone number so duplicates can be detected reliably")
            else:
                seen_name[nk] = l

        keep.append(l)

    accepted = keep

    # ------------------------------------------------------------ 3. data-quality exceptions
    closed_no_order = set(cfg["canonical"]["closed_no_order"])
    for l in accepted:
        if not l["assigned"]:
            ex("unassigned", "critical",
               "Lead has no salesperson assigned",
               f'{l["source"]} · {l["customer"] or "(blank)"} · {l["qty"] or 0} pcs', l,
               "Make Salesperson mandatory at lead capture")
        if l["qty_defect"]:
            ex("qty_null", "medium",
               'Quantity field contains text instead of a number',
               f'{l["source"]} · {l["customer"]}', l,
               "Stop writing 'null' into a numeric field")
        if l["stage"] in closed_no_order and l["disposition"]["id"] == "not_logged":
            ex("no_reason", "medium",
               f'Closed as {l["stage"]} with no reason recorded',
               f'{l["source"]} · {l["customer"]}', l,
               "Make Loss Reason mandatory before closing a lead")
        if l["disposition"]["severity"] == "recoverable" and l["stage"] in closed_no_order:
            ex("recoverable", "high",
               "Closed without an order although the note records live interest",
               f'{l["qty"]} pcs · {l["customer"]} · note: "{l["note"]}"', l,
               "Re-open and chase")
        if (l["quality"] in ("Hot Lead", "Warm Lead")) and l["stage"] in closed_no_order:
            ex("hot_lost", "high",
               f'Lead rated {l["quality"]} was closed as {l["stage"]}',
               f'{l["qty"]} pcs · {l["customer"]}', l,
               "Review with the rep before writing off")
        if l["quality"] == "Not Interested" and re.search(r"requir", (l["note"] or "").lower()) \
                and not re.search(r"not\s*requ|no\s*requ", (l["note"] or "").lower()):
            ex("contradiction", "medium",
               'Rated "Not Interested" but the note records a requirement',
               f'{l["customer"]} · note: "{l["note"]}"', l,
               'Block "Not Interested" when a requirement is logged')

    for f in files:
        for u in f["unmapped_stage_values"]:
            exceptions.append({
                "kind": "unmapped_stage", "severity": "high",
                "message": f'Stage value "{u["value"]}" is not in the mapping and defaulted to New',
                "detail": f'field {u["field"]}', "fix": "Add this value to config/sources.json",
                "ref": f'{u["file"]} · {u["sheet"]} · row {u["row"]}', "customer": "",
            })

    # ------------------------------------------------------------ 4. control totals
    per_file = []
    for f in files:
        fname = f["sheets"][0]["file"] if f["sheets"] else "?"
        f_leads = f["leads"]
        acc = sum(1 for l in accepted if l["src_file"] == fname)
        per_file.append({
            "file": fname,
            "sheets": f["sheets"],
            "rows_read": f["rows_read"],
            "rows_blank": f["rows_blank"],
            "parsed": len(f_leads),
            "no_date": sum(1 for l in no_date if l["src_file"] == fname),
            "out_of_period": sum(1 for l in out_period if l["src_file"] == fname),
            "duplicates": sum(1 for l in dropped if l["src_file"] == fname),
            "accepted": acc,
        })

    totals = {
        "rows_read": sum(p["rows_read"] for p in per_file),
        "rows_blank": sum(p["rows_blank"] for p in per_file),
        "parsed": len(all_leads),
        "no_date": len(no_date),
        "out_of_period": len(out_period),
        "duplicates": len(dropped),
        "accepted": len(accepted),
    }
    totals["balanced"] = (
        totals["parsed"] - totals["no_date"] - totals["out_of_period"]
        - totals["duplicates"] == totals["accepted"]
    )

    sev_rank = {"critical": 0, "high": 1, "medium": 2, "low": 3}
    exceptions.sort(key=lambda e: (sev_rank.get(e["severity"], 9), e["kind"]))

    return {
        "period": {"start": p_start, "end": p_end},
        "leads": accepted,
        "excluded": {"no_date": no_date, "out_of_period": out_period, "duplicates": dropped},
        "control_totals": totals,
        "per_file": per_file,
        "exceptions": exceptions,
        "exception_summary": [
            {"kind": k, "count": c,
             "severity": next(e["severity"] for e in exceptions if e["kind"] == k)}
            for k, c in Counter(e["kind"] for e in exceptions).most_common()
        ],
        "reps": sorted({l["rep"] for l in accepted} - {unassigned}) + (
            [unassigned] if any(not l["assigned"] for l in accepted) else []),
        "sources": sorted({l["source"] for l in accepted}),
    }
