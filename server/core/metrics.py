"""
Analytics. Takes reconciled leads and produces every figure the UI and the
Excel export need, as one plain-JSON structure.

Conventions used throughout:
  touch rate        = (leads - New) / leads
  qualification rate= (Qualified + Quoted/Proposal + Won) / leads
  conversion %      = Won / leads
  loss rate         = (Lost + Not Qualified) / leads
"""
from __future__ import annotations

import datetime as dt
from collections import Counter, defaultdict


def pct(a: float, b: float) -> float:
    return round(100.0 * a / b, 2) if b else 0.0


def _block(rows: list[dict], cfg: dict) -> dict:
    C = cfg["canonical"]
    stages = C["stages"]
    c = Counter(r["stage"] for r in rows)
    total = len(rows)
    new = c[C["untouched_stage"]]
    won = c[C["won_stage"]]
    closed_no = sum(c[s] for s in C["closed_no_order"])
    qplus = c["Qualified"] + c["Quoted / Proposal"] + won
    worked = total - new
    return {
        **{s: c[s] for s in stages},
        "total": total,
        "worked": worked,
        "qplus": qplus,
        "quoted_plus": c["Quoted / Proposal"] + won,
        "closed_no_order": closed_no,
        "open": total - won - closed_no,
        "qty": sum(r["qty"] for r in rows),
        "hot": sum(1 for r in rows if r["quality"] == "Hot Lead"),
        "unassigned": sum(1 for r in rows if not r["assigned"]),
        "touch_rate": pct(worked, total),
        "qual_rate": pct(qplus, total),
        "quoted_rate": pct(c["Quoted / Proposal"] + won, total),
        "conversion": pct(won, total),
        "win_on_worked": pct(won, worked),
        "loss_rate": pct(closed_no, total),
    }


def compute(rec: dict, cfg: dict) -> dict:
    C = cfg["canonical"]
    leads = rec["leads"]
    stages, quals = C["stages"], C["qualities"]
    reps, sources = rec["reps"], rec["sources"]
    end = dt.date.fromisoformat(rec["period"]["end"])
    asof = end + dt.timedelta(days=1)

    by = lambda **kw: [l for l in leads
                       if all((l["rep"] if k == "rep" else l[k]) == v for k, v in kw.items())]

    overall = _block(leads, cfg)
    by_rep = {r: _block(by(rep=r), cfg) for r in reps}
    by_source = {s: _block(by(source=s), cfg) for s in sources}
    by_detail = {d: _block([l for l in leads if l["source_detail"] == d], cfg)
                 for d in sorted({l["source_detail"] for l in leads})}

    cube = {r: {"ALL": by_rep[r],
                **{s: _block([l for l in leads if l["rep"] == r and l["source"] == s], cfg)
                   for s in sources}} for r in reps}
    tot_cube = {"ALL": overall, **{s: by_source[s] for s in sources}}

    # ---- stage x source matrix (the "stage wise snapshot")
    stage_matrix = [{
        "stage": s,
        "total": sum(1 for l in leads if l["stage"] == s),
        "share": pct(sum(1 for l in leads if l["stage"] == s), len(leads)),
        **{src: sum(1 for l in leads if l["stage"] == s and l["source"] == src) for src in sources},
    } for s in stages]

    quality_matrix = [{
        "quality": q,
        "total": sum(1 for l in leads if l["quality"] == q),
        "share": pct(sum(1 for l in leads if l["quality"] == q), len(leads)),
        "qualified": sum(1 for l in leads if l["quality"] == q and l["stage"] == "Qualified"),
        "not_qualified": sum(1 for l in leads if l["quality"] == q and l["stage"] == "Not Qualified"),
        **{src: sum(1 for l in leads if l["quality"] == q and l["source"] == src) for src in sources},
    } for q in quals]

    # ---- untouched ageing
    new_leads = [l for l in leads if l["stage"] == C["untouched_stage"]]
    age_of = lambda l: (asof - dt.date.fromisoformat(l["date"])).days
    ageing = dict(sorted(Counter(age_of(l) for l in new_leads).items()))
    ageing_by_rep = {r: dict(sorted(Counter(age_of(l) for l in new_leads if l["rep"] == r).items()))
                     for r in reps}
    aged = sum(1 for l in new_leads if age_of(l) >= 2)

    # ---- unassigned register
    unassigned = sorted(
        [{"qty": l["qty"], "age": age_of(l), "date": l["date"], "source": l["source"],
          "customer": l["customer"], "phone": l["phone"], "city": l["city"],
          "product": l["product"]} for l in leads if not l["assigned"]],
        key=lambda x: (-x["qty"], x["date"]))

    # ---- dispositions and demand lost
    closed = [l for l in leads if l["stage"] in C["closed_no_order"]]
    disp = defaultdict(lambda: {"leads": 0, "qty": 0, "severity": "normal"})
    for l in closed:
        d = disp[l["disposition"]["label"]]
        d["leads"] += 1
        d["qty"] += l["qty"]
        d["severity"] = l["disposition"]["severity"]
    dispositions = sorted(
        [{"label": k, **v, "share": pct(v["leads"], len(closed))} for k, v in disp.items()],
        key=lambda x: -x["leads"])
    demand = sorted(
        [{"label": k, "leads": v["leads"], "qty": v["qty"], "severity": v["severity"]}
         for k, v in disp.items() if v["qty"] > 0], key=lambda x: -x["qty"])

    demand_totals = {
        "total": sum(l["qty"] for l in leads),
        "closed": sum(l["qty"] for l in closed),
        "open": sum(l["qty"] for l in leads
                    if l["stage"] not in C["closed_no_order"] + [C["won_stage"]]),
        "won": sum(l["qty"] for l in leads if l["stage"] == C["won_stage"]),
        "product_gap": sum(l["qty"] for l in closed if l["disposition"]["severity"] == "product"),
    }

    # ---- leads worth chasing again
    recoverable = sorted(
        [{"qty": l["qty"], "date": l["date"], "source": l["source"], "rep": l["rep"],
          "customer": l["customer"], "product": l["product"], "quality": l["quality"],
          "stage": l["stage"], "note": l["note"], "why": l["disposition"]["label"]}
         for l in closed
         if l["qty"] >= 10 or l["quality"] in ("Hot Lead", "Warm Lead")
         or l["disposition"]["severity"] in ("recoverable", "product")],
        key=lambda x: -x["qty"])[:40]

    # ---- daily
    daily = []
    d0 = dt.date.fromisoformat(rec["period"]["start"])
    while d0 <= end:
        k = d0.isoformat()
        rows = [l for l in leads if l["date"] == k]
        daily.append({"date": k, "day": d0.strftime("%a"), "total": len(rows),
                      **{s: sum(1 for l in rows if l["stage"] == s) for s in stages},
                      **{f"src_{s}": sum(1 for l in rows if l["source"] == s) for s in sources}})
        d0 += dt.timedelta(days=1)

    return {
        "period": rec["period"], "asof": asof.isoformat(),
        "stages": stages, "qualities": quals, "reps": reps, "sources": sources,
        "overall": overall, "by_rep": by_rep, "by_source": by_source, "by_detail": by_detail,
        "cube": cube, "tot_cube": tot_cube,
        "stage_matrix": stage_matrix, "quality_matrix": quality_matrix,
        "ageing": {str(k): v for k, v in ageing.items()},
        "ageing_by_rep": {r: {str(k): v for k, v in d.items()} for r, d in ageing_by_rep.items()},
        "aged_untouched": aged,
        "unassigned": unassigned,
        "dispositions": dispositions, "demand": demand, "demand_totals": demand_totals,
        "recoverable": recoverable, "daily": daily,
    }


def build_insights(m: dict, rec: dict, cfg: dict) -> list[dict]:
    """Findings written from the numbers, ranked by what they cost the business."""
    C = cfg["canonical"]
    out, o = [], m["overall"]
    n = lambda x: f"{x:,}"

    if m["unassigned"]:
        q = sum(u["qty"] for u in m["unassigned"])
        top = m["unassigned"][0]
        out.append({"sev": "critical", "title":
            f'{len(m["unassigned"])} leads have no salesperson assigned at all',
            "why": f'They carry {n(q)} pieces of demand'
                   + (f', including a single {n(top["qty"])}-piece enquiry from {top["customer"]}'
                      if top["qty"] > 0 else '')
                   + f'. Oldest is {max(u["age"] for u in m["unassigned"])} days old.',
            "action": "Assign every one of them today, largest quantity first."})

    dt_ = m["demand_totals"]
    if dt_["total"] and dt_["closed"]:
        out.append({"sev": "critical", "title":
            f'{n(dt_["closed"])} of {n(dt_["total"])} pieces of demand were closed without an order',
            "why": f'{pct(dt_["closed"], dt_["total"])}% of everything asked for this week. '
                   f'{n(dt_["open"])} pieces are still open and winnable.',
            "action": "Work the open pieces before they age out."})
    if dt_["product_gap"]:
        out.append({"sev": "high", "title":
            f'{n(dt_["product_gap"])} pieces were rejected on product range, not on buyer intent',
            "why": "These buyers wanted something the range does not cover. "
                   "That is a sourcing decision sitting inside a sales report.",
            "action": "Put the range gap in front of management with the piece count attached."})

    if o["New"]:
        out.append({"sev": "high", "title":
            f'{o["New"]} of {o["total"]} leads ({pct(o["New"], o["total"])}%) were never worked',
            "why": f'{m["aged_untouched"]} of them are already 2+ days old. '
                   f'Online lead intent decays within hours.',
            "action": "Clear the aged backlog before taking on new spend."})

    worst = sorted([(r, b) for r, b in m["by_rep"].items()
                    if r != C["unassigned_label"] and b["total"] >= 10],
                   key=lambda x: x[1]["touch_rate"])
    if worst and worst[0][1]["touch_rate"] < 70:
        r, b = worst[0]
        out.append({"sev": "high", "title":
            f'{r} worked only {b["touch_rate"]}% of {b["total"]} assigned leads',
            "why": f'{b["New"]} leads never touched. '
                   f'Best in team is {max(x[1]["touch_rate"] for x in worst)}%.',
            "action": "Redistribute the aged untouched leads and review capacity."})
        # channel-specific check - is it the person or the channel?
        split = [(s, m["cube"][r][s]) for s in m["sources"] if m["cube"][r][s]["total"] >= 8]
        good = [s for s, b2 in split if b2["qual_rate"] >= 20]
        bad = [s for s, b2 in split if b2["qual_rate"] == 0]
        if good and bad:
            out.append({"sev": "medium", "title":
                f'{r}’s gap is channel-specific, not a blanket skill problem',
                "why": f'Qualifies well on {", ".join(good)} but produces nothing '
                       f'from {", ".join(bad)}.',
                "action": f'Weight their allocation toward {good[0]} and coach the rest.'})

    ranked = sorted([(s, b) for s, b in m["by_source"].items() if b["total"] >= 10],
                    key=lambda x: -x[1]["qual_rate"])
    if len(ranked) >= 2:
        best, low = ranked[0], ranked[-1]
        out.append({"sev": "high", "title":
            f'{best[0]} qualifies {best[1]["qual_rate"]}% of its leads; '
            f'{low[0]} manages {low[1]["qual_rate"]}%',
            "why": f'{best[0]}: {best[1]["total"]} leads, {n(best[1]["qty"])} pcs of demand. '
                   f'{low[0]}: {low[1]["total"]} leads, {n(low[1]["qty"])} pcs.',
            "action": f'Shift budget and attention from {low[0]} toward {best[0]}.'})

    overload = sorted(((r, b) for r, b in m["by_rep"].items() if r != C["unassigned_label"]),
                      key=lambda x: -x[1]["total"])
    if len(overload) >= 2 and overload[0][1]["total"] > 1.4 * overload[-1][1]["total"]:
        r, b = overload[0]
        out.append({"sev": "medium", "title":
            f'{r} carries {b["total"]} leads — {pct(b["total"], o["total"])}% of everything',
            "why": f'Loss rate {b["loss_rate"]}%. The lightest load in the team is '
                   f'{overload[-1][1]["total"]}.',
            "action": "Rebalance the allocation before quality of follow-up slips further."})

    gaps = {e["kind"]: e["count"] for e in rec["exception_summary"]}
    if gaps.get("no_reason"):
        out.append({"sev": "medium", "title":
            f'{gaps["no_reason"]} leads were closed with no reason recorded',
            "why": "Those losses cannot be diagnosed or coached against.",
            "action": "Make Loss Reason mandatory before a lead can be closed."})
    if gaps.get("recoverable") or gaps.get("hot_lost"):
        k = gaps.get("recoverable", 0) + gaps.get("hot_lost", 0)
        out.append({"sev": "high", "title":
            f'{k} leads were written off despite recorded interest',
            "why": "Either rated Hot/Warm or carrying a note that records a live requirement.",
            "action": "Re-open these first — cheapest pipeline available."})
    if not any(l["qty"] for l in rec["leads"] if l["source"] == "Meta"):
        out.append({"sev": "medium", "title": "No deal value is captured in any feed",
            "why": "Revenue conversion, average order value and ROI per source cannot be "
                   "calculated. Only lead counts and piece quantities are measurable.",
            "action": "Add an Order Value field at Qualified stage."})

    rank = {"critical": 0, "high": 1, "medium": 2}
    return sorted(out, key=lambda x: rank.get(x["sev"], 3))
