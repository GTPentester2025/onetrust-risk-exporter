# zone_reports/stats.py
"""Pure cross-tab computation for zone risk reports. No Excel dependency."""
import csv
from dataclasses import dataclass, field

CAT_ORDER = ["COMMERCIAL", "LOGISTICS", "NCI", "PACKAGING", "TECHNOLOGY", "RAU", "Fees"]

# sheet code -> Organization value; None means "all organizations"
ZONES = {
    "GHQ":     "GHQ",
    "AFR":     "Africa",
    "SAZ":     "South America Zone",
    "MAZ":     "Middle America Zone",
    "NAZ":     "North America Zone",
    "APAC":    "APAC",
    "EUR":     "Europe",
    "GRO":     ["BEES", "BEES | FINTECH"],
    "Overall": None,
}

TREATED_STAGES = {"monitoring"}


def is_treated(stage):
    return (stage or "").strip().lower() in TREATED_STAGES


@dataclass
class ZoneStats:
    organization: object
    total: int
    by_cat: dict                      # Cat -> count, in CAT_ORDER (+extras)
    grand_by_cat: list                # [(Cat, count)] same order, for charts
    by_domain: list                   # [(domain, count)] desc
    crosstab: dict                    # domain -> {Cat: count}
    domain_pct: dict                  # domain -> int percent of total
    top_cats: dict                    # domain -> [(Cat, count)] selected
    treated: int = 0
    treated_pct: int = 0
    by_stage: list = field(default_factory=list)


def load_rows(path):
    for enc in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            with open(path, newline="", encoding=enc) as f:
                return list(csv.DictReader(f))
        except UnicodeDecodeError:
            continue
    raise SystemExit(f"Could not decode {path} as CSV.")


def _pct(n, total):
    return round(100 * n / total) if total else 0


def select_top_cats(pairs_desc, total, threshold=0.8, min_items=2, max_items=4):
    """Take cats (already sorted desc) until cumulative share >= threshold,
    clamped to [min_items, max_items]. Drops zero-count cats."""
    pairs = [(c, n) for c, n in pairs_desc if n > 0]
    if not pairs:
        return []
    out, cum = [], 0
    for cat, n in pairs:
        out.append((cat, n))
        cum += n
        if len(out) >= min_items and total and cum / total >= threshold:
            break
        if len(out) >= max_items:
            break
    return out


def _cat_sort_key(cat):
    return (CAT_ORDER.index(cat) if cat in CAT_ORDER else len(CAT_ORDER), cat)


def _match_org(org_val, organization):
    if organization is None:
        return True
    o = (org_val or "").strip()
    if isinstance(organization, (list, tuple, set)):
        return o in organization
    return o == organization


def compute_zone_stats(rows, organization=None):
    sel = [r for r in rows if _match_org(r.get("Organization"), organization)]
    total = len(sel)

    crosstab, dom_total, cat_total = {}, {}, {}
    for r in sel:
        dom = (r.get("Category") or "").strip() or "(blank)"
        cat = (r.get("Cat") or "").strip() or "(blank)"
        crosstab.setdefault(dom, {}).setdefault(cat, 0)
        crosstab[dom][cat] += 1
        dom_total[dom] = dom_total.get(dom, 0) + 1
        cat_total[cat] = cat_total.get(cat, 0) + 1

    cats = sorted(cat_total, key=_cat_sort_key)
    by_cat = {c: cat_total[c] for c in cats}
    grand_by_cat = [(c, cat_total[c]) for c in cats]

    by_domain = sorted(dom_total.items(), key=lambda kv: (-kv[1], kv[0]))
    domain_pct = {d: _pct(n, total) for d, n in by_domain}

    top_cats = {}
    for dom, _ in by_domain:
        pairs = sorted(crosstab[dom].items(), key=lambda kv: (-kv[1], _cat_sort_key(kv[0])))
        top_cats[dom] = select_top_cats(pairs, dom_total[dom])

    treated, stage_total = 0, {}
    for r in sel:
        stg = (r.get("Stage") or "").strip() or "(blank)"
        stage_total[stg] = stage_total.get(stg, 0) + 1
        if is_treated(r.get("Stage")):
            treated += 1
    treated_pct = _pct(treated, total)
    by_stage = sorted(stage_total.items(), key=lambda kv: (-kv[1], kv[0]))

    return ZoneStats(organization, total, by_cat, grand_by_cat,
                     by_domain, crosstab, domain_pct, top_cats,
                     treated, treated_pct, by_stage)
