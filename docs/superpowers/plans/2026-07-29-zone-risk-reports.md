# Zone Risk Reports Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** From `refined.csv`, build a multi-sheet Excel workbook with one live-PivotTable sheet per business zone plus an Overall sheet, each with native charts and an auto-generated narrative, cloned from the hand-built MAZ template.

**Architecture:** Two pure Python modules — `stats.py` (cross-tab computation) and `narrative.py` (placeholder-filled text) — have zero Excel dependency and are unit-tested against the known MAZ figures. A third module, `build_zone_reports.py`, drives Excel via COM (pywin32) to write the data sheet, clone MAZ per zone, set each pivot's Organization filter, and inject the regenerated title + narrative. A `--dry-run` flag exercises the pure path (prints stats/narrative for every zone) without launching Excel.

**Tech Stack:** Python 3, standard library (`csv`, `dataclasses`, `datetime`, `argparse`), `pywin32` (auto-installed) for Excel COM. Tests use `pytest`.

## Global Constraints

- Platform: Windows with Microsoft Excel installed (COM automation). Pure modules and `--dry-run` run anywhere.
- `pywin32` is auto-installed at runtime if `import win32com.client` fails; never assume it is preinstalled.
- Input file: `refined.csv` — the 19 columns produced by `refine_columns.py`, headers exactly: `Risk GUID, ID, Risk name, Source, Cat, Risk owners, Description, Treatment plan, Organization, Stage, Inherent risk score, Inherent risk level, Residual risk score, Residual risk level, Date created, Category, Result, Date closed, Risk approver`.
- Pivot definition (all sheets): Rows = `Category` field, Columns = `Cat` field, Values = Count of `ID`, Page filters = `Organization` (per zone) + `Stage` (all).
- Fixed Cat order: `COMMERCIAL, LOGISTICS, NCI, PACKAGING, TECHNOLOGY, RAU, Fees` (extras appended after, before Grand Total).
- Zone map (sheet → Organization): `GHQ→GHQ, AFR→Africa, SAZ→South America Zone, MAZ→Middle America Zone, NAZ→North America Zone, APAC→APAC, EUR→Europe, BEES→BEES, BEES-FT→BEES | FINTECH`; plus `Overall` (no Organization filter).
- Always write a `*_backup.xlsx` copy before modifying the workbook; always quit the Excel COM app in a `finally`.

---

### Task 1: `stats.py` — cross-tab computation

**Files:**
- Create: `zone_reports/__init__.py` (empty)
- Create: `zone_reports/stats.py`
- Test: `tests/test_stats.py`

**Interfaces:**
- Produces:
  - `CAT_ORDER: list[str]`
  - `ZONES: dict[str, str | None]` (sheet code → Organization value; `Overall` → `None`)
  - `@dataclass ZoneStats` with fields: `organization: str | None`, `total: int`, `by_cat: dict[str,int]`, `grand_by_cat: list[tuple[str,int]]`, `by_domain: list[tuple[str,int]]`, `crosstab: dict[str, dict[str,int]]`, `domain_pct: dict[str,int]`, `top_cats: dict[str, list[tuple[str,int]]]`
  - `load_rows(path: str) -> list[dict]`
  - `compute_zone_stats(rows: list[dict], organization: str | None = None) -> ZoneStats`
  - `select_top_cats(pairs_desc: list[tuple[str,int]], total: int, threshold: float = 0.8, min_items: int = 2, max_items: int = 4) -> list[tuple[str,int]]`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_stats.py
import csv
from zone_reports.stats import (
    compute_zone_stats, select_top_cats, CAT_ORDER, ZONES,
)

# The MAZ cross-tab from the reference workbook (rows=domain, cols=Cat).
MAZ = {
    "Business Continuity Management": {"COMMERCIAL":14,"LOGISTICS":1,"NCI":9,"TECHNOLOGY":2,"RAU":5,"Fees":1},
    "Information Security":           {"COMMERCIAL":4,"LOGISTICS":4,"PACKAGING":3},
    "Privacy":                        {"COMMERCIAL":2,"LOGISTICS":2,"NCI":3},
    "Security":                       {"COMMERCIAL":7,"LOGISTICS":1,"NCI":6,"PACKAGING":2},
}

def _maz_rows():
    rows, n = [], 0
    for domain, cats in MAZ.items():
        for cat, count in cats.items():
            for _ in range(count):
                n += 1
                rows.append({"ID": f"R{n}", "Organization": "Middle America Zone",
                             "Category": domain, "Cat": cat, "Stage": "Open"})
    # add noise from another zone that must be filtered out
    rows.append({"ID":"X1","Organization":"Africa","Category":"Privacy","Cat":"NCI","Stage":"Open"})
    return rows

def test_total_and_grand_by_cat():
    s = compute_zone_stats(_maz_rows(), "Middle America Zone")
    assert s.total == 66
    assert s.by_cat == {"COMMERCIAL":27,"LOGISTICS":8,"NCI":18,"PACKAGING":5,
                        "TECHNOLOGY":2,"RAU":5,"Fees":1}

def test_domain_percentages():
    s = compute_zone_stats(_maz_rows(), "Middle America Zone")
    assert s.domain_pct["Business Continuity Management"] == 48
    assert s.domain_pct["Security"] == 24
    assert s.domain_pct["Information Security"] == 17
    assert s.domain_pct["Privacy"] == 11
    # domains sorted descending by count
    assert [d for d, _ in s.by_domain][0] == "Business Continuity Management"

def test_top_cats_match_reference_narrative():
    s = compute_zone_stats(_maz_rows(), "Middle America Zone")
    assert s.top_cats["Business Continuity Management"] == [("COMMERCIAL",14),("NCI",9),("RAU",5)]
    assert s.top_cats["Security"] == [("COMMERCIAL",7),("NCI",6)]
    assert s.top_cats["Information Security"] == [("COMMERCIAL",4),("LOGISTICS",4),("PACKAGING",3)]
    assert s.top_cats["Privacy"] == [("NCI",3),("COMMERCIAL",2),("LOGISTICS",2)]

def test_overall_includes_all_orgs():
    s = compute_zone_stats(_maz_rows(), None)
    assert s.total == 67  # 66 MAZ + 1 Africa noise row

def test_select_top_cats_rule():
    # accumulate desc until cumulative >= 80% of total, min 2, max 4
    assert select_top_cats([("A",7),("B",6),("C",2),("D",1)], 16) == [("A",7),("B",6)]
    assert select_top_cats([("A",14),("B",9),("C",5),("D",2)], 32) == [("A",14),("B",9),("C",5)]
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_stats.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'zone_reports'`

- [ ] **Step 3: Write minimal implementation**

```python
# zone_reports/__init__.py  (empty file)
```

```python
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
    "BEES":    "BEES",
    "BEES-FT": "BEES | FINTECH",
    "Overall": None,
}


@dataclass
class ZoneStats:
    organization: str | None
    total: int
    by_cat: dict                      # Cat -> count, in CAT_ORDER (+extras)
    grand_by_cat: list                # [(Cat, count)] same order, for charts
    by_domain: list                   # [(domain, count)] desc
    crosstab: dict                    # domain -> {Cat: count}
    domain_pct: dict                  # domain -> int percent of total
    top_cats: dict                    # domain -> [(Cat, count)] selected


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


def compute_zone_stats(rows, organization=None):
    sel = [r for r in rows
           if organization is None or (r.get("Organization") or "").strip() == organization]
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

    return ZoneStats(organization, total, by_cat, grand_by_cat,
                     by_domain, crosstab, domain_pct, top_cats)
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_stats.py -v`
Expected: PASS (5 tests)

- [ ] **Step 5: Commit**

```bash
git add zone_reports/__init__.py zone_reports/stats.py tests/test_stats.py
git commit -m "feat: zone cross-tab stats module with MAZ-verified tests"
```

---

### Task 2: `narrative.py` — placeholder-filled report text

**Files:**
- Create: `zone_reports/narrative.py`
- Test: `tests/test_narrative.py`

**Interfaces:**
- Consumes: `ZoneStats` from `zone_reports.stats`.
- Produces:
  - `format_date_ordinal(d: datetime.date) -> str` (e.g. `"29th July 2026"`)
  - `build_title(zone_name: str, date_str: str) -> str`
  - `build_narrative(stats: ZoneStats, date_str: str) -> list[str]` (text lines for the narrative block)

- [ ] **Step 1: Write the failing test**

```python
# tests/test_narrative.py
from datetime import date
from zone_reports.stats import compute_zone_stats
from zone_reports.narrative import build_narrative, build_title, format_date_ordinal
from tests.test_stats import _maz_rows

def test_format_date_ordinal():
    assert format_date_ordinal(date(2026, 5, 14)) == "14th May 2026"
    assert format_date_ordinal(date(2026, 5, 1)) == "1st May 2026"
    assert format_date_ordinal(date(2026, 5, 22)) == "22nd May 2026"
    assert format_date_ordinal(date(2026, 5, 3)) == "3rd May 2026"

def test_title():
    assert build_title("Middle America Zone", "14th May 2026") == \
        "Middle America Zone Risk Analysis (Till 14th May 2026)"

def test_narrative_key_lines():
    s = compute_zone_stats(_maz_rows(), "Middle America Zone")
    text = "\n".join(build_narrative(s, "14th May 2026"))
    assert "A total of 66 risks identified across business units" in text
    assert "Commercial (27 risks) and NCI (18 risks) together account for 68% of total exposure" in text
    assert "Business Continuity Management (48% of total risk" in text
    assert "Security (24% of total risk)" in text
    # top cats appear with title-cased names and counts
    assert "Commercial (14)" in text
    assert "NCI (9)" in text
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_narrative.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'zone_reports.narrative'`

- [ ] **Step 3: Write minimal implementation**

```python
# zone_reports/narrative.py
"""Placeholder-filled narrative text for a zone report, from ZoneStats."""

# Cat/domain names are Title-cased for display; acronyms preserved.
_ACRONYMS = {"NCI", "RAU", "GHQ", "BEES"}

# Connector phrase per domain rank (clamped to the last for deeper ranks).
_PHRASES = ["heavily concentrated in", "primarily in", "led by", "distributed across"]


def _disp(name):
    return name if name.upper() in _ACRONYMS else name.title()


def format_date_ordinal(d):
    suffix = "th" if 11 <= d.day % 100 <= 13 else {1: "st", 2: "nd", 3: "rd"}.get(d.day % 10, "th")
    return f"{d.day}{suffix} {d.strftime('%B %Y')}"


def build_title(zone_name, date_str):
    return f"{zone_name} Risk Analysis (Till {date_str})"


def build_narrative(stats, date_str):
    lines = []
    total = stats.total
    lines.append(f"A total of {total} risks identified across business units")
    if total == 0:
        return lines

    lines.append("Risk concentration is uneven, with strong clustering in a few key areas")
    top2 = stats.grand_by_cat and sorted(stats.grand_by_cat, key=lambda kv: -kv[1])[:2]
    if len(top2) == 2:
        (c1, n1), (c2, n2) = top2
        pct = round(100 * (n1 + n2) / total)
        lines.append(f"{_disp(c1)} ({n1} risks) and {_disp(c2)} ({n2} risks) "
                     f"together account for {pct}% of total exposure")

    lines.append("")
    lines.append("Key Insights by Risk Domain")
    for rank, (dom, dcount) in enumerate(stats.by_domain):
        superl = " – highest" if rank == 0 else ""
        lines.append(f"{_disp(dom)} ({stats.domain_pct[dom]}% of total risk{superl})")
        phrase = _PHRASES[min(rank, len(_PHRASES) - 1)]
        lines.append(f"{dcount} risks, {phrase}:")
        for cat, n in stats.top_cats[dom]:
            lines.append(f"    {_disp(cat)} ({n})")
        lines.append("")
    return lines
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_narrative.py -v`
Expected: PASS (3 tests)

- [ ] **Step 5: Commit**

```bash
git add zone_reports/narrative.py tests/test_narrative.py
git commit -m "feat: placeholder narrative generator, verified against MAZ text"
```

---

### Task 3: `build_zone_reports.py` — CLI + `--dry-run`

**Files:**
- Create: `build_zone_reports.py`
- Test: `tests/test_dryrun.py`

**Interfaces:**
- Consumes: `load_rows`, `compute_zone_stats`, `ZONES` from `zone_reports.stats`; `build_title`, `build_narrative`, `format_date_ordinal` from `zone_reports.narrative`.
- Produces:
  - `iter_zone_reports(rows, date_str) -> list[dict]` where each dict is `{"sheet": str, "org": str|None, "zone_name": str, "stats": ZoneStats, "title": str, "narrative": list[str]}`
  - `zone_display_name(sheet: str, org: str | None) -> str` (Overall → "Overall"; else the Organization value)
  - `main(argv=None)` with `--data`, `--workbook`, `--out`, `--template-sheet`, `--dry-run`.

- [ ] **Step 1: Write the failing test**

```python
# tests/test_dryrun.py
from build_zone_reports import iter_zone_reports, zone_display_name
from tests.test_stats import _maz_rows

def test_zone_display_name():
    assert zone_display_name("MAZ", "Middle America Zone") == "Middle America Zone"
    assert zone_display_name("Overall", None) == "Overall"

def test_iter_zone_reports_covers_all_zones():
    reports = iter_zone_reports(_maz_rows(), "14th May 2026")
    sheets = {r["sheet"] for r in reports}
    assert {"GHQ","AFR","SAZ","MAZ","NAZ","APAC","EUR","BEES","BEES-FT","Overall"} <= sheets
    maz = next(r for r in reports if r["sheet"] == "MAZ")
    assert maz["stats"].total == 66
    assert maz["title"] == "Middle America Zone Risk Analysis (Till 14th May 2026)"
    overall = next(r for r in reports if r["sheet"] == "Overall")
    assert overall["stats"].total == 67
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_dryrun.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'build_zone_reports'`

- [ ] **Step 3: Write minimal implementation** (pure path only; COM added in Task 4)

```python
#!/usr/bin/env python3
"""Build per-zone + Overall Excel risk reports from refined.csv.

Pure computation and --dry-run need no Excel. A full run drives Excel via COM
(pywin32, auto-installed) to clone the MAZ template per zone. See
docs/superpowers/specs/2026-07-29-zone-risk-reports-design.md.
"""
import argparse
import sys
from datetime import date

from zone_reports.stats import load_rows, compute_zone_stats, ZONES
from zone_reports.narrative import build_title, build_narrative, format_date_ordinal


def zone_display_name(sheet, org):
    return "Overall" if org is None else org


def iter_zone_reports(rows, date_str):
    reports = []
    for sheet, org in ZONES.items():
        stats = compute_zone_stats(rows, org)
        name = zone_display_name(sheet, org)
        reports.append({
            "sheet": sheet, "org": org, "zone_name": name, "stats": stats,
            "title": build_title(name, date_str),
            "narrative": build_narrative(stats, date_str),
        })
    return reports


def _print_dry_run(reports):
    for r in reports:
        print("=" * 70)
        print(r["title"])
        print(f"  sheet={r['sheet']}  org={r['org']}  total={r['stats'].total}")
        print("  by Cat:", r["stats"].by_cat)
        print("  " + "\n  ".join(r["narrative"]))


def main(argv=None):
    ap = argparse.ArgumentParser(description="Build zone risk report sheets.")
    ap.add_argument("--data", default="refined.csv")
    ap.add_argument("--workbook", help="Workbook holding the MAZ template (required unless --dry-run).")
    ap.add_argument("--out", help="Output path (default: overwrite --workbook).")
    ap.add_argument("--template-sheet", default="MAZ")
    ap.add_argument("--dry-run", action="store_true",
                    help="Compute + print stats/narrative for every zone; no Excel.")
    args = ap.parse_args(argv)

    rows = load_rows(args.data)
    date_str = format_date_ordinal(date.today())
    reports = iter_zone_reports(rows, date_str)

    if args.dry_run:
        _print_dry_run(reports)
        return 0

    if not args.workbook:
        sys.exit("--workbook is required for a full run (or use --dry-run).")
    from zone_reports.excel import render_workbook       # Task 4
    render_workbook(args.workbook, args.out or args.workbook,
                    args.template_sheet, args.data, reports)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_dryrun.py -v`
Expected: PASS (2 tests). Also spot-check: `python build_zone_reports.py --data refined.csv --dry-run` (needs a refined.csv; skip if absent).

- [ ] **Step 5: Commit**

```bash
git add build_zone_reports.py tests/test_dryrun.py
git commit -m "feat: zone report CLI with pure --dry-run path"
```

---

### Task 4: `zone_reports/excel.py` — Excel COM rendering

**Files:**
- Create: `zone_reports/excel.py`

**Interfaces:**
- Consumes: report dicts from `iter_zone_reports` (Task 3); `CAT_ORDER` from stats.
- Produces: `render_workbook(workbook_path, out_path, template_sheet, data_csv, reports) -> None`
- Helpers (module-private): `_ensure_pywin32()`, `_write_data_sheet(wb, data_csv) -> table_range`, `_find_pivot(ws)`, `_configure_pivot(pt, org)`, `_find_title_cell(ws)`, `_write_narrative(ws, anchor, title, lines)`.

**Note:** This task's deliverable is validated interactively on this Windows machine against a **copy** of the real workbook (the code always writes a `*_backup.xlsx` first). There is no automated unit test for the COM path; the pure logic it depends on is already covered by Tasks 1–3.

- [ ] **Step 1: Implement pywin32 bootstrap + data sheet writer**

```python
# zone_reports/excel.py
"""Excel COM (pywin32) rendering for zone reports. Windows + Excel required."""
import os
import shutil
import subprocess
import sys

from zone_reports.stats import CAT_ORDER

DATA_SHEET = "Data"
TABLE_NAME = "RiskData"


def _ensure_pywin32():
    try:
        import win32com.client  # noqa: F401
        return
    except ImportError:
        print("pywin32 not found -> installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pywin32"])
        import win32com.client  # noqa: F401


def _write_data_sheet(wb, data_csv):
    """Load refined.csv into a fresh Data sheet as a ListObject Table; return the
    table's range address (e.g. 'Data!$A$1:$S$4491')."""
    import csv
    with open(data_csv, newline="", encoding="utf-8-sig") as f:
        rows = list(csv.reader(f))
    if not rows:
        raise SystemExit("refined.csv is empty.")

    # remove an existing Data sheet, then recreate
    for ws in list(wb.Worksheets):
        if ws.Name == DATA_SHEET:
            ws.Delete()
    ws = wb.Worksheets.Add()
    ws.Name = DATA_SHEET

    nrows, ncols = len(rows), len(rows[0])
    # bulk write via a 2D array assignment (fast)
    top_left = ws.Cells(1, 1)
    bottom_right = ws.Cells(nrows, ncols)
    ws.Range(top_left, bottom_right).Value = rows

    # define the Table (ListObject) over the used range
    xlSrcRange = 1
    rng = ws.Range(top_left, bottom_right)
    lo = ws.ListObjects.Add(xlSrcRange, rng, None, 1)  # 1 = xlYes (has headers)
    lo.Name = TABLE_NAME
    return f"{DATA_SHEET}!{rng.Address}"
```

- [ ] **Step 2: Implement pivot discovery + configuration**

```python
# append to zone_reports/excel.py

def _find_pivot(ws):
    """Return the first PivotTable on a sheet, or None."""
    try:
        if ws.PivotTables().Count >= 1:
            return ws.PivotTables(1)
    except Exception:
        pass
    return None


def _repoint_pivot_source(wb, pt, table_range):
    """Point a pivot's cache at the RiskData table range and refresh."""
    xlDatabase = 1
    cache = wb.PivotCaches().Create(SourceType=xlDatabase, SourceData=table_range)
    pt.ChangePivotCache(cache)
    pt.RefreshTable()


def _configure_pivot(pt, org):
    """Set the Organization page-filter (org=None -> All), Stage -> All, and
    enforce the fixed Cat column order. Field names come from the template
    pivot, not hard-coded captions."""
    # Organization page filter
    try:
        f = pt.PivotFields("Organization")
        f.ClearAllFilters()
        if org is None:
            f.CurrentPage = "(All)"
        else:
            f.CurrentPage = org
    except Exception as e:
        print(f"  WARN: Organization filter: {e}")

    # Stage -> all
    try:
        s = pt.PivotFields("Stage")
        s.ClearAllFilters()
        s.CurrentPage = "(All)"
    except Exception:
        pass

    # Fixed Cat order
    try:
        cf = pt.PivotFields("Cat")
        pos = 1
        for cat in CAT_ORDER:
            try:
                cf.PivotItems(cat).Position = pos
                pos += 1
            except Exception:
                continue  # cat not present in this zone
    except Exception as e:
        print(f"  WARN: Cat order: {e}")

    pt.RefreshTable()
```

- [ ] **Step 2b: Run a smoke check that the module imports**

Run: `python -c "import zone_reports.excel"`
Expected: no error (pywin32 not imported at module load; only inside `_ensure_pywin32`).

- [ ] **Step 3: Implement title/narrative discovery + writer**

```python
# append to zone_reports/excel.py

def _find_title_cell(ws):
    """Find the cell whose text contains 'Risk Analysis' (the MAZ title)."""
    try:
        found = ws.UsedRange.Find("Risk Analysis")
        if found is not None:
            return found.Row, found.Column
    except Exception:
        pass
    return None


def _write_title_and_narrative(ws, title_rc, title, lines):
    """Overwrite the title cell and the narrative block below it."""
    if title_rc is None:
        print("  WARN: no title cell found; skipping title/narrative for this sheet.")
        return
    r, c = title_rc
    ws.Cells(r, c).Value = title
    # narrative goes into the first column of the block, a couple rows below
    start = r + 1
    for i, line in enumerate(lines):
        ws.Cells(start + i, c).Value = line
```

- [ ] **Step 4: Implement the orchestrator `render_workbook`**

```python
# append to zone_reports/excel.py

def render_workbook(workbook_path, out_path, template_sheet, data_csv, reports):
    _ensure_pywin32()
    import win32com.client as win32

    backup = os.path.splitext(workbook_path)[0] + "_backup.xlsx"
    shutil.copyfile(workbook_path, backup)
    print(f"Backup written: {backup}")

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(os.path.abspath(workbook_path))

        if template_sheet not in [ws.Name for ws in wb.Worksheets]:
            raise SystemExit(f"Template sheet '{template_sheet}' not found in workbook.")

        table_range = _write_data_sheet(wb, data_csv)

        tmpl = wb.Worksheets(template_sheet)

        for rep in reports:
            sheet = rep["sheet"]
            # start each zone from a clean clone of the template
            if sheet != template_sheet and sheet in [ws.Name for ws in wb.Worksheets]:
                wb.Worksheets(sheet).Delete()
            if sheet == template_sheet:
                ws = tmpl
            else:
                tmpl.Copy(After=wb.Worksheets(wb.Worksheets.Count))
                ws = wb.Worksheets(wb.Worksheets.Count)
                ws.Name = sheet

            pt = _find_pivot(ws)
            if pt is not None:
                _repoint_pivot_source(wb, pt, table_range)
                _configure_pivot(pt, rep["org"])
            else:
                print(f"  WARN: no pivot on sheet {sheet}.")

            _write_title_and_narrative(ws, _find_title_cell(ws),
                                       rep["title"], rep["narrative"])
            print(f"  built {sheet}: total={rep['stats'].total}")

        wb.RefreshAll()
        excel.CalculateUntilAsyncQueriesDone()
        wb.SaveAs(os.path.abspath(out_path))
        print(f"Saved: {out_path}")
    finally:
        excel.Quit()
```

- [ ] **Step 5: Validate on a copy of the real workbook**

Prereqs: `refined.csv` present; the real workbook copied to `report_test.xlsx` in the project folder.

Run: `python build_zone_reports.py --data refined.csv --workbook report_test.xlsx --out report_out.xlsx`
Expected: prints backup path, "built GHQ/AFR/.../Overall", "Saved: report_out.xlsx"; no orphaned `EXCEL.EXE` in Task Manager. Open `report_out.xlsx` and confirm each zone sheet has the pivot filtered to its Organization, charts populated, and the regenerated title/narrative.

- [ ] **Step 6: Commit**

```bash
git add zone_reports/excel.py
git commit -m "feat: Excel COM rendering — data sheet, per-zone pivot clone, narrative"
```

---

## Self-Review

**Spec coverage:**
- Per-zone sheets + Overall → Task 3 `ZONES` (includes `Overall`), Task 4 clone loop. ✓
- Live PivotTable via COM → Task 4 `_repoint_pivot_source` / `_configure_pivot`. ✓
- Fixed Cat order → `CAT_ORDER`, enforced in `_configure_pivot`. ✓
- Data as `Data` sheet Table `RiskData` → `_write_data_sheet`. ✓
- Narrative regenerated per zone, placeholder-based → Task 2. ✓
- Title "Till <today>" ordinal → `format_date_ordinal`, `build_title`. ✓
- pywin32 auto-install → `_ensure_pywin32`. ✓
- Backup before touching + always Quit → `render_workbook`. ✓
- Stage = all → `_configure_pivot` sets Stage "(All)". ✓
- Zero-row zone still builds → `iter_zone_reports` runs for every ZONES entry; narrative handles total==0. ✓

**Placeholder scan:** No TBD/TODO; all code steps contain full code. `_configure_pivot`/`_find_title_cell` use `except` fallbacks with explicit warnings rather than vague "handle errors". ✓

**Type consistency:** `iter_zone_reports` returns dicts with keys `sheet/org/zone_name/stats/title/narrative`; `render_workbook` reads exactly those (`rep["sheet"]`, `rep["org"]`, `rep["title"]`, `rep["narrative"]`, `rep["stats"]`). `ZoneStats` fields used in narrative (`grand_by_cat`, `by_domain`, `domain_pct`, `top_cats`, `total`) all defined in Task 1. ✓

**Note on Overall filter:** `CurrentPage = "(All)"` requires the Organization field to allow all items; `ClearAllFilters` first ensures no leftover filter. If the installed Excel rejects `"(All)"`, fallback is to set `.EnableMultiplePageItems` — flagged here so the implementer checks during Task 4 Step 5 validation.
