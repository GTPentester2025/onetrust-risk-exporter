# GRO Merge, Risk-Treatment, Insight Parity & Glass Polish — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Merge BEES + BEES-FT into a GRO zone, add treated-risk analytics (Stage=Monitoring) as a KPI + donut on dashboard and PPT, bring the dashboard Key Insights to full PPT parity, and restyle the UI to professional glassmorphism.

**Architecture:** All numbers originate in `zone_reports/stats.py` (GRO membership, treated/by_stage), flow through `narrative.py` (treated line) and `dashboard.build_payload` into both `index.html` and `ppt_export.py`. UI work is split into a functional task (GRO tabs, treated tiles, insight parity) and a visual glass-polish task.

**Tech Stack:** Python 3.14 stdlib + python-pptx; vanilla JS + Chart.js 4 + chartjs-plugin-datalabels + html2canvas; CSS `backdrop-filter` glassmorphism.

## Global Constraints

- `ZONES` order (drives tab + slide order): GHQ, AFR, SAZ, MAZ, NAZ, APAC, EUR, GRO, Overall. `GRO` value = `["BEES", "BEES | FINTECH"]`. No BEES/BEES-FT entries remain.
- `compute_zone_stats(rows, organization)` accepts `str | list | None` (None=all, list=membership, str=equality).
- Treated: `TREATED_STAGES = {"monitoring"}`; `is_treated(stage)` = `(stage or "").strip().lower() in TREATED_STAGES`.
- `ZoneStats` gains `treated:int`, `treated_pct:int`, `by_stage:list[(stage,count)]` (desc, then name); blank stage → "(blank)".
- Narrative treated line (only when total>0), placed after the concentration line: `"{treated} of {total} risks ({treated_pct}%) are actively monitored (treated)"`.
- Payload zone dict gains `treated`, `treated_pct`, `by_stage`; zone-set == {GHQ,AFR,SAZ,MAZ,NAZ,APAC,EUR,GRO,Overall}.
- Treated donut: Treated `#34D399`, Open `#E8C810` (open = total-treated). PPT slide count == title + 9 == 10.
- Glass: tiles `rgba(255,255,255,0.05)` + `backdrop-filter: blur(14px) saturate(1.2)` + 1px `rgba(255,255,255,0.09)` border + radius 16px; charts/table on inner panel `rgba(20,16,10,0.55)`; text cream `#f3f2f1`/muted `#c8c6c4` ≥4.5:1; gold `#e8c810` accent only; backdrop-filter only on top-level tiles.
- Legacy `build_zone_reports.py` must not crash on a list-valued ZONES entry (skip it).
- Server 127.0.0.1; existing tests stay green except those asserting BEES/BEES-FT or slide-count 11 (update them); no network/threads/sockets in tests.

---

### Task 1: stats.py — GRO membership + treated analytics

**Files:**
- Modify: `zone_reports/stats.py`, `build_zone_reports.py`
- Test: `tests/test_stats.py`

**Interfaces:**
- Produces: `ZONES` (GRO), `TREATED_STAGES`, `is_treated(stage)->bool`, `compute_zone_stats(rows, organization=None)` accepting `str|list|None`; `ZoneStats` with new fields `treated`, `treated_pct`, `by_stage`.

- [ ] **Step 1: Write failing tests**

```python
# append to tests/test_stats.py
from zone_reports.stats import is_treated

def _rows_mixed():
    rows = []
    n = 0
    def add(org, dom, cat, stage, k=1):
        nonlocal n
        for _ in range(k):
            n += 1
            rows.append({"ID": f"R{n}", "Organization": org, "Category": dom,
                         "Cat": cat, "Stage": stage})
    add("BEES", "Security", "NCI", "Monitoring", 3)
    add("BEES | FINTECH", "Privacy", "COMMERCIAL", "Assessment", 2)
    add("BEES | FINTECH", "Security", "NCI", "Monitoring", 1)
    add("GHQ", "Security", "NCI", "Closed", 4)
    return rows

def test_gro_combines_both_orgs():
    s = compute_zone_stats(_rows_mixed(), ["BEES", "BEES | FINTECH"])
    assert s.total == 6                       # 3 + 2 + 1, excludes GHQ's 4
    assert s.treated == 4                      # 3 + 1 monitoring
    assert s.treated_pct == 67                 # round(100*4/6)

def test_compute_accepts_str_and_none():
    rows = _rows_mixed()
    assert compute_zone_stats(rows, "BEES").total == 3
    assert compute_zone_stats(rows, None).total == 10

def test_by_stage_desc():
    s = compute_zone_stats(_rows_mixed(), None)
    stages = dict(s.by_stage)
    assert stages["Monitoring"] == 4 and stages["Closed"] == 4 and stages["Assessment"] == 2
    assert s.by_stage[0][1] >= s.by_stage[-1][1]     # sorted desc

def test_is_treated():
    assert is_treated("Monitoring") and is_treated(" monitoring ")
    assert not is_treated("Closed") and not is_treated("") and not is_treated(None)

def test_zones_has_gro_not_bees():
    assert "GRO" in ZONES and "BEES" not in ZONES and "BEES-FT" not in ZONES
    assert ZONES["GRO"] == ["BEES", "BEES | FINTECH"]
    assert list(ZONES)[-1] == "Overall"
```

Also update `test_overall_includes_all_orgs` if it references removed zones (it uses org=None — unaffected; leave). No BEES/BEES-FT assertions exist in test_stats currently besides ZONES import.

- [ ] **Step 2: Run to verify failure**

Run: `python -m pytest tests/test_stats.py -v`
Expected: FAIL (`is_treated` import error; ZONES still has BEES).

- [ ] **Step 3: Implement**

In `zone_reports/stats.py`:

Replace the `ZONES` dict:

```python
ZONES = {
    "GHQ":  "GHQ",
    "AFR":  "Africa",
    "SAZ":  "South America Zone",
    "MAZ":  "Middle America Zone",
    "NAZ":  "North America Zone",
    "APAC": "APAC",
    "EUR":  "Europe",
    "GRO":  ["BEES", "BEES | FINTECH"],
    "Overall": None,
}

TREATED_STAGES = {"monitoring"}


def is_treated(stage):
    return (stage or "").strip().lower() in TREATED_STAGES
```

Add fields to `ZoneStats` (with defaults so positional construction stays safe):

```python
@dataclass
class ZoneStats:
    organization: object
    total: int
    by_cat: dict
    grand_by_cat: list
    by_domain: list
    crosstab: dict
    domain_pct: dict
    top_cats: dict
    treated: int = 0
    treated_pct: int = 0
    by_stage: list = field(default_factory=list)
```

Add an org matcher and extend `compute_zone_stats`:

```python
def _match_org(org_val, organization):
    if organization is None:
        return True
    o = (org_val or "").strip()
    if isinstance(organization, (list, tuple, set)):
        return o in organization
    return o == organization
```

Change the `sel` line to `sel = [r for r in rows if _match_org(r.get("Organization"), organization)]`, then before the return add:

```python
    treated, stage_total = 0, {}
    for r in sel:
        stg = (r.get("Stage") or "").strip() or "(blank)"
        stage_total[stg] = stage_total.get(stg, 0) + 1
        if is_treated(r.get("Stage")):
            treated += 1
    treated_pct = _pct(treated, total)
    by_stage = sorted(stage_total.items(), key=lambda kv: (-kv[1], kv[0]))
```

and return `ZoneStats(organization, total, by_cat, grand_by_cat, by_domain, crosstab, domain_pct, top_cats, treated, treated_pct, by_stage)`.

In `build_zone_reports.py`, `iter_zone_reports` (line ~21) is PURE and works with GRO unchanged — do NOT skip it. But `zone_display_name` returns the raw org, which is a list for GRO; fix it to fall back to the sheet code for non-string orgs:

```python
def zone_display_name(sheet, org):
    if org is None:
        return "Overall"
    if not isinstance(org, str):     # list-valued (e.g. GRO)
        return sheet
    return org
```

(The Excel COM path in `zone_reports/excel.py` — `render_workbook`, only reached on a real non-dry-run Excel build — cannot set a multi-org pivot page filter for GRO. That is out of scope per the spec; it is not exercised by tests or the web app. Leave it; do not attempt to support GRO in the Excel pivot here.)

Also update `tests/test_dryrun.py` line 11: replace the expected sheet set to the GRO shape:

```python
    assert {"GHQ","AFR","SAZ","MAZ","NAZ","APAC","EUR","GRO","Overall"} <= sheets
```

and add an assertion that GRO resolves to a clean display name:

```python
    assert zone_display_name("GRO", ["BEES", "BEES | FINTECH"]) == "GRO"
```

- [ ] **Step 4: Run tests**

Run: `python -m pytest tests/test_stats.py -v`
Expected: PASS (existing MAZ tests still pass; new tests pass).

- [ ] **Step 5: Commit**

```bash
git add zone_reports/stats.py build_zone_reports.py tests/test_stats.py tests/test_dryrun.py
git commit -m "feat: GRO zone merge + treated/by_stage analytics in stats"
```

---

### Task 2: narrative treated line + payload fields

**Files:**
- Modify: `zone_reports/narrative.py`, `dashboard.py`
- Test: `tests/test_narrative.py`, `tests/test_dashboard.py`

**Interfaces:**
- Consumes: Task 1 `ZoneStats`.
- Produces: narrative includes the treated line; `zone_to_dict` adds `treated`, `treated_pct`, `by_stage` to each zone dict.

- [ ] **Step 1: Write failing tests**

```python
# append to tests/test_narrative.py  (import compute_zone_stats + build_narrative already used there)
import datetime as _dt
from zone_reports.stats import compute_zone_stats
from zone_reports.narrative import build_narrative

def test_narrative_has_treated_line():
    rows = [{"ID": f"R{i}", "Organization": "GHQ", "Category": "Security",
             "Cat": "NCI", "Stage": "Monitoring" if i < 3 else "Assessment"}
            for i in range(5)]
    s = compute_zone_stats(rows, "GHQ")
    lines = build_narrative(s, "7th August 2026")
    joined = "\n".join(lines)
    assert "3 of 5 risks (60%) are actively monitored (treated)" in joined

def test_narrative_no_treated_line_when_empty():
    s = compute_zone_stats([], "GHQ")
    lines = build_narrative(s, "7th August 2026")
    assert not any("monitored (treated)" in ln for ln in lines)
```

```python
# append to tests/test_dashboard.py
def test_payload_zone_set_has_gro():
    rows = [{"ID": "1", "Organization": "BEES", "Category": "Security",
             "Cat": "NCI", "Stage": "Monitoring"}]
    p = dashboard.build_payload(rows, view="V",
                                generated_at=datetime.datetime(2026, 8, 7, 6, 0))
    assert set(p["zones"]) == {"GHQ","AFR","SAZ","MAZ","NAZ","APAC","EUR","GRO","Overall"}
    gro = p["zones"]["GRO"]
    assert gro["total"] == 1 and gro["treated"] == 1 and gro["treated_pct"] == 100
    assert isinstance(gro["by_stage"], list)
```

- [ ] **Step 2: Run to verify failure**

Run: `python -m pytest tests/test_narrative.py tests/test_dashboard.py -k "treated or gro" -v`
Expected: FAIL (no treated line; payload lacks treated / has BEES not GRO).

- [ ] **Step 3: Implement**

`zone_reports/narrative.py` — in `build_narrative`, after the concentration `if len(top2) == 2:` block and before `lines.append("")`, add:

```python
    if getattr(stats, "treated", 0) is not None and total:
        lines.append(f"{stats.treated} of {total} risks "
                     f"({stats.treated_pct}%) are actively monitored (treated)")
```

`dashboard.py` — in `zone_to_dict`, add the three keys:

```python
def zone_to_dict(zs, narrative_lines):
    return {
        "organization": zs.organization,
        "total": zs.total,
        "by_cat": zs.by_cat,
        "grand_by_cat": zs.grand_by_cat,
        "by_domain": zs.by_domain,
        "crosstab": zs.crosstab,
        "domain_pct": zs.domain_pct,
        "top_cats": zs.top_cats,
        "treated": zs.treated,
        "treated_pct": zs.treated_pct,
        "by_stage": zs.by_stage,
        "narrative": narrative_lines,
    }
```

- [ ] **Step 4: Run tests**

Run: `python -m pytest tests/ -q`
Expected: PASS (whole suite; note any test asserting zone-set with BEES must be updated — search `BEES` in tests and fix to GRO).

- [ ] **Step 5: Commit**

```bash
git add zone_reports/narrative.py dashboard.py tests/test_narrative.py tests/test_dashboard.py
git commit -m "feat: treated narrative line + treated/by_stage in payload"
```

---

### Task 3: PPT — treated donut + 3-chart row layout

**Files:**
- Modify: `ppt_export.py`
- Test: `tests/test_ppt_export.py`

**Interfaces:**
- Consumes: payload zone dict with `treated`, `total`, plus existing fields.
- Produces: same public API; each zone slide has pie + stacked bar + treated donut (3 chart shapes) within slide bounds; slide count title+9.

- [ ] **Step 1: Update + write tests**

Update the fixture `_payload()` in `tests/test_ppt_export.py` to the 9-zone set (replace the `BEES`/`BEES-FT` codes in its zone-code list with `GRO`; keep 9 codes incl. Overall) and add `"treated": 1, "treated_pct": 33, "by_stage": [["Monitoring",1],["Assessment",2]]` to each zone dict it builds. Update `test_deck_has_title_plus_one_slide_per_zone` to assert `len(prs.slides) == 10`. Then add:

```python
def test_zone_slide_has_three_charts():
    prs = Presentation(io.BytesIO(ppt_export.deck_to_bytes(_payload())))
    slide = prs.slides[1]
    charts = [sh for sh in slide.shapes if sh.has_chart]
    assert len(charts) >= 3      # pie, stacked bar, treated donut
    assert any(sh.has_table for sh in slide.shapes)
```

Keep `test_zone_slide_labeled_and_fits` (bounds check) — it now also covers the donut.

- [ ] **Step 2: Run to verify failure**

Run: `python -m pytest tests/test_ppt_export.py -v`
Expected: FAIL (slide count 10 vs old; only 2 charts).

- [ ] **Step 3: Implement**

In `ppt_export.py` `build_deck`, per-zone slide: keep header band + KPI strip (append `"  ·  Treated {z['treated']} ({z['treated_pct']}%)"` to the KPI text). Change the charts to a 3-across row (all `Inches`):

- Pie (risks by domain): position `(0.4, 1.6)`, size `(4.3, 3.0)`.
- Stacked bar: position `(4.9, 1.6)`, size `(4.6, 3.0)`.
- Treated donut (native pie with 2 points): position `(9.7, 1.6)`, size `(3.2, 3.0)`:

```python
        treated = z.get("treated", 0)
        openn = max(0, z["total"] - treated)
        if z["total"]:
            cd = CategoryChartData()
            cd.categories = ["Treated", "Open"]
            cd.add_series("Risks", (treated, openn))
            gr = slide.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, Inches(9.7), Inches(1.6),
                                        Inches(3.2), Inches(3.0), cd)
            ch = gr.chart
            ch.has_title = True; ch.chart_title.text_frame.text = "Risk Treatment"
            pts = ch.plots[0].series[0].points
            pts[0].format.fill.solid(); pts[0].format.fill.fore_color.rgb = RGBColor(0x34,0xD3,0x99)
            pts[1].format.fill.solid(); pts[1].format.fill.fore_color.rgb = RGBColor(0xE8,0xC8,0x10)
            ch.has_legend = True; ch.legend.position = XL_LEGEND_POSITION.BOTTOM
            ch.legend.include_in_layout = False
            ch.font.color.rgb = THEME["cream"]; ch.font.size = Pt(9)
            plot = ch.plots[0]; plot.has_data_labels = True
            plot.data_labels.number_format = '0'; plot.data_labels.font.size = Pt(9)
            plot.data_labels.font.color.rgb = THEME["cream"]
```

Move the table + Key Insights to the bottom row: table at `(0.4, 4.9)` size `(7.6, 2.3)`; insights at `(8.3, 4.9)` size `(4.6, 2.3)`. Ensure `XL_CHART_TYPE.DOUGHNUT` is imported (from `pptx.enum.chart`).

- [ ] **Step 4: Run tests**

Run: `python -m pytest tests/test_ppt_export.py -v` then `python -m pytest tests/ -q`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add ppt_export.py tests/test_ppt_export.py
git commit -m "feat: PPT treated donut + 3-chart zone slide layout"
```

---

### Task 4: UI — GRO tabs, Risks-Treated KPI + donut, insight parity

**Files:**
- Modify: `index.html`

**Interfaces:**
- Consumes: payload with 9 zones incl. GRO; each zone `treated`, `treated_pct`, `by_stage`, `narrative`, and existing fields.
- Produces: leaf.

- [ ] **Step 1: Implement**

1. **Zone tabs:** change the hardcoded zone code list to
   `['GHQ','AFR','SAZ','MAZ','NAZ','APAC','EUR','GRO','Overall']` (remove BEES/BEES-FT). Grep the file for `'BEES'`/`"BEES-FT"` and remove all remaining references.

2. **Risks Treated KPI card:** add a 4th KPI card in `#kpis` (after top-supplier-category): big number `z.treated` (tabular, `toLocaleString`), sublabel `of {z.total} · {z.treated_pct}%`, green accent border/number (`--live #34d399`). Zero-total → "0 / 0 · 0%".

3. **Risk Treatment tile + donut:** add a new tile (in the pie/bar row or its own row) titled "Risk Treatment" with `<canvas id="treatedChart">`. Render a Chart.js **doughnut**: labels ["Treated","Open"], data `[z.treated, z.total - z.treated]`, colors `['#34D399','#E8C810']`, `cutout:'62%'`, legend bottom, datalabels showing `value` + `(pct%)` in dark bold, tooltip "label: n (P%)". Destroy previous instance before recreate (add `treatedChart` to the destroy logic). Empty state text when `z.total === 0`. Add `renderTreated(z)` and call it from `render()`.

4. **Insight parity (renderNarrative):** add a `dispName(s)` helper — keep `{'NCI','RAU','GHQ','BEES','GRO'}` uppercase, else Title Case each word — and apply it to domain names (`.ki-domain-name`) and cat chip labels so the dashboard matches the PPT's `_disp`. Add a treated line under the concentration line: `"<span class='ki-num'>{treated}</span> of {total} ({treated_pct}%) actively monitored (treated)"`. Ensure every domain block shows ALL `top_cats` chips with counts (already does). Keep copy/download exporting `z.narrative`.

- [ ] **Step 2: Verify**

`node --check` the extracted `<script>` (must pass). Headless (dummy `refined.csv` already present):
```
python dashboard.py --no-browser --port 8071  (background)
curl -s http://127.0.0.1:8071/ | grep -c "treatedChart"        # >=1
curl -s http://127.0.0.1:8071/ | grep -c "Risks Treated"       # >=1
curl -s http://127.0.0.1:8071/ | grep -oE "'GRO'|\"GRO\"" | head -1   # GRO present
curl -s http://127.0.0.1:8071/ | grep -c "BEES"                # 0
curl -s http://127.0.0.1:8071/api/data | python -c "import sys,json;d=json.load(sys.stdin);print('GRO' in d['zones'], d['zones']['GRO']['treated'])"
stop server
```
`python -m pytest tests/ -q` green.

- [ ] **Step 3: Commit**

```bash
git add index.html
git commit -m "feat: GRO tabs, Risks-Treated KPI + donut, insight parity"
```

---

### Task 5: UI — glassmorphism professional polish

**Files:**
- Modify: `index.html`

**Interfaces:**
- Consumes: the Task 4 DOM. Produces: leaf (visual only; no behavior change).

- [ ] **Step 1: Implement (CSS-first; do not change JS behavior)**

1. **Ambient base:** body background = `linear-gradient(160deg,#14100a,#0a0806)`; add a fixed full-viewport `::before` (or a `#bg-blobs` div) with 2–3 radial-gradient blobs (gold `rgba(232,200,16,0.09)`, teal `rgba(52,211,153,0.06)`) blurred, `pointer-events:none`, `z-index:0`; content above at `z-index:1`. No pure black.

2. **Glass tokens + surfaces:** add `--glass: rgba(255,255,255,0.05); --glass-border: rgba(255,255,255,0.09); --glass-blur: blur(14px) saturate(1.2); --panel: rgba(20,16,10,0.55);`. Apply to `.tile`, KPI cards, the header/app bar, and the settings modal card: `background:var(--glass); backdrop-filter:var(--glass-blur); -webkit-backdrop-filter:var(--glass-blur); border:1px solid var(--glass-border); border-radius:16px; box-shadow:0 8px 24px rgba(0,0,0,.45), inset 0 1px 0 rgba(255,255,255,.06);`. Apply `backdrop-filter` ONLY at this top level (not nested elements).

3. **Legibility:** wrap each chart `<canvas>` and the table in an inner container with `background:var(--panel); border-radius:12px; padding:8px;` so data reads crisp over blur. KPI numbers get a faint scrim (`text-shadow:0 1px 2px rgba(0,0,0,.5)` or a subtle inner panel). Confirm body text uses `--fg #f3f2f1` / `--fg-muted #c8c6c4` (≥4.5:1); reserve gold for numbers/active state only (audit labels currently gold that should be cream).

4. **Chrome polish:** zone tab strip — active tab gets a single sliding gold underline (`::after` bar or a moved indicator), clear `:hover` and `:focus-visible` (2px gold ring) states; all numeric outputs use `tabular-nums` and `toLocaleString` thousands separators (KPIs, table cells, chart tooltips/labels where applicable); enforce an 8px spacing rhythm on tile padding/gaps; give chart tiles a `min-height` so switching zones doesn't reflow; polish empty/error/loading states to match the glass theme.

5. **Perf + motion:** keep transitions 150–300ms; `@media (prefers-reduced-motion: reduce){ *{animation:none!important;transition:none!important} }`; blobs are static.

- [ ] **Step 2: Verify**

`node --check` extracted script (unchanged JS still valid). Headless:
```
python dashboard.py --no-browser --port 8072  (background)
curl -s http://127.0.0.1:8072/ | grep -c "backdrop-filter"     # >=1
curl -s http://127.0.0.1:8072/ | grep -c -- "--glass"          # >=1
curl -s http://127.0.0.1:8072/ | head -3                        # HTML
stop server
```
`python -m pytest tests/ -q` still green (no Python touched). Visual check performed manually by the human against the dummy data (glass tiles, ambient bg, contrast, GRO tab, treated donut).

- [ ] **Step 3: Commit**

```bash
git add index.html
git commit -m "feat: glassmorphism professional polish (ambient bg, glass tiles, legibility, chrome)"
```

---

## Self-Review

**Spec coverage:** C1 GRO → Task 1 (+Task 4 tabs). C2 treated stats → Task 1; narrative/payload → Task 2; dashboard tiles → Task 4; PPT → Task 3. C3 insight parity → Task 4. C4 treated tiles → Task 4 (dash) + Task 3 (ppt). C5 glass → Task 5. Legacy Excel guard → Task 1. Test updates (BEES/slide-11) → Tasks 1-3. All covered.

**Placeholder scan:** none — concrete code for Python tasks; explicit ids/values/requirements for UI tasks.

**Type consistency:** `compute_zone_stats(rows, organization)` str|list|None consistent; `ZoneStats` new fields (`treated`,`treated_pct`,`by_stage`) flow through `zone_to_dict` → payload → UI/PPT with identical names; `is_treated`/`TREATED_STAGES` single definition (Task 1); donut colors `#34D399`/`#E8C810` identical in Task 3 (PPT) and Task 4 (UI); zone-code list identical in Task 4 and stats `ZONES`. ✓
