# Untreated Risk Aging Analytics — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add untreated-risk aging analytics (avg/median/oldest age, per-domain aging, SLA buckets) to stats, narrative, payload, dashboard, and PPT — age = today − Date created.

**Architecture:** Aging is computed purely in `zone_reports/stats.py` (relative to a `today` param), flows through `narrative.py` + `dashboard.build_payload` (today = generated_at date) into `index.html` (KPI + 2 charts) and `ppt_export.py` (compact text). One computation, identical numbers.

**Tech Stack:** Python 3.14 stdlib (`datetime`, `statistics`) + python-pptx; vanilla JS + Chart.js 4 + chartjs-plugin-datalabels.

## Global Constraints

- Untreated = `not is_treated(Stage)` (is_treated already = Stage lower == "monitoring").
- Age = `(today - created).days`; drop ages `< 0`; drop rows with unparseable `Date created`.
- `parse_date(s)`: strip, first 10 chars, `date.fromisoformat`, None on failure (handles `YYYY-MM-DD` and full ISO).
- `AGE_BUCKETS = ["0-30","31-90","91-180","180+"]`; bands 0–30, 31–90, 91–180, 180+.
- ZoneStats new fields (defaults; positional construction stays valid): `untreated=0, aged_count=0, avg_age=0, median_age=0, oldest_age=0, age_by_domain=[] (list of (domain,avg_age,count) desc by avg_age then domain), age_buckets=[] (list of (label,count) in AGE_BUCKETS order)`.
- `compute_zone_stats(rows, organization=None, today=None)`; `today` default `date.today()`; `build_payload` passes `today=generated_at.date()`.
- avg_age/median_age = rounded ints (0 when aged_count 0); oldest_age = max (0 when none).
- Narrative aging line (when untreated>0 and aged_count>0): `Untreated risks average {avg_age} days old (median {median_age}); oldest {oldest_age} days`.
- Dashboard bucket colors: 0-30 #34D399, 31-90 #E8C810, 91-180 #FFB74D, 180+ #F0736A. Per-domain bar warm ramp gold→orange→red by share of max.
- PPT: KPI strip appends "  ·  Avg untreated age {avg_age}d"; Key Insights box appends the aging line + top-3 age_by_domain as "{domain}: {avg} d avg ({count})". No new PPT chart.
- Tests no network/threads/sockets; existing 72 tests stay green.

---

### Task 1: stats.py — aging computation

**Files:** Modify `zone_reports/stats.py`; Test `tests/test_stats.py`

**Interfaces:** Produces `parse_date(s)->date|None`, `AGE_BUCKETS`, extended `compute_zone_stats(rows, organization=None, today=None)`, ZoneStats aging fields.

- [ ] **Step 1: Write failing tests**

```python
# append to tests/test_stats.py
import datetime as dt
from zone_reports.stats import parse_date, AGE_BUCKETS

def test_parse_date_variants():
    assert parse_date("2026-06-15") == dt.date(2026,6,15)
    assert parse_date("2026-06-15T08:30:00Z") == dt.date(2026,6,15)
    assert parse_date("") is None and parse_date("garbage") is None and parse_date(None) is None

def _aging_rows():
    # today = 2026-08-07; untreated unless Monitoring
    def r(i, org, dom, stage, created):
        return {"ID": f"R{i}", "Organization": org, "Category": dom, "Cat": "NCI",
                "Stage": stage, "Date created": created}
    return [
        r(1,"GHQ","Security","Assessment","2026-08-01"),     # age 6, untreated
        r(2,"GHQ","Security","Identification","2026-06-08"), # age 60, untreated
        r(3,"GHQ","Privacy","Remediation","2026-02-08"),     # age 180, untreated
        r(4,"GHQ","Privacy","Monitoring","2026-01-01"),      # TREATED -> excluded
        r(5,"GHQ","Security","Assessment",""),               # untreated, no date
        r(6,"GHQ","Privacy","Assessment","2027-01-01"),      # future -> dropped
    ]

def test_aging_metrics():
    today = dt.date(2026,8,7)
    s = compute_zone_stats(_aging_rows(), "GHQ", today=today)
    assert s.untreated == 5           # rows 1,2,3,5,6 (row4 treated)
    assert s.aged_count == 3          # rows 1,2,3 (5 no date, 6 future)
    assert s.avg_age == 82            # round((6+60+180)/3)=82
    assert s.median_age == 60
    assert s.oldest_age == 180

def test_age_by_domain_desc():
    today = dt.date(2026,8,7)
    s = compute_zone_stats(_aging_rows(), "GHQ", today=today)
    doms = {d: (a, c) for d, a, c in s.age_by_domain}
    assert doms["Privacy"][0] == 180 and doms["Privacy"][1] == 1
    assert doms["Security"][0] == 33   # round((6+60)/2)
    assert s.age_by_domain[0][0] == "Privacy"   # oldest first

def test_age_buckets():
    today = dt.date(2026,8,7)
    s = compute_zone_stats(_aging_rows(), "GHQ", today=today)
    b = dict(s.age_buckets)
    assert [x[0] for x in s.age_buckets] == AGE_BUCKETS
    assert b["0-30"] == 1 and b["31-90"] == 1 and b["91-180"] == 1 and b["180+"] == 0
```

- [ ] **Step 2: Run to verify failure**

Run: `python -m pytest tests/test_stats.py -k "parse_date or aging or age_" -v`
Expected: FAIL (`parse_date`/`AGE_BUCKETS` missing).

- [ ] **Step 3: Implement**

In `zone_reports/stats.py` add `import datetime` and `import statistics` at top (keep existing imports). Add:

```python
AGE_BUCKETS = ["0-30", "31-90", "91-180", "180+"]


def parse_date(s):
    s = (s or "").strip()
    if len(s) < 10:
        return None
    try:
        return datetime.date.fromisoformat(s[:10])
    except ValueError:
        return None


def _bucket(age):
    if age <= 30:
        return "0-30"
    if age <= 90:
        return "31-90"
    if age <= 180:
        return "91-180"
    return "180+"
```

Add the aging fields to `ZoneStats` (after `by_stage`):

```python
    untreated: int = 0
    aged_count: int = 0
    avg_age: int = 0
    median_age: int = 0
    oldest_age: int = 0
    age_by_domain: list = field(default_factory=list)
    age_buckets: list = field(default_factory=list)
```

Change the signature to `def compute_zone_stats(rows, organization=None, today=None):` and near the top of the body add `today = today or datetime.date.today()`. Before the return, after the treated/by_stage block, add:

```python
    untreated = 0
    ages, dom_ages, bucket_ct = [], {}, {b: 0 for b in AGE_BUCKETS}
    for r in sel:
        if is_treated(r.get("Stage")):
            continue
        untreated += 1
        d = parse_date(r.get("Date created"))
        if d is None:
            continue
        age = (today - d).days
        if age < 0:
            continue
        ages.append(age)
        dom = (r.get("Category") or "").strip() or "(blank)"
        dom_ages.setdefault(dom, []).append(age)
        bucket_ct[_bucket(age)] += 1
    aged_count = len(ages)
    avg_age = round(sum(ages) / aged_count) if aged_count else 0
    median_age = round(statistics.median(ages)) if aged_count else 0
    oldest_age = max(ages) if ages else 0
    age_by_domain = sorted(
        ((d, round(sum(a) / len(a)), len(a)) for d, a in dom_ages.items()),
        key=lambda t: (-t[1], t[0]))
    age_buckets = [(b, bucket_ct[b]) for b in AGE_BUCKETS]
```

Append the seven new fields to the `ZoneStats(...)` return, in the field order declared above.

- [ ] **Step 4: Run tests**

Run: `python -m pytest tests/test_stats.py -v`
Expected: PASS (existing MAZ/GRO/treated tests still green + new aging tests).

- [ ] **Step 5: Commit**

```bash
git add zone_reports/stats.py tests/test_stats.py
git commit -m "feat: untreated risk aging metrics in stats"
```

---

### Task 2: narrative aging line + payload fields (today from generated_at)

**Files:** Modify `zone_reports/narrative.py`, `dashboard.py`; Test `tests/test_narrative.py`, `tests/test_dashboard.py`

**Interfaces:** Consumes Task 1 ZoneStats. Produces narrative aging line; `zone_to_dict` adds the 7 aging keys; `build_payload` passes `today=generated_at.date()`.

- [ ] **Step 1: Write failing tests**

```python
# append to tests/test_narrative.py
def test_narrative_has_aging_line():
    import datetime as dt
    rows = [{"ID": f"R{i}", "Organization": "GHQ", "Category": "Security",
             "Cat": "NCI", "Stage": "Assessment",
             "Date created": "2026-06-08"} for i in range(3)]
    s = compute_zone_stats(rows, "GHQ", today=dt.date(2026,8,7))
    lines = build_narrative(s, "7th August 2026")
    assert any("Untreated risks average 60 days old (median 60); oldest 60 days" in ln
               for ln in lines)
```

```python
# append to tests/test_dashboard.py
def test_payload_has_aging_and_uses_generated_date():
    rows = [{"ID":"1","Organization":"GHQ","Category":"Security","Cat":"NCI",
             "Stage":"Assessment","Date created":"2026-06-08"}]
    p = dashboard.build_payload(rows, view="V",
                                generated_at=datetime.datetime(2026,8,7,6,0))
    z = p["zones"]["GHQ"]
    assert z["untreated"] == 1 and z["aged_count"] == 1
    assert z["avg_age"] == 60 and z["oldest_age"] == 60
    assert isinstance(z["age_by_domain"], list) and isinstance(z["age_buckets"], list)
```

- [ ] **Step 2: Run to verify failure**

Run: `python -m pytest tests/test_narrative.py tests/test_dashboard.py -k "aging" -v`
Expected: FAIL (no aging line; payload lacks aging / today defaults to real today).

- [ ] **Step 3: Implement**

`zone_reports/narrative.py` — in `build_narrative`, after the treated-line block, add:

```python
    if getattr(stats, "untreated", 0) and getattr(stats, "aged_count", 0):
        lines.append(f"Untreated risks average {stats.avg_age} days old "
                     f"(median {stats.median_age}); oldest {stats.oldest_age} days")
```

`dashboard.py` — in `build_payload`, change the per-zone call to pass the date:

```python
    for code, org in stats.ZONES.items():
        zs = stats.compute_zone_stats(rows, org, today=generated_at.date())
        ...
```

and in `zone_to_dict` add the seven keys:

```python
        "untreated": zs.untreated,
        "aged_count": zs.aged_count,
        "avg_age": zs.avg_age,
        "median_age": zs.median_age,
        "oldest_age": zs.oldest_age,
        "age_by_domain": zs.age_by_domain,
        "age_buckets": zs.age_buckets,
```

- [ ] **Step 4: Run tests**

Run: `python -m pytest tests/ -q`
Expected: PASS (whole suite).

- [ ] **Step 5: Commit**

```bash
git add zone_reports/narrative.py dashboard.py tests/test_narrative.py tests/test_dashboard.py
git commit -m "feat: aging narrative line + aging payload fields (today=generated_at)"
```

---

### Task 3: PPT aging text

**Files:** Modify `ppt_export.py`; Test `tests/test_ppt_export.py`

**Interfaces:** Consumes payload aging fields. Produces aging text in KPI strip + Key Insights box.

- [ ] **Step 1: Update fixture + write tests**

Add aging fields to each zone dict in `_payload()` in `tests/test_ppt_export.py` (e.g. `"untreated":2,"aged_count":2,"avg_age":90,"median_age":85,"oldest_age":150,"age_by_domain":[["Security",120,1],["Privacy",60,1]],"age_buckets":[["0-30",0],["31-90",1],["91-180",1],["180+",0]]`). Then add:

```python
def test_slide_has_aging_text():
    prs = Presentation(io.BytesIO(ppt_export.deck_to_bytes(_payload())))
    slide = prs.slides[1]
    texts = " ".join(sh.text_frame.text for sh in slide.shapes if sh.has_text_frame)
    assert "Avg untreated age 90d" in texts
    assert "days old" in texts or "d avg" in texts   # aging appears in insights
```

- [ ] **Step 2: Run to verify failure**

Run: `python -m pytest tests/test_ppt_export.py -k "aging" -v`
Expected: FAIL.

- [ ] **Step 3: Implement**

In `ppt_export.py` `build_deck`, where the KPI strip text is built, append when untreated:

```python
        if z.get("untreated"):
            kpi_text += f"  ·  Avg untreated age {z.get('avg_age', 0)}d"
```

(Use the actual KPI text variable name in the file.) Where the Key Insights box lines are assembled (the narrative list rendered into the text box), append after the narrative lines:

```python
        if z.get("untreated") and z.get("aged_count"):
            extra = [f"Untreated avg {z['avg_age']}d · median {z['median_age']}d · oldest {z['oldest_age']}d"]
            for d, avg, c in (z.get("age_by_domain") or [])[:3]:
                extra.append(f"{_label(d)}: {avg} d avg ({c})")
            # append these as additional paragraphs in the insights text box
```

Add these `extra` lines to the same text frame used for the narrative (respect the existing 16-line truncation — count `extra` toward it; if it would exceed, still show the overall aging line first).

- [ ] **Step 4: Run tests**

Run: `python -m pytest tests/test_ppt_export.py -v` then `python -m pytest tests/ -q`
Expected: PASS (aging text test + still 10 slides + bounds).

- [ ] **Step 5: Commit**

```bash
git add ppt_export.py tests/test_ppt_export.py
git commit -m "feat: PPT aging text in KPI strip + key insights"
```

---

### Task 4: index.html — aging KPI + 2 charts + insight line

**Files:** Modify `index.html`

**Interfaces:** Consumes payload aging fields. Produces leaf.

- [ ] **Step 1: Implement**

1. **KPI card "Avg Untreated Age"** in `#kpis` (after Risks Treated): big `z.avg_age + 'd'`; sublabel `median {z.median_age}d · oldest {z.oldest_age}d · {z.untreated} untreated`. When `z.untreated === 0`: big `0d`, sublabel `no untreated risks`.

2. **"Risk Aging by Domain" tile** with `<canvas id="agingChart">` inside a `.chart-inner-panel`: horizontal bar (`indexAxis:'y'`), labels = `z.age_by_domain` domains via `dispName` (oldest-first, already sorted), data = avg days, bar color per row = warm ramp by share of max age: helper `ageColor(v, maxv)` interpolating `#E8C810`→`#FFB74D`→`#F0736A` (v/maxv 0→0.5→1); datalabels = value + 'd' (cream/dark), x-axis title "days". Empty state div when `age_by_domain` empty.

3. **"Aging Buckets" tile** with `<canvas id="bucketChart">` inside a `.chart-inner-panel`: vertical bar of `z.age_buckets` counts, per-bar colors fixed `['#34D399','#E8C810','#FFB74D','#F0736A']` for the 4 bands in order; datalabels = counts; empty state when all counts 0.

4. Both charts: add module-scope `agingChart`/`bucketChart` vars; destroy-before-recreate; add `renderAging(z)` + `renderBuckets(z)` and call from `render()`; per-chart `plugins.datalabels`; tile `min-height` consistent with other chart tiles; glass `.tile` + inner `--panel` wrappers.

5. **Key Insights**: add an aging line under the treated line: `Untreated risks average <span class="ki-num">{avg_age}</span> days old (median {median_age}d; oldest {oldest_age}d)` — only when `z.untreated>0 && z.aged_count>0`.

Do not change any other JS logic (zone switch no fetch, other charts, exports, settings).

- [ ] **Step 2: Verify**

`node --check` extracted `<script>` (must pass). Headless (dummy refined.csv present):
```
python dashboard.py --no-browser --port 8081  (background)
curl -s http://127.0.0.1:8081/ | grep -c "agingChart"      # >=1
curl -s http://127.0.0.1:8081/ | grep -c "bucketChart"     # >=1
curl -s http://127.0.0.1:8081/ | grep -c "Avg Untreated Age"  # >=1
curl -s http://127.0.0.1:8081/api/data | python -c "import sys,json;d=json.load(sys.stdin);z=d['zones']['Overall'];print(z['avg_age'],z['untreated'],z['age_buckets'])"
stop server
```
`python -m pytest tests/ -q` green. Do NOT commit refined.csv/config.json/*.pptx.

- [ ] **Step 3: Commit**

```bash
git add index.html
git commit -m "feat: dashboard untreated-aging KPI + per-domain + buckets charts"
```

---

## Self-Review

**Spec coverage:** C1 stats aging → Task 1. C2 narrative → Task 2. C3 payload/today → Task 2. C4 dashboard KPI+2 charts+insight → Task 4. C5 PPT text → Task 3. Error cases (blank/future date, 0 untreated, no valid dates) → Task 1 logic + Task 4 empty states. All covered.

**Placeholder scan:** none — concrete code for Python; explicit ids/values for UI.

**Type consistency:** `compute_zone_stats(rows, organization=None, today=None)`, `parse_date`, `AGE_BUCKETS`, aging field names (`untreated, aged_count, avg_age, median_age, oldest_age, age_by_domain, age_buckets`) identical across Tasks 1-4; `age_by_domain` = list of (domain,avg,count); `age_buckets` = list of (label,count); bucket colors identical in Task 3/4; `build_payload` today=generated_at.date() single definition (Task 2). ✓
