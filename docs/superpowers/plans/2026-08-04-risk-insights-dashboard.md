# Risk Insights Dashboard Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A one-command local dashboard that fetches OneTrust risk data via the existing pipeline, caches it to `refined.csv`, and renders KPI/pie/stacked-bar/narrative/table insights with an instant zone filter.

**Architecture:** A stdlib `http.server` (`dashboard.py`) serves a single `index.html` and a small JSON API. `/api/fetch` shells out to the existing `onetrust_view_export.py` + `refine_columns.py`, then imports `zone_reports.stats` + `zone_reports.narrative` to build one all-zones JSON payload. The browser caches that payload and switches zones client-side with no refetch.

**Tech Stack:** Python 3.14 stdlib (`http.server`, `subprocess`, `json`, `webbrowser`, `datetime`); reused modules `zone_reports/stats.py`, `zone_reports/narrative.py`; vanilla HTML/CSS/JS; Chart.js via CDN. No new pip dependencies.

## Global Constraints

- No new pip dependencies — stdlib + existing modules only.
- Server binds `127.0.0.1` only (never `0.0.0.0`).
- Client secret is never written to disk and never logged; passed to child processes as args only.
- Reuse `zone_reports.stats` and `zone_reports.narrative` unchanged — do not modify existing scripts.
- Zone set and organization mapping come from `stats.ZONES` (includes `Overall: None`); Cat order from `stats.CAT_ORDER`.
- Cache artifact is `refined.csv`; raw intermediate is `raw_export.csv`.
- Tests use `pytest`, live in `tests/`, and require no network, no Excel, no socket.
- Default port `8000`, overridable with `--port`.

---

### Task 1: All-zones payload builder (pure function)

**Files:**
- Create: `dashboard.py`
- Test: `tests/test_dashboard.py`

**Interfaces:**
- Consumes: `zone_reports.stats.compute_zone_stats`, `stats.ZONES`, `stats.CAT_ORDER`, `stats.load_rows`; `zone_reports.narrative.build_narrative`, `narrative.format_date_ordinal`.
- Produces: `build_payload(rows, view=None, generated_at=None) -> dict` and `zone_to_dict(zs, narrative_lines) -> dict`. Payload shape: `{source?, generated_at, view, cat_order, zones: {code: zone_dict}}` where `zone_dict = {organization, total, by_cat, grand_by_cat, by_domain, crosstab, domain_pct, top_cats, narrative}`. Later tasks (server) call `build_payload`.

- [ ] **Step 1: Write the failing test**

```python
# tests/test_dashboard.py
import datetime
from dashboard import build_payload

def _rows():
    rows = []
    n = 0
    for cat in ("COMMERCIAL", "NCI", "COMMERCIAL"):
        n += 1
        rows.append({"ID": f"R{n}", "Organization": "GHQ",
                     "Category": "Security", "Cat": cat, "Stage": "Open"})
    rows.append({"ID": "A1", "Organization": "Africa",
                 "Category": "Privacy", "Cat": "NCI", "Stage": "Open"})
    return rows

def test_payload_shape_and_zone_totals():
    p = build_payload(_rows(), view="V",
                      generated_at=datetime.datetime(2026, 8, 4, 12, 0, 0))
    assert p["view"] == "V"
    assert p["generated_at"] == "2026-08-04T12:00:00"
    assert p["cat_order"] == ["COMMERCIAL", "LOGISTICS", "NCI",
                              "PACKAGING", "TECHNOLOGY", "RAU", "Fees"]
    # every zone code present, incl. Overall
    assert set(p["zones"]) == {"GHQ", "AFR", "SAZ", "MAZ", "NAZ",
                               "APAC", "EUR", "BEES", "BEES-FT", "Overall"}
    ghq = p["zones"]["GHQ"]
    assert ghq["total"] == 3
    assert ghq["by_cat"] == {"COMMERCIAL": 2, "NCI": 1}
    assert p["zones"]["Overall"]["total"] == 4      # includes Africa row
    assert p["zones"]["AFR"]["total"] == 1
    assert ghq["narrative"][0] == "A total of 3 risks identified across business units"

def test_zero_row_zone_is_present():
    p = build_payload(_rows(), view="V",
                      generated_at=datetime.datetime(2026, 8, 4, 12, 0, 0))
    naz = p["zones"]["NAZ"]
    assert naz["total"] == 0
    assert naz["narrative"] == ["A total of 0 risks identified across business units"]
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_dashboard.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'dashboard'` (or `ImportError: cannot import name 'build_payload'`).

- [ ] **Step 3: Write minimal implementation**

Create `dashboard.py` with only the builder for now:

```python
#!/usr/bin/env python3
"""Local dashboard server for OneTrust zone risk insights."""
import datetime

from zone_reports import stats
from zone_reports import narrative


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
        "narrative": narrative_lines,
    }


def build_payload(rows, view=None, generated_at=None, source=None):
    generated_at = generated_at or datetime.datetime.now()
    date_str = narrative.format_date_ordinal(generated_at.date())
    zones = {}
    for code, org in stats.ZONES.items():
        zs = stats.compute_zone_stats(rows, org)
        lines = narrative.build_narrative(zs, date_str)
        zones[code] = zone_to_dict(zs, lines)
    payload = {
        "generated_at": generated_at.replace(microsecond=0).isoformat(),
        "view": view,
        "cat_order": stats.CAT_ORDER,
        "zones": zones,
    }
    if source:
        payload["source"] = source
    return payload
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_dashboard.py -v`
Expected: PASS (both tests).

- [ ] **Step 5: Commit**

```bash
git add dashboard.py tests/test_dashboard.py
git commit -m "feat: all-zones payload builder for dashboard"
```

---

### Task 2: Fetch runner (subprocess pipeline)

**Files:**
- Modify: `dashboard.py`
- Test: `tests/test_dashboard.py`

**Interfaces:**
- Consumes: `build_payload` (Task 1); `stats.load_rows`.
- Produces:
  - `export_cmd(view, hostname, client_id, client_secret, raw, py) -> list[str]` — builds the exporter argv.
  - `refine_cmd(raw, data, cat_file, py) -> list[str]` — builds the refiner argv.
  - `run_fetch(body, *, runner=subprocess.run, py=sys.executable) -> dict` — validates body, runs both commands, loads `refined.csv`, returns `build_payload(..., source="api")`. On validation failure raises `ValueError`; on a non-zero step raises `RuntimeError(message)`. `runner` is injectable so tests avoid real processes.
  - `load_cached() -> dict | None` — if `refined.csv` exists, return `build_payload(rows, source="cache")`, else `None`.
  - `list_views() -> list[str]` — read names from `views.json`, `[]` if absent.

- [ ] **Step 1: Write the failing test**

```python
# append to tests/test_dashboard.py
import subprocess
import pytest
import dashboard

def test_export_cmd_has_creds_and_all_columns():
    cmd = dashboard.export_cmd("My View", "app-de.onetrust.com", "cid", "sec",
                               "raw_export.csv", py="PY")
    assert cmd[0] == "PY"
    assert "--view" in cmd and "My View" in cmd
    assert "--all-columns" in cmd
    assert "--client-secret" in cmd and "sec" in cmd
    assert cmd[cmd.index("--out") + 1] == "raw_export.csv"

def test_run_fetch_missing_view_raises():
    with pytest.raises(ValueError):
        dashboard.run_fetch({"hostname": "h", "client_id": "c",
                             "client_secret": "s"})

def test_run_fetch_reports_step_failure():
    def fake_runner(cmd, **kw):
        return subprocess.CompletedProcess(cmd, 1, "", "403 Forbidden")
    with pytest.raises(RuntimeError) as e:
        dashboard.run_fetch({"view": "V", "hostname": "h",
                             "client_id": "c", "client_secret": "s"},
                            runner=fake_runner)
    assert "403" in str(e.value)
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_dashboard.py -v`
Expected: FAIL (`AttributeError: module 'dashboard' has no attribute 'export_cmd'`).

- [ ] **Step 3: Write minimal implementation**

Add to `dashboard.py` (imports at top: `import json, subprocess, sys`; `from pathlib import Path`; `HERE = Path(__file__).resolve().parent`):

```python
RAW = "raw_export.csv"
DATA = "refined.csv"
CAT_FILE = "Supplier_Category_List.xlsx"
VIEWS_FILE = HERE / "views.json"


def export_cmd(view, hostname, client_id, client_secret, raw, py=sys.executable):
    return [py, str(HERE / "onetrust_view_export.py"),
            "--view", view, "--all-columns", "--out", raw,
            "--hostname", hostname, "--client-id", client_id,
            "--client-secret", client_secret]


def refine_cmd(raw, data, cat_file, py=sys.executable):
    return [py, str(HERE / "refine_columns.py"),
            "--in", raw, "--out", data, "--cat-file", cat_file]


def _run_step(cmd, runner):
    r = runner(cmd, capture_output=True, text=True)
    if r.returncode != 0:
        tail = (r.stderr or r.stdout or "").strip().splitlines()
        msg = tail[-1] if tail else f"exit {r.returncode}"
        raise RuntimeError(msg)


def run_fetch(body, *, runner=subprocess.run, py=sys.executable):
    for field in ("view", "hostname", "client_id", "client_secret"):
        if not (body.get(field) or "").strip():
            raise ValueError(f"Missing {field}")
    _run_step(export_cmd(body["view"], body["hostname"], body["client_id"],
                         body["client_secret"], RAW, py), runner)
    _run_step(refine_cmd(RAW, DATA, CAT_FILE, py), runner)
    rows = stats.load_rows(DATA)
    return build_payload(rows, view=body["view"], source="api")


def load_cached():
    if not Path(DATA).exists():
        return None
    rows = stats.load_rows(DATA)
    return build_payload(rows, source="cache")


def list_views():
    if not VIEWS_FILE.exists():
        return []
    data = json.loads(VIEWS_FILE.read_text(encoding="utf-8"))
    return [v.get("name", "") for v in data.get("views", []) if v.get("name")]
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_dashboard.py -v`
Expected: PASS (all tests). Note: `_run_step` passes `capture_output/text` to `runner`; the fake runner in the test accepts `**kw`, so this is compatible.

- [ ] **Step 5: Commit**

```bash
git add dashboard.py tests/test_dashboard.py
git commit -m "feat: fetch runner wiring exporter+refiner into payload"
```

---

### Task 3: HTTP server + routes + startup

**Files:**
- Modify: `dashboard.py`
- Test: `tests/test_dashboard.py`

**Interfaces:**
- Consumes: `run_fetch`, `load_cached`, `list_views`, `build_payload` (Tasks 1-2).
- Produces:
  - `handle_get(path) -> (status:int, content_type:str, body:bytes)` — pure router for GET; `/` → index.html bytes (`text/html`), `/api/views` → `{"views":[...]}`, `/api/data` → cached payload or `{"empty": true}`.
  - `handle_post(path, body_bytes) -> (status, content_type, body)` — `/api/fetch` → parse JSON, `run_fetch`, return payload; `ValueError` → 400 `{"error"}`, `RuntimeError`/other → 500 `{"error"}`.
  - `make_handler()` returning a `BaseHTTPRequestHandler` subclass that delegates to the two functions above.
  - `main(argv=None)` — parse `--port` (default 8000), start `HTTPServer(("127.0.0.1", port), ...)`, open browser, `serve_forever`.

- [ ] **Step 1: Write the failing test**

```python
# append to tests/test_dashboard.py
import json as _json

def test_handle_get_views(monkeypatch):
    monkeypatch.setattr(dashboard, "list_views", lambda: ["A", "B"])
    status, ctype, body = dashboard.handle_get("/api/views")
    assert status == 200
    assert "application/json" in ctype
    assert _json.loads(body)["views"] == ["A", "B"]

def test_handle_get_data_empty(monkeypatch):
    monkeypatch.setattr(dashboard, "load_cached", lambda: None)
    status, _c, body = dashboard.handle_get("/api/data")
    assert status == 200
    assert _json.loads(body)["empty"] is True

def test_handle_get_root_serves_html():
    status, ctype, body = dashboard.handle_get("/")
    assert status == 200
    assert "text/html" in ctype
    assert b"<!DOCTYPE html>" in body or b"<!doctype html>" in body

def test_handle_post_fetch_bad_request(monkeypatch):
    def boom(b): raise ValueError("Missing view")
    monkeypatch.setattr(dashboard, "run_fetch", boom)
    status, _c, body = dashboard.handle_post("/api/fetch", b"{}")
    assert status == 400
    assert "Missing view" in _json.loads(body)["error"]

def test_handle_post_fetch_error(monkeypatch):
    def boom(b): raise RuntimeError("403 Forbidden")
    monkeypatch.setattr(dashboard, "run_fetch", boom)
    status, _c, body = dashboard.handle_post("/api/fetch", b'{"view":"V"}')
    assert status == 500
    assert "403" in _json.loads(body)["error"]
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_dashboard.py -v`
Expected: FAIL (`AttributeError: module 'dashboard' has no attribute 'handle_get'`). `test_handle_get_root_serves_html` also fails until Task 4 creates `index.html`; it is expected to pass once Task 4 lands — mark it xfail-free by ordering: create a minimal `index.html` stub in this task's Step 3 so the test passes now.

- [ ] **Step 3: Write minimal implementation**

Add to `dashboard.py` (add imports: `import argparse, webbrowser`; `from http.server import BaseHTTPRequestHandler, HTTPServer`):

```python
INDEX = HERE / "index.html"


def _json_resp(obj, status=200):
    return status, "application/json; charset=utf-8", json.dumps(obj).encode("utf-8")


def handle_get(path):
    if path == "/" or path == "/index.html":
        return 200, "text/html; charset=utf-8", INDEX.read_bytes()
    if path == "/api/views":
        return _json_resp({"views": list_views()})
    if path == "/api/data":
        payload = load_cached()
        return _json_resp(payload if payload else {"empty": True})
    return _json_resp({"error": "not found"}, 404)


def handle_post(path, body_bytes):
    if path != "/api/fetch":
        return _json_resp({"error": "not found"}, 404)
    try:
        body = json.loads(body_bytes or b"{}")
        return _json_resp(run_fetch(body))
    except ValueError as e:
        return _json_resp({"error": str(e)}, 400)
    except Exception as e:  # RuntimeError, subprocess, decode
        return _json_resp({"error": str(e)}, 500)


def make_handler():
    class Handler(BaseHTTPRequestHandler):
        def _send(self, status, ctype, body):
            self.send_response(status)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def do_GET(self):
            self._send(*handle_get(self.path))

        def do_POST(self):
            length = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(length) if length else b""
            self._send(*handle_post(self.path, body))

        def log_message(self, *a):
            pass  # silence; never log request bodies (may hold secrets)
    return Handler


def main(argv=None):
    ap = argparse.ArgumentParser(description="OneTrust risk insights dashboard.")
    ap.add_argument("--port", type=int, default=8000)
    ap.add_argument("--no-browser", action="store_true")
    args = ap.parse_args(argv)
    url = f"http://127.0.0.1:{args.port}/"
    server = HTTPServer(("127.0.0.1", args.port), make_handler())
    print(f"Dashboard at {url}  (Ctrl-C to stop)")
    if not args.no_browser:
        webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")
        server.server_close()


if __name__ == "__main__":
    main()
```

Also create a minimal `index.html` stub so the root route test passes (Task 4 replaces it fully):

```html
<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>Risk Insights</title></head>
<body><div id="app">Loading…</div></body></html>
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_dashboard.py -v`
Expected: PASS (all tests, including `test_handle_get_root_serves_html`).

- [ ] **Step 5: Commit**

```bash
git add dashboard.py index.html tests/test_dashboard.py
git commit -m "feat: dashboard http server routes + startup"
```

---

### Task 4: `index.html` — full UI

**Files:**
- Modify: `index.html` (replace the stub)

**Interfaces:**
- Consumes the API from Task 3: `GET /api/views`, `GET /api/data`, `POST /api/fetch`. Uses the payload shape from Task 1 (`zones[code]` with `total, by_cat, grand_by_cat, by_domain, crosstab, domain_pct, top_cats, narrative`, plus top-level `cat_order, generated_at, view, source`).
- Produces: no downstream consumers (leaf).

This task is UI; verification is manual (browser). No unit test. Build it in one focused pass, then run the manual smoke checklist.

- [ ] **Step 1: Write the full page**

Replace `index.html` with a single self-contained file. Requirements to implement (each maps to the spec):

Structure & regions:
- `<!DOCTYPE html>`, `<meta name="viewport" content="width=device-width, initial-scale=1">`, `<title>Risk Insights Dashboard</title>`.
- Chart.js via CDN: `<script src="https://cdn.jsdelivr.net/npm/chart.js@4"></script>`.
- Header bar: app title; text inputs `#hostname`, `#clientId`, `#clientSecret` (`type="password"`); `<select id="view">`; `<button id="fetchBtn">Fetch</button>`; status pill `<span id="status" role="status" aria-live="polite">`.
- Zone filter: a segmented control `<div id="zones" role="tablist">` — one button per zone code in this order: `GHQ AFR SAZ MAZ NAZ APAC EUR BEES BEES-FT Overall`. Selected button gets `aria-selected="true"` and a highlighted style.
- Insights container with: KPI card row (`#kpis`), a two-column chart row (`<canvas id="pie">`, `<canvas id="bar">`), narrative block (`#narrative`), crosstab table (`#table`).
- Empty state element `#empty` shown when no data.

Behavior (vanilla JS, no framework):
- `let payload = null; let currentZone = "Overall";`
- On load: `fetch('/api/views')` → fill `#view` options; `fetch('/api/data')` → if not `{empty:true}`, set `payload`, render, set status to `Loaded N risks · <generated_at> · cache`.
- Fetch button click: disable button, status → `Fetching…`; `POST /api/fetch` with JSON body `{hostname, client_id, client_secret, view}` from the inputs; on success set `payload`, `currentZone="Overall"`, render, status green `Loaded N risks · <time> · api`; on `{error}` or non-2xx, status red `Error: <msg>` (`#status` gets an `error` class); re-enable button in a `finally`.
- Zone button click: set `currentZone`, update `aria-selected`, call `render()` — no network.
- `render()`:
  - `const z = payload.zones[currentZone];`
  - KPIs: total = `z.total`; top-2 domains from `z.by_domain` with combined `z.domain_pct` sum → "Top domains: X, Y (NN%)"; top supplier cat = max of `z.by_cat` → "name (n)". Handle `z.total === 0` (show zeros / "No risks").
  - Pie: labels = domains from `z.by_domain`, data = counts. Destroy previous Chart instance before re-creating.
  - Stacked bar: x-axis = domains (`z.by_domain` order); one dataset per cat in `payload.cat_order` (skip cats absent from `z.by_cat`); each dataset value per domain = `z.crosstab[domain]?.[cat] || 0`; `options.scales.x.stacked = options.scales.y.stacked = true`. Destroy previous instance first.
  - Narrative: join `z.narrative` lines into `#narrative` (preserve blank lines as spacing; render each line in its own element).
  - Table: header row = `Domain` + `payload.cat_order` (only cats present in `z.by_cat`) + `Total`; body row per domain (in `z.by_domain` order) with `z.crosstab[domain][cat] || 0` and the row total = the domain count; use `font-variant-numeric: tabular-nums`. Make header cells clickable to sort rows by that column (toggle asc/desc; set `aria-sort`).

Styling (inline `<style>`): professional dashboard per the design system — CSS custom properties for slate/indigo tokens, `prefers-color-scheme: dark` overrides, Inter via `font-family: Inter, system-ui, sans-serif`, 8px spacing scale, card surfaces with subtle border/shadow, ≥4.5:1 text contrast, visible `:focus-visible` outlines, buttons ≥44px touch height. Status pill colors: neutral/idle, amber/fetching, green/loaded, red/error — each with a text label (not color alone). No emoji; if icons are wanted use inline Lucide SVG paths.

Before writing, pull exact tokens/palette:

```bash
python skills/ui-ux-pro-max/scripts/search.py "risk compliance analytics dashboard professional" --design-system -f markdown
```

Apply the returned palette, typography, and effects. (The `skills/` path is inside the ui-ux-pro-max skill base dir shown when the skill loads; if not present, use the design-system defaults: slate `#0f172a`/`#1e293b`, indigo `#4f46e5`, surface `#ffffff`/`#0b1220`, Inter.)

- [ ] **Step 2: Manual smoke test**

Run: `python dashboard.py --port 8000`
Then in the browser verify:
1. Page loads; if `refined.csv` exists it renders and status shows `… cache`; else empty state shows.
2. View dropdown is populated from `views.json`.
3. Zone buttons switch KPIs/charts/narrative/table instantly with no network request (check DevTools Network tab).
4. Charts render (pie + stacked bar); table sorts on header click; numbers match the `Overall` figures.
5. Enter host/id/secret + view, click Fetch → status goes `Fetching…` then green `Loaded N risks … api`; `refined.csv` is written.
6. Force an error (blank secret) → red `Error: …` pill; button re-enabled.
7. Toggle OS dark mode → contrast holds.

If `refined.csv` is not available for a no-network check, create a tiny one by hand (headers `ID,Organization,Category,Cat,Stage` + a few rows) to exercise steps 1-4 without hitting the API.

- [ ] **Step 3: Commit**

```bash
git add index.html
git commit -m "feat: dashboard UI - fetch, zone filter, KPI/pie/bar/narrative/table"
```

---

## Self-Review

**Spec coverage:**
- One-command start + auto-open browser → Task 3 `main()`. ✓
- Fetch button + live API pull + cache to `refined.csv` → Task 2 `run_fetch`. ✓
- `/api/views`, `/api/data`, `/api/fetch` → Task 3. ✓
- Credentials entered in UI, posted to localhost, never on disk → Task 2 (args only) + Task 4 (inputs) + Global Constraints. ✓
- Zone filter (10 zones incl. Overall), instant client-side → Task 1 payload has all zones; Task 4 render. ✓
- Insights: KPI, domain pie, stacked bar, narrative, crosstab table → Task 4; data from Task 1. ✓
- All-zones single JSON payload shape → Task 1. ✓
- Reuse `stats.py`/`narrative.py` unchanged → Tasks 1-2. ✓
- Design system (slate/indigo, Inter, dark, a11y) → Task 4. ✓
- Error handling matrix (missing views.json, bad creds/403, missing cat-file, no refined.csv, zero-row zone) → Tasks 2 (`list_views`/`run_fetch`), 3 (handlers), 1 (zero-row zone). ✓
- Security: 127.0.0.1 bind, no logging of bodies → Task 3 (`main`, `log_message`). ✓
- Test `tests/test_dashboard.py` payload shape + totals → Task 1; runner + routes → Tasks 2-3. ✓

**Placeholder scan:** No TBD/TODO; every code step has concrete code; UI step enumerates exact element ids, payload fields, and Chart.js options. ✓

**Type consistency:** `build_payload(rows, view, generated_at, source)` and `zone_to_dict` names match across Tasks 1-3; `run_fetch(body, runner=)`, `export_cmd`/`refine_cmd` args match tests; handler functions `handle_get`/`handle_post` names match tests; payload keys (`by_cat, grand_by_cat, by_domain, crosstab, domain_pct, top_cats, narrative, cat_order`) consistent between Task 1 producer and Task 4 consumer. ✓
