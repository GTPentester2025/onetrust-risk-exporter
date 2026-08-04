# Dashboard v2 (Power BI shell, editable PPT, config + scheduling) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Turn the dashboard into a user-facing Power BI-style dark console with credentials/schedule behind a gear popup, asynchronous + scheduled server-side fetches, and a fully editable PowerPoint export whose numbers match the screen.

**Architecture:** A new `config_store.py` (pure) owns config load/save/mask + schedule-due logic, persisted to a gitignored `config.json`. `dashboard.py` gains an async fetch job (background thread + shared status), a scheduler daemon, config/status/ppt routes, and drops the creds-in-body fetch. A new `ppt_export.py` builds a native-object deck with python-pptx (auto-installed). `index.html` is rebuilt as a dark BI report canvas with a settings modal and an export menu.

**Tech Stack:** Python 3.14 stdlib (`http.server`, `threading`, `subprocess`, `json`, `datetime`); reused `zone_reports/stats.py` + `narrative.py`, `onetrust_view_export.get_token`; `python-pptx` (auto-installed) for PPT; existing `requests` (already required by the exporter) for the connection test; vanilla HTML/CSS/JS + Chart.js + html2canvas via CDN.

## Global Constraints

- HTTP server stays stdlib-only; the only new pip dependency is `python-pptx`, auto-installed on first PPT use (same pattern as `build_zone_reports.py` for pywin32).
- Server binds `127.0.0.1` only. Request bodies are never logged.
- Client secret is stored only in gitignored `config.json`; `GET /api/config` returns it masked (`has_secret: bool`), never the raw secret; it is never logged or placed in the DOM.
- `.gitignore` must include `config.json` and `*.pptx`.
- Reuse `zone_reports.stats` + `zone_reports.narrative` and `build_payload` unchanged; PPT and UI both derive from the same `build_payload` output so numbers match.
- Zones/order and Cat order come from `stats.ZONES` (10 incl. Overall) and `stats.CAT_ORDER`; `(blank)` Cat is displayed as "Uncategorized".
- View is hard-coded "TPRM Global View".
- Tests use `pytest`, live in `tests/`, require no network, no sockets, no live threads. Existing 20 tests stay green (update only removed-`run_fetch` references).
- Dark theme tokens (exact): `--canvas #1b1a19; --tile #252423; --tile-2 #2f2e2d; --fg #f3f2f1; --fg-muted #c8c6c4; --fg-dim #979593; --border #3b3a39; --gold #e8c810; --gold-deep #c99a1e; --live #34d399; --danger #f0736a; --radius 8px;` Fira Sans / Fira Code.

---

### Task 1: Config store (load / save / mask / is_configured)

**Files:**
- Create: `config_store.py`
- Modify: `.gitignore`
- Test: `tests/test_config_store.py`

**Interfaces:**
- Produces: `DEFAULT_CONFIG` (dict); `load_config(path) -> dict`; `save_config(path, incoming, existing) -> dict`; `mask_config(cfg) -> dict`; `is_configured(cfg) -> bool`. Later tasks import these.

- [ ] **Step 1: Write the failing test**

```python
# tests/test_config_store.py
import json
import config_store as cs

def test_load_missing_returns_default(tmp_path):
    cfg = cs.load_config(str(tmp_path / "nope.json"))
    assert cfg["hostname"] == "" and cfg["client_id"] == "" and cfg["client_secret"] == ""
    assert cfg["schedule"]["mode"] == "off"

def test_load_corrupt_returns_default(tmp_path):
    p = tmp_path / "config.json"; p.write_text("{not json", encoding="utf-8")
    assert cs.load_config(str(p))["schedule"]["mode"] == "off"

def test_save_preserves_secret_when_blank(tmp_path):
    p = str(tmp_path / "config.json")
    existing = {"hostname": "h", "client_id": "c", "client_secret": "SECRET",
                "schedule": {"mode": "off", "time": "06:00", "interval_hours": 6}}
    merged = cs.save_config(p, {"hostname": "h2", "client_secret": ""}, existing)
    assert merged["client_secret"] == "SECRET"      # kept
    assert merged["hostname"] == "h2"               # overwritten
    assert json.loads(open(p, encoding="utf-8").read())["client_secret"] == "SECRET"

def test_save_overwrites_secret_when_provided(tmp_path):
    p = str(tmp_path / "config.json")
    existing = {"hostname": "h", "client_id": "c", "client_secret": "OLD",
                "schedule": {"mode": "off", "time": "06:00", "interval_hours": 6}}
    merged = cs.save_config(p, {"client_secret": "NEW"}, existing)
    assert merged["client_secret"] == "NEW"

def test_mask_hides_secret():
    m = cs.mask_config({"hostname": "h", "client_id": "c", "client_secret": "S",
                        "schedule": {"mode": "off", "time": "06:00", "interval_hours": 6}})
    assert "client_secret" not in m
    assert m["has_secret"] is True
    assert m["hostname"] == "h"

def test_is_configured():
    base = {"hostname": "h", "client_id": "c", "client_secret": "s"}
    assert cs.is_configured(base) is True
    assert cs.is_configured({**base, "client_secret": ""}) is False
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_config_store.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'config_store'`.

- [ ] **Step 3: Write minimal implementation**

```python
# config_store.py
"""Pure config load/save/mask + schedule logic for the dashboard. No HTTP/threads."""
import copy
import json

DEFAULT_CONFIG = {
    "hostname": "",
    "client_id": "",
    "client_secret": "",
    "schedule": {"mode": "off", "time": "06:00", "interval_hours": 6},
}


def load_config(path):
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, json.JSONDecodeError):
        return copy.deepcopy(DEFAULT_CONFIG)
    cfg = copy.deepcopy(DEFAULT_CONFIG)
    for k in ("hostname", "client_id", "client_secret"):
        if isinstance(data.get(k), str):
            cfg[k] = data[k]
    sched = data.get("schedule") or {}
    if isinstance(sched, dict):
        cfg["schedule"].update({k: sched[k] for k in ("mode", "time", "interval_hours")
                                if k in sched})
    return cfg


def save_config(path, incoming, existing):
    merged = copy.deepcopy(existing)
    for k in ("hostname", "client_id"):
        if k in incoming and isinstance(incoming[k], str):
            merged[k] = incoming[k]
    sec = incoming.get("client_secret")
    if isinstance(sec, str) and sec.strip():
        merged["client_secret"] = sec
    if isinstance(incoming.get("schedule"), dict):
        merged.setdefault("schedule", {}).update(incoming["schedule"])
    with open(path, "w", encoding="utf-8") as f:
        json.dump(merged, f, indent=2)
    return merged


def mask_config(cfg):
    m = {k: v for k, v in cfg.items() if k != "client_secret"}
    m["has_secret"] = bool(cfg.get("client_secret"))
    return m


def is_configured(cfg):
    return all(bool((cfg.get(k) or "").strip())
               for k in ("hostname", "client_id", "client_secret"))
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_config_store.py -v`
Expected: PASS (6 tests).

- [ ] **Step 5: Add gitignore entries and commit**

Append to `.gitignore` (each on its own line if not already present): `config.json` and `*.pptx`.

```bash
git add config_store.py tests/test_config_store.py .gitignore
git commit -m "feat: config store (load/save/mask/is_configured) + gitignore"
```

---

### Task 2: Schedule "is due" computation

**Files:**
- Modify: `config_store.py`
- Test: `tests/test_config_store.py`

**Interfaces:**
- Consumes: nothing new.
- Produces: `next_run_due(schedule, last_run, now) -> bool` where `schedule` is the config's `schedule` dict, `last_run` is a `datetime`/`None`, `now` is a `datetime`. Deterministic (no internal clock). The scheduler thread (Task 4) calls this.

- [ ] **Step 1: Write the failing test**

```python
# append to tests/test_config_store.py
import datetime as dt

def test_due_off_never():
    assert cs.next_run_due({"mode": "off"}, None, dt.datetime(2026,8,5,10,0)) is False

def test_due_hourly():
    now = dt.datetime(2026,8,5,10,0)
    assert cs.next_run_due({"mode": "hourly"}, None, now) is True
    assert cs.next_run_due({"mode": "hourly"}, now - dt.timedelta(minutes=30), now) is False
    assert cs.next_run_due({"mode": "hourly"}, now - dt.timedelta(hours=2), now) is True

def test_due_interval():
    now = dt.datetime(2026,8,5,10,0)
    s = {"mode": "interval", "interval_hours": 6}
    assert cs.next_run_due(s, now - dt.timedelta(hours=5), now) is False
    assert cs.next_run_due(s, now - dt.timedelta(hours=7), now) is True

def test_due_daily():
    s = {"mode": "daily", "time": "06:00"}
    # before scheduled time today, never ran -> not due
    assert cs.next_run_due(s, None, dt.datetime(2026,8,5,5,0)) is False
    # after scheduled time today, never ran -> due
    assert cs.next_run_due(s, None, dt.datetime(2026,8,5,7,0)) is True
    # already ran after today's scheduled instant -> not due
    assert cs.next_run_due(s, dt.datetime(2026,8,5,6,30), dt.datetime(2026,8,5,7,0)) is False
    # last ran yesterday, now past today's time -> due
    assert cs.next_run_due(s, dt.datetime(2026,8,4,6,30), dt.datetime(2026,8,5,7,0)) is True
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_config_store.py -k due -v`
Expected: FAIL — `AttributeError: module 'config_store' has no attribute 'next_run_due'`.

- [ ] **Step 3: Write minimal implementation**

Add to `config_store.py` (`import datetime` at top):

```python
def next_run_due(schedule, last_run, now):
    mode = (schedule or {}).get("mode", "off")
    if mode == "off":
        return False
    if mode == "hourly":
        return last_run is None or (now - last_run) >= datetime.timedelta(hours=1)
    if mode == "interval":
        hrs = schedule.get("interval_hours", 6) or 6
        return last_run is None or (now - last_run) >= datetime.timedelta(hours=hrs)
    if mode == "daily":
        hh, mm = (schedule.get("time", "06:00") or "06:00").split(":")
        scheduled_today = now.replace(hour=int(hh), minute=int(mm),
                                      second=0, microsecond=0)
        if now < scheduled_today:
            return False
        return last_run is None or last_run < scheduled_today
    return False
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_config_store.py -v`
Expected: PASS (all config_store tests).

- [ ] **Step 5: Commit**

```bash
git add config_store.py tests/test_config_store.py
git commit -m "feat: schedule next_run_due computation"
```

---

### Task 3: Async fetch job + scheduler + connection test (dashboard.py)

**Files:**
- Modify: `dashboard.py`
- Test: `tests/test_dashboard.py`

**Interfaces:**
- Consumes: `config_store` (Task 1-2); existing `export_cmd`, `refine_cmd`, `_run_step`, `stats.load_rows`, `build_payload`; `onetrust_view_export.get_token`.
- Produces:
  - `CONFIG_FILE` (Path), `VIEW = "TPRM Global View"`, module `JOB` dict, `reset_job()` (test helper).
  - `do_fetch(cfg, *, runner=subprocess.run, py=sys.executable) -> int` — run export+refine using `cfg` creds, return refined row count; raise `ValueError` if `cfg` not configured, `RuntimeError` on a failed step.
  - `run_fetch_job(*, config_loader=None, runner=subprocess.run, py=sys.executable, now_fn=None) -> None` — update `JOB` around `do_fetch`; single-flight (no-op if already running).
  - `start_fetch_async() -> dict` → `{"started": bool, "running": bool}`.
  - `scheduler_tick(now, *, config_loader=None) -> bool` — pure-ish: if `next_run_due`, call `start_fetch_async`; return whether it started (the daemon loop calls this on a timer).
  - `test_connection(hostname, client_id, client_secret) -> dict` → `{"ok": bool, "message": str}`.
- Note: the old `run_fetch(body, ...)` is REMOVED; `handle_post` stops calling it (Task 4). Keep `export_cmd`/`refine_cmd`/`_run_step`.

- [ ] **Step 1: Write the failing test**

```python
# append to tests/test_dashboard.py
import subprocess, datetime
import dashboard

CFG = {"hostname": "h", "client_id": "c", "client_secret": "s",
       "schedule": {"mode": "off", "time": "06:00", "interval_hours": 6}}

def test_do_fetch_requires_config(monkeypatch, tmp_path):
    with pytest.raises(ValueError):
        dashboard.do_fetch({"hostname": "", "client_id": "", "client_secret": ""})

def test_do_fetch_runs_both_steps_and_counts(monkeypatch, tmp_path):
    calls = []
    def fake_runner(cmd, **kw):
        calls.append(cmd)
        return subprocess.CompletedProcess(cmd, 0, "", "")
    monkeypatch.setattr(dashboard.stats, "load_rows", lambda p: [{"a": 1}, {"a": 2}])
    n = dashboard.do_fetch(CFG, runner=fake_runner)
    assert n == 2
    assert len(calls) == 2  # export then refine

def test_run_fetch_job_sets_done(monkeypatch):
    dashboard.reset_job()
    monkeypatch.setattr(dashboard, "do_fetch", lambda cfg, **kw: 7)
    fixed = datetime.datetime(2026, 8, 5, 6, 0)
    dashboard.run_fetch_job(config_loader=lambda: CFG, now_fn=lambda: fixed)
    assert dashboard.JOB["state"] == "done"
    assert dashboard.JOB["rows"] == 7
    assert dashboard.JOB["last_run"] == fixed

def test_run_fetch_job_error(monkeypatch):
    dashboard.reset_job()
    def boom(cfg, **kw): raise RuntimeError("403 Forbidden")
    monkeypatch.setattr(dashboard, "do_fetch", boom)
    dashboard.run_fetch_job(config_loader=lambda: CFG, now_fn=lambda: datetime.datetime(2026,8,5))
    assert dashboard.JOB["state"] == "error"
    assert "403" in dashboard.JOB["message"]

def test_test_connection_ok(monkeypatch):
    monkeypatch.setattr(dashboard, "get_token", lambda h, c, s: "tok")
    r = dashboard.test_connection("h", "c", "s")
    assert r["ok"] is True

def test_test_connection_fail(monkeypatch):
    def boom(h, c, s): raise SystemExit("No access_token: bad creds")
    monkeypatch.setattr(dashboard, "get_token", boom)
    r = dashboard.test_connection("h", "c", "s")
    assert r["ok"] is False and "access_token" in r["message"]
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_dashboard.py -k "do_fetch or run_fetch_job or test_connection" -v`
Expected: FAIL — `AttributeError: module 'dashboard' has no attribute 'do_fetch'`.

- [ ] **Step 3: Write minimal implementation**

In `dashboard.py`: add `import threading`; add `from config_store import (load_config, save_config, mask_config, is_configured, next_run_due)`; add `from onetrust_view_export import get_token`. Replace the `run_fetch` function with the following block (keep `export_cmd`, `refine_cmd`, `_run_step`, `load_cached`, `list_views`):

```python
CONFIG_FILE = HERE / "config.json"
VIEW = "TPRM Global View"

_job_lock = threading.Lock()
JOB = {"state": "idle", "last_run": None, "rows": None, "message": "", "source": None}


def reset_job():
    global JOB
    JOB = {"state": "idle", "last_run": None, "rows": None, "message": "", "source": None}


def do_fetch(cfg, *, runner=subprocess.run, py=sys.executable):
    if not is_configured(cfg):
        raise ValueError("Not configured")
    _run_step(export_cmd(VIEW, cfg["hostname"], cfg["client_id"],
                         cfg["client_secret"], RAW, py), runner)
    _run_step(refine_cmd(RAW, DATA, CAT_FILE, py), runner)
    return len(stats.load_rows(DATA))


def run_fetch_job(*, config_loader=None, runner=subprocess.run,
                  py=sys.executable, now_fn=None):
    config_loader = config_loader or (lambda: load_config(str(CONFIG_FILE)))
    now_fn = now_fn or datetime.datetime.now
    with _job_lock:
        if JOB["state"] == "running":
            return
        JOB["state"] = "running"; JOB["message"] = ""
    try:
        n = do_fetch(config_loader(), runner=runner, py=py)
        JOB.update(state="done", rows=n, last_run=now_fn(), source="api", message="")
    except Exception as e:
        JOB.update(state="error", message=str(e))


def start_fetch_async():
    with _job_lock:
        if JOB["state"] == "running":
            return {"started": False, "running": True}
    t = threading.Thread(target=run_fetch_job, daemon=True)
    t.start()
    return {"started": True, "running": False}


def scheduler_tick(now, *, config_loader=None):
    cfg = (config_loader or (lambda: load_config(str(CONFIG_FILE))))()
    if not is_configured(cfg):
        return False
    if next_run_due(cfg.get("schedule", {}), JOB["last_run"], now) and JOB["state"] != "running":
        return start_fetch_async().get("started", False)
    return False


def _scheduler_loop(stop_event):
    while not stop_event.wait(30):
        try:
            scheduler_tick(datetime.datetime.now())
        except Exception:
            pass  # never let the daemon die on a transient error


def test_connection(hostname, client_id, client_secret):
    try:
        get_token(hostname, client_id, client_secret)
        return {"ok": True, "message": "Connection OK"}
    except SystemExit as e:
        return {"ok": False, "message": str(e)}
    except Exception as e:
        return {"ok": False, "message": f"{type(e).__name__}: {e}"}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_dashboard.py -k "do_fetch or run_fetch_job or test_connection" -v`
Expected: PASS. Also run the whole file: `python -m pytest tests/test_dashboard.py -v` — the old `run_fetch` tests will now FAIL because `run_fetch` was removed; DELETE those three obsolete tests (`test_run_fetch_missing_view_raises`, `test_run_fetch_reports_step_failure`, and any other referencing `dashboard.run_fetch`) in this step, then re-run so the file is green.

- [ ] **Step 5: Commit**

```bash
git add dashboard.py tests/test_dashboard.py
git commit -m "feat: async fetch job, scheduler tick, connection test"
```

---

### Task 4: Config/status/fetch routes + scheduler startup (dashboard.py)

**Files:**
- Modify: `dashboard.py`
- Test: `tests/test_dashboard.py`

**Interfaces:**
- Consumes: Task 3 (`JOB`, `start_fetch_async`, `test_connection`, `CONFIG_FILE`), `config_store` (`load_config`, `save_config`, `mask_config`).
- Produces: extended `handle_get` (`/api/config`, `/api/status`) and `handle_post` (`/api/config`, `/api/config/test`, `/api/fetch` now body-less/async), `_job_status()` helper, scheduler thread started in `main()`.

- [ ] **Step 1: Write the failing test**

```python
# append to tests/test_dashboard.py
import json as _json

def test_get_config_masked(monkeypatch):
    monkeypatch.setattr(dashboard, "load_config",
        lambda p: {"hostname": "h", "client_id": "c", "client_secret": "S",
                   "schedule": {"mode": "off", "time": "06:00", "interval_hours": 6}})
    status, ctype, body = dashboard.handle_get("/api/config")
    j = _json.loads(body)
    assert status == 200 and "client_secret" not in j and j["has_secret"] is True

def test_get_status(monkeypatch):
    dashboard.reset_job()
    status, _c, body = dashboard.handle_get("/api/status")
    j = _json.loads(body)
    assert j["state"] == "idle" and j["last_run"] is None

def test_post_fetch_starts(monkeypatch):
    monkeypatch.setattr(dashboard, "start_fetch_async", lambda: {"started": True, "running": False})
    status, _c, body = dashboard.handle_post("/api/fetch", b"")
    assert status == 202 and _json.loads(body)["started"] is True

def test_post_config_test(monkeypatch):
    monkeypatch.setattr(dashboard, "test_connection", lambda h, c, s: {"ok": True, "message": "Connection OK"})
    body = _json.dumps({"hostname": "h", "client_id": "c", "client_secret": "s"}).encode()
    status, _c, out = dashboard.handle_post("/api/config/test", body)
    assert status == 200 and _json.loads(out)["ok"] is True

def test_post_config_saves_and_masks(monkeypatch):
    saved = {}
    monkeypatch.setattr(dashboard, "load_config", lambda p: {"hostname": "", "client_id": "",
        "client_secret": "OLD", "schedule": {"mode": "off", "time": "06:00", "interval_hours": 6}})
    def fake_save(p, incoming, existing):
        saved.update(existing); saved.update({k: v for k, v in incoming.items() if v})
        return {**existing, **{k: v for k, v in incoming.items() if v}}
    monkeypatch.setattr(dashboard, "save_config", fake_save)
    body = _json.dumps({"hostname": "h2"}).encode()
    status, _c, out = dashboard.handle_post("/api/config", body)
    j = _json.loads(out)
    assert status == 200 and "client_secret" not in j and j["hostname"] == "h2"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_dashboard.py -k "config or get_status or post_fetch" -v`
Expected: FAIL — `/api/config` currently routes to 404 / `handle_post` only knows `/api/fetch` with a body.

- [ ] **Step 3: Write minimal implementation**

Replace `handle_get` and `handle_post` in `dashboard.py` with:

```python
def _job_status():
    j = JOB
    return {"state": j["state"],
            "last_run": j["last_run"].isoformat() if j["last_run"] else None,
            "rows": j["rows"], "message": j["message"], "source": j["source"]}


def handle_get(path):
    if path == "/" or path == "/index.html":
        try:
            return 200, "text/html; charset=utf-8", INDEX.read_bytes()
        except FileNotFoundError:
            return _json_resp({"error": "index.html not found"}, 500)
    if path == "/api/config":
        return _json_resp(mask_config(load_config(str(CONFIG_FILE))))
    if path == "/api/status":
        return _json_resp(_job_status())
    if path == "/api/data":
        payload = load_cached()
        return _json_resp(payload if payload else {"empty": True})
    if path == "/api/ppt":
        return _handle_ppt()          # defined in Task 5
    return _json_resp({"error": "not found"}, 404)


def handle_post(path, body_bytes):
    try:
        if path == "/api/fetch":
            r = start_fetch_async()
            return _json_resp(r, 202)
        if path == "/api/config":
            body = json.loads(body_bytes or b"{}")
            merged = save_config(str(CONFIG_FILE), body, load_config(str(CONFIG_FILE)))
            return _json_resp(mask_config(merged))
        if path == "/api/config/test":
            body = json.loads(body_bytes or b"{}")
            existing = load_config(str(CONFIG_FILE))
            host = body.get("hostname") or existing["hostname"]
            cid = body.get("client_id") or existing["client_id"]
            sec = body.get("client_secret") or existing["client_secret"]
            return _json_resp(test_connection(host, cid, sec))
        return _json_resp({"error": "not found"}, 404)
    except Exception as e:
        return _json_resp({"error": str(e)}, 500)
```

Add a temporary stub so imports resolve until Task 5 lands the real one — place near the other helpers:

```python
def _handle_ppt():
    return _json_resp({"error": "PPT not available"}, 500)
```

In `main()`, start the scheduler thread before `serve_forever()`:

```python
    stop_event = threading.Event()
    threading.Thread(target=_scheduler_loop, args=(stop_event,), daemon=True).start()
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_dashboard.py -v`
Expected: PASS (all dashboard tests, including the updated set from Task 3).

- [ ] **Step 5: Commit**

```bash
git add dashboard.py tests/test_dashboard.py
git commit -m "feat: config/status/async-fetch routes + scheduler startup"
```

---

### Task 5: Editable PowerPoint export (ppt_export.py + /api/ppt)

**Files:**
- Create: `ppt_export.py`
- Modify: `dashboard.py` (replace the `_handle_ppt` stub)
- Test: `tests/test_ppt_export.py`

**Interfaces:**
- Consumes: `build_payload` output shape (`zones[code]` with `total, by_cat, by_domain, crosstab, domain_pct, narrative`, top-level `cat_order, view, generated_at`).
- Produces: `ensure_pptx()`; `build_deck(payload) -> Presentation`; `deck_to_bytes(payload) -> bytes`. `dashboard._handle_ppt()` calls `deck_to_bytes`.

- [ ] **Step 1: Write the failing test**

```python
# tests/test_ppt_export.py
import io
import pytest

pptx = pytest.importorskip("pptx")  # skip cleanly if python-pptx not installed
from pptx import Presentation
import ppt_export

def _payload():
    zones = {}
    for i, code in enumerate(["GHQ", "AFR", "SAZ", "MAZ", "NAZ", "APAC",
                              "EUR", "BEES", "BEES-FT", "Overall"]):
        zones[code] = {
            "organization": code, "total": 3,
            "by_cat": {"COMMERCIAL": 2, "NCI": 1},
            "grand_by_cat": [["COMMERCIAL", 2], ["NCI", 1]],
            "by_domain": [["Security", 2], ["Privacy", 1]],
            "crosstab": {"Security": {"COMMERCIAL": 2}, "Privacy": {"NCI": 1}},
            "domain_pct": {"Security": 67, "Privacy": 33},
            "top_cats": {"Security": [["COMMERCIAL", 2]]},
            "narrative": ["A total of 3 risks identified across business units"],
        }
    return {"view": "TPRM Global View", "generated_at": "2026-08-05T06:00:00",
            "cat_order": ["COMMERCIAL", "LOGISTICS", "NCI", "PACKAGING",
                          "TECHNOLOGY", "RAU", "Fees"], "zones": zones}

def test_deck_has_title_plus_one_slide_per_zone():
    b = ppt_export.deck_to_bytes(_payload())
    prs = Presentation(io.BytesIO(b))
    assert len(prs.slides) == 11  # title + 10 zones

def test_zone_slide_has_chart_and_table():
    prs = Presentation(io.BytesIO(ppt_export.deck_to_bytes(_payload())))
    zone_slide = prs.slides[1]
    has_chart = any(sh.has_chart for sh in zone_slide.shapes)
    has_table = any(sh.has_table for sh in zone_slide.shapes)
    assert has_chart and has_table
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_ppt_export.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'ppt_export'` (or skips if pptx missing — install first: `python -m pip install python-pptx`).

- [ ] **Step 3: Write minimal implementation**

```python
# ppt_export.py
"""Build an editable PowerPoint deck (native charts/tables) from a dashboard payload."""
import io
import subprocess
import sys


def ensure_pptx():
    try:
        import pptx  # noqa: F401
    except ImportError:
        subprocess.run([sys.executable, "-m", "pip", "install", "python-pptx"], check=True)


def _label(cat):
    return "Uncategorized" if cat == "(blank)" else cat


def _active_cats(payload, z):
    known = [c for c in payload["cat_order"] if c in z["by_cat"]]
    extras = [c for c in z["by_cat"] if c not in payload["cat_order"]]
    return known + extras


def build_deck(payload):
    ensure_pptx()
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # Title slide
    s = prs.slides.add_slide(blank)
    tb = s.shapes.add_textbox(Inches(0.7), Inches(2.6), Inches(12), Inches(2))
    tf = tb.text_frame
    tf.text = "OneTrust Risk Insights"
    tf.paragraphs[0].runs[0].font.size = Pt(40)
    p = tf.add_paragraph()
    p.text = f'{payload.get("view", "")}  ·  generated {payload.get("generated_at", "")}'
    p.font.size = Pt(16)

    for code, z in payload["zones"].items():
        s = prs.slides.add_slide(blank)
        head = s.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.3), Inches(0.8))
        head.text_frame.text = f"{code} — {z['total']} risks"
        head.text_frame.paragraphs[0].runs[0].font.size = Pt(26)

        # KPI text
        kpi = s.shapes.add_textbox(Inches(0.5), Inches(1.1), Inches(12.3), Inches(0.6))
        doms = z["by_domain"]
        top2 = ", ".join(f"{d} ({z['domain_pct'].get(d,0)}%)" for d, _ in doms[:2]) or "—"
        topcat = max(z["by_cat"].items(), key=lambda kv: kv[1])[0] if z["by_cat"] else "—"
        kpi.text_frame.text = (f"Total {z['total']}   |   Top domains: {top2}"
                               f"   |   Top category: {_label(topcat)}")
        kpi.text_frame.paragraphs[0].font.size = Pt(12)

        # Pie: risks by domain (native, editable)
        if doms:
            cd = CategoryChartData()
            cd.categories = [d for d, _ in doms]
            cd.add_series("Risks", [n for _, n in doms])
            s.shapes.add_chart(XL_CHART_TYPE.PIE, Inches(0.5), Inches(1.9),
                               Inches(6), Inches(2.6), cd)

        # Stacked bar: domain x Cat (native, editable)
        cats = _active_cats(payload, z)
        if doms and cats:
            cd = CategoryChartData()
            cd.categories = [d for d, _ in doms]
            for c in cats:
                cd.add_series(_label(c), [z["crosstab"].get(d, {}).get(c, 0) for d, _ in doms])
            ch = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, Inches(6.7), Inches(1.9),
                                    Inches(6.1), Inches(2.6), cd)
            ch.chart.has_legend = True

        # Crosstab table (native)
        rows = len(doms) + 1
        colcats = cats
        ncol = 2 + len(colcats)
        tbl = s.shapes.add_table(rows if rows > 1 else 2, ncol,
                                 Inches(0.5), Inches(4.7), Inches(8.3), Inches(2.4)).table
        tbl.cell(0, 0).text = "Domain"
        for j, c in enumerate(colcats):
            tbl.cell(0, 1 + j).text = _label(c)
        tbl.cell(0, ncol - 1).text = "Total"
        for i, (d, dn) in enumerate(doms, start=1):
            tbl.cell(i, 0).text = d
            for j, c in enumerate(colcats):
                tbl.cell(i, 1 + j).text = str(z["crosstab"].get(d, {}).get(c, 0))
            tbl.cell(i, ncol - 1).text = str(dn)

        # Narrative text box
        nb = s.shapes.add_textbox(Inches(9.0), Inches(4.7), Inches(3.8), Inches(2.4))
        nb.text_frame.word_wrap = True
        nb.text_frame.text = "\n".join(z["narrative"]) or "No data"
        for para in nb.text_frame.paragraphs:
            para.font.size = Pt(9)

    return prs


def deck_to_bytes(payload):
    prs = build_deck(payload)
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()
```

Replace the `_handle_ppt` stub in `dashboard.py` with the real handler:

```python
def _handle_ppt():
    payload = load_cached()
    if not payload:
        return _json_resp({"error": "No data — refresh first"}, 400)
    try:
        import ppt_export
        data = ppt_export.deck_to_bytes(payload)
    except Exception as e:
        return _json_resp({"error": f"PPT build failed: {e}"}, 500)
    ctype = ("application/vnd.openxmlformats-officedocument"
             ".presentationml.presentation")
    return 200, ctype, data
```

Because the `.pptx` response needs a download filename, extend the `Handler._send`/`do_GET` path so an `/api/ppt` 200 sets `Content-Disposition`. Simplest: in `make_handler()`'s `do_GET`, after computing the tuple, if `self.path == "/api/ppt"` and status == 200, add the header:

```python
        def do_GET(self):
            status, ctype, body = handle_get(self.path)
            extra = None
            if self.path == "/api/ppt" and status == 200:
                extra = ("Content-Disposition", 'attachment; filename="risk-insights.pptx"')
            self._send(status, ctype, body, extra)
```

and update `_send` to accept an optional `extra` header tuple:

```python
        def _send(self, status, ctype, body, extra=None):
            self.send_response(status)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(len(body)))
            if extra:
                self.send_header(*extra)
            self.end_headers()
            self.wfile.write(body)
```

(Update `do_POST` to call `self._send(*handle_post(...))` — unchanged signature works since `extra` defaults to None.)

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_ppt_export.py -v` then `python -m pytest tests/ -v`
Expected: PASS (PPT tests pass with python-pptx installed; whole suite green).

- [ ] **Step 5: Commit**

```bash
git add ppt_export.py dashboard.py tests/test_ppt_export.py
git commit -m "feat: editable PowerPoint deck export + /api/ppt route"
```

---

### Task 6: Power BI dark UI redesign (index.html)

**Files:**
- Modify: `index.html` (rebuild)

**Interfaces:**
- Consumes the routes: `GET /api/config` (masked: `has_secret`, hostname, client_id, schedule), `POST /api/config`, `POST /api/config/test`, `POST /api/fetch` (202 async), `GET /api/status` (poll), `GET /api/data` (payload/empty), `GET /api/ppt` (download). Payload shape unchanged from v1.
- Produces: leaf (no consumers).

This task is UI; verification is manual (browser) plus a headless wiring smoke. Build in one focused pass.

- [ ] **Step 1: Rebuild index.html**

Keep all existing data rendering logic (KPI, pie, bar, narrative, table, `catLabel`, `activeCatsFor`, chart destroy-before-recreate, per-tile copy/download, html2canvas zone screenshot). Change the shell, remove the credentials from the main page, and add the config modal + async refresh + export menu. Requirements:

Head: keep Fira fonts, Chart.js, html2canvas CDNs.

Theme: single dark theme using the Global-Constraints tokens. Remove any prior gold-on-black background glows in favor of a flat Power BI canvas: `--canvas` page background (flat, no radial glow), `--tile` cards with 1px `--border`, `--radius` 8px, subtle shadow `0 1px 2px rgba(0,0,0,.4)`. Fira Sans body, Fira Code for numbers/mono. Contrast ≥4.5:1 (use `--fg`/`--fg-muted`, avoid `--fg-dim` for small body text). `prefers-reduced-motion` disables transitions. No emoji; inline Lucide-style SVG icons only.

App bar (`#appbar`, sticky top): left product mark in gold; a **zone tab strip** `#zones role="tablist"` (the 10 zones GHQ AFR SAZ MAZ NAZ APAC EUR BEES BEES-FT Overall; active tab has a gold underline + `aria-selected="true"`; switching only re-renders from cached payload — NO fetch); right cluster: **Refresh** button (`#refreshBtn`), a **status chip** `#status role="status" aria-live="polite"`, an **Export ▾** menu button (`#exportBtn`) opening a small menu with items "PowerPoint (.pptx)", "Zone image (PNG)", "Crosstab CSV", and a **Settings** gear button (`#settingsBtn`, aria-label "Settings").

Report canvas (`#main`): flat `--canvas`, centered max-width ~1280px, tiles in a responsive grid:
- KPI tiles row (`#kpis`): total risks; top-2 domains combined %; top supplier category. Big Fira Code numbers in gold; cream labels.
- Row: pie tile (`<canvas id="pie">`) | stacked-bar tile (`<canvas id="bar">`).
- Row: crosstab table tile (`#table`) | narrative tile (`#narrative`).
- Each tile = header (title + hover-revealed toolbar with Copy + Download icon buttons, aria-labels + titles) + body. Keep the existing per-element copy/download behaviors (pie/bar PNG composited on opaque `--tile`; table CSV/TSV; narrative txt/text; KPI png/text) and the toast.

Onboarding banner (`#banner`): shown when `GET /api/config` returns `has_secret === false` (or hostname/client_id empty). Text "Connect your OneTrust data to load insights." + a button that opens Settings. Dismissible.

Settings modal (`#settings`, hidden by default): scrim, centered card, Escape + close button, focus moves into the modal on open and back to the gear on close.
- Connection: `#cfgHost` (hostname), `#cfgId` (client id), `#cfgSecret` (password; placeholder "•••• saved (leave blank to keep)" when `has_secret`, else "client secret"); **Test connection** button (`#cfgTest`) → POST /api/config/test with the three fields → show a green "Connection OK" or red message inline (`#cfgTestResult`).
- Schedule: `#cfgMode` select (Off / Hourly / Daily / Every N hours = values off|hourly|daily|interval); `#cfgTime` (type=time, shown when Daily); `#cfgInterval` (number hours, shown when interval).
- **Save** (`#cfgSave`) → POST /api/config with only changed fields (omit `client_secret` if blank so the server keeps it) → on success close modal, refresh the banner state, toast "Settings saved".
- Never prefill the secret field from the server (server never sends it).

Refresh + polling: `#refreshBtn` → POST /api/fetch; then poll `GET /api/status` every 2s. Status chip states: idle ""; running "Refreshing…" + spinner (disable Refresh); done → set chip green "Updated <last_run local time>" and call `loadData()` to re-render; error → red "Error: <message>". Stop polling when state is done/error/idle. On initial page load: GET /api/config (banner + never fill secret), GET /api/status (reflect an in-progress/last run), GET /api/data (render cache if present). Also poll status while the page is open at a slow cadence (e.g., every 30s) so scheduled runs refresh the view automatically; when a poll shows a newer `last_run` than last seen, call `loadData()`.

Export menu:
- "PowerPoint (.pptx)": if payload is null → toast "No data yet"; else navigate/download `GET /api/ppt` (e.g., set `window.location = '/api/ppt'` or fetch→blob→download); the server streams the file. Show a brief "Building deck…" toast.
- "Zone image (PNG)": existing html2canvas of the current zone's tiles → `<zone>-dashboard.png`.
- "Crosstab CSV": existing current-zone CSV download.

Guards: all export/copy handlers no-op with a toast when payload is null. Zone switching never triggers a network call.

- [ ] **Step 2: Manual + headless smoke**

Headless wiring (no browser):
```
python dashboard.py --no-browser --port 8041   # background
curl -s http://127.0.0.1:8041/ | grep -c "settingsBtn"      # >=1
curl -s http://127.0.0.1:8041/ | grep -c "exportBtn"        # >=1
curl -s http://127.0.0.1:8041/ | grep -c -- "--canvas"      # >=1 (dark tokens)
curl -s http://127.0.0.1:8041/api/config                    # masked JSON, no client_secret
curl -s http://127.0.0.1:8041/api/status                    # {"state":"idle",...}
# stop the server
```
Also extract the `<script>` and run `node --check` — must pass.
Browser (manual, user): open Settings → enter creds → Test connection → Save; click Refresh → chip shows Refreshing then Updated; switch zones (no network in DevTools); Export ▾ → PowerPoint downloads and opens in PowerPoint with editable charts/table; toggle a Daily schedule and confirm it persists in config.json.

- [ ] **Step 3: Commit**

```bash
git add index.html
git commit -m "feat: Power BI dark UI, settings modal, async refresh, export menu"
```

---

## Self-Review

**Spec coverage:**
- Config store (load/save/mask/is_configured) → Task 1. ✓
- Schedule due logic → Task 2. ✓
- Async fetch job + scheduler + connection test → Task 3. ✓
- Config/status/fetch routes + scheduler startup → Task 4. ✓
- Editable PPT deck + /api/ppt + download filename → Task 5. ✓
- Power BI dark UI, settings modal, export menu, async refresh + polling, onboarding banner → Task 6. ✓
- Secret masked/never in browser → Tasks 1 (mask), 4 (routes), 6 (never prefill). ✓
- .gitignore config.json + *.pptx → Task 1. ✓
- python-pptx auto-install → Task 5 (`ensure_pptx`). ✓
- Data accuracy (PPT from build_payload) → Task 5 consumes payload; UI + PPT same source. ✓
- 127.0.0.1 bind / no body logging → unchanged from v1 (`main`, `log_message`); scheduler added in `main` (Task 4). ✓
- Reduced-motion, no-emoji, contrast, zone-switch-no-refetch → Task 6. ✓
- Error matrix (not configured, test fail, fetch fail, ppt no-data, pptx install fail, corrupt config) → Tasks 1 (corrupt), 3 (fetch/test), 5 (ppt no-data/install), 6 (UI surfacing). ✓
- Tests no network/threads/sockets; existing 20 green (minus removed run_fetch, replaced) → Tasks 1-5. ✓

**Placeholder scan:** No TBD/TODO; every code step has concrete code; UI task enumerates exact element ids, routes, and behaviors. The Task 4 `_handle_ppt` stub is explicitly replaced in Task 5. ✓

**Type consistency:** `do_fetch(cfg,...)`, `run_fetch_job(config_loader=,runner=,py=,now_fn=)`, `start_fetch_async()->{started,running}`, `scheduler_tick(now,config_loader=)`, `test_connection(host,cid,sec)->{ok,message}`, `JOB` keys (`state,last_run,rows,message,source`), `_job_status()`, `mask_config`→`has_secret`, `deck_to_bytes(payload)->bytes`, `_handle_ppt()` — names consistent across Tasks 3-6 and the tests. Route set consistent between Task 4 handlers and Task 6 consumers. ✓
