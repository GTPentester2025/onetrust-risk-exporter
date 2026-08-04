import datetime
import subprocess
import pytest
import json as _json
from dashboard import build_payload
import dashboard

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

def test_export_cmd_has_creds_and_all_columns():
    cmd = dashboard.export_cmd("My View", "app-de.onetrust.com", "cid", "sec",
                               "raw_export.csv", py="PY")
    assert cmd[0] == "PY"
    assert "--view" in cmd and "My View" in cmd
    assert "--all-columns" in cmd
    assert "--client-secret" in cmd and "sec" in cmd
    assert cmd[cmd.index("--out") + 1] == "raw_export.csv"


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
    assert dashboard.JOB["source"] == "api"

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

def test_start_fetch_async_single_flight(monkeypatch):
    # Branch 1: already running -> no new thread, returns {started:False, running:True}
    dashboard.reset_job()
    dashboard.JOB["state"] = "running"
    result = dashboard.start_fetch_async()
    assert result == {"started": False, "running": True}

    # Branch 2: idle -> claims slot under lock, spawns thread, returns {started:True}
    dashboard.reset_job()
    monkeypatch.setattr(dashboard, "run_fetch_job", lambda **kw: None)
    result = dashboard.start_fetch_async()
    assert result["started"] is True
    # The slot was claimed under the lock before the (no-op) thread ran,
    # so state remains "running" (no-op never sets done/error).
    assert dashboard.JOB["state"] == "running"

def test_scheduler_tick_starts_when_due(monkeypatch):
    started_calls = []
    monkeypatch.setattr(dashboard, "start_fetch_async",
                        lambda: started_calls.append(1) or {"started": True, "running": False})
    cfg = {
        "hostname": "h", "client_id": "c", "client_secret": "s",
        "schedule": {"mode": "hourly", "time": "06:00", "interval_hours": 1},
    }
    dashboard.reset_job()
    result = dashboard.scheduler_tick(
        datetime.datetime(2026, 8, 5, 10, 0),
        config_loader=lambda: cfg,
    )
    assert result is True
    assert started_calls  # start_fetch_async was called

def test_scheduler_tick_skips_when_not_configured(monkeypatch):
    cfg = {"hostname": "", "client_id": "", "client_secret": "", "schedule": {}}
    dashboard.reset_job()
    result = dashboard.scheduler_tick(
        datetime.datetime(2026, 8, 5, 10, 0),
        config_loader=lambda: cfg,
    )
    assert result is False

def test_test_connection_exception_branch(monkeypatch):
    monkeypatch.setattr(dashboard, "get_token",
                        lambda h, c, s: (_ for _ in ()).throw(ConnectionError("boom")))
    r = dashboard.test_connection("h", "c", "s")
    assert r["ok"] is False
    assert "boom" in r["message"]

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
