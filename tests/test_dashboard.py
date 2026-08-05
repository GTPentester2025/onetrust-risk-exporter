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

def test_get_config_masked(monkeypatch):
    monkeypatch.setattr(dashboard, "load_config",
        lambda p: {"hostname": "h", "client_id": "c", "client_secret": "S",
                   "schedule": {"mode": "off", "time": "06:00", "interval_hours": 6}})
    status, ctype, body = dashboard.handle_get("/api/config")
    j = _json.loads(body)
    assert status == 200 and "client_secret" not in j and j["has_secret"] is True
    # strict shape — no raw hostname/client_id either
    assert "hostname" not in j and "client_id" not in j
    assert "configured" in j and "host_hint" in j

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
    # strict mask: no raw credential fields
    assert status == 200 and "client_secret" not in j
    assert "hostname" not in j and "client_id" not in j
    assert "configured" in j and "has_secret" in j


def test_post_config_test_uses_saved_secret_fallback(monkeypatch):
    saved_cfg = {
        "hostname": "savedhost",
        "client_id": "savedid",
        "client_secret": "SAVEDSEC",
        "schedule": {"mode": "off", "time": "06:00", "interval_hours": 6},
    }
    monkeypatch.setattr(dashboard, "load_config", lambda p: saved_cfg)

    recorded = {}

    def fake_test_connection(host, cid, sec):
        recorded["host"] = host
        recorded["cid"] = cid
        recorded["sec"] = sec
        return {"ok": True, "message": "ok"}

    monkeypatch.setattr(dashboard, "test_connection", fake_test_connection)

    body = _json.dumps({"hostname": "h2"}).encode()
    status, _c, out = dashboard.handle_post("/api/config/test", body)
    assert status == 200
    assert _json.loads(out)["ok"] is True
    assert recorded["host"] == "h2"
    assert recorded["sec"] == "SAVEDSEC"


def test_post_config_real_roundtrip_preserves_secret(monkeypatch, tmp_path):
    from config_store import save_config as real_save_config

    cfg_file = tmp_path / "config.json"
    monkeypatch.setattr(dashboard, "CONFIG_FILE", cfg_file)

    initial = {
        "hostname": "orighost",
        "client_id": "origid",
        "client_secret": "ORIGSEC",
        "schedule": {"mode": "off", "time": "06:00", "interval_hours": 6},
    }
    real_save_config(str(cfg_file), initial, {
        "hostname": "", "client_id": "", "client_secret": "",
        "schedule": {"mode": "off", "time": "06:00", "interval_hours": 6},
    })

    body = _json.dumps({"hostname": "newhost"}).encode()
    status, _c, out = dashboard.handle_post("/api/config", body)
    j = _json.loads(out)

    assert status == 200
    assert "client_secret" not in j
    assert j.get("has_secret") is True
    # strict mask: no raw credential fields
    assert "hostname" not in j and "client_id" not in j
    assert "configured" in j

    on_disk = _json.loads(cfg_file.read_text(encoding="utf-8"))
    assert on_disk["client_secret"] == "ORIGSEC"
    assert on_disk["hostname"] == "newhost"


def test_get_config_strict_no_raw_values(monkeypatch):
    monkeypatch.setattr(dashboard, "load_config",
        lambda p: {"hostname": "app-de.onetrust.com", "client_id": "cid",
                   "client_secret": "S",
                   "schedule": {"mode": "off", "time": "06:00", "interval_hours": 6}})
    status, _c, body = dashboard.handle_get("/api/config")
    j = _json.loads(body)
    assert status == 200
    assert "hostname" not in j and "client_id" not in j and "client_secret" not in j
    assert j["configured"] is True and j["has_secret"] is True
    assert j["host_hint"].startswith("ap") and j["host_hint"].endswith(".onetrust.com")


def test_files_status_empty(monkeypatch, tmp_path):
    monkeypatch.setattr(dashboard, "CAT_FILE", str(tmp_path / "cat.xlsx"))
    monkeypatch.setattr(dashboard, "VIEWS_FILE", tmp_path / "views.json")
    j = dashboard.files_status()
    assert j["catfile"]["present"] is False
    assert j["views"]["present"] is False


def test_upload_views_valid_and_invalid(monkeypatch, tmp_path):
    monkeypatch.setattr(dashboard, "VIEWS_FILE", tmp_path / "views.json")
    monkeypatch.setattr(dashboard, "CAT_FILE", str(tmp_path / "cat.xlsx"))
    ok = _json.dumps({"views": [{"name": "V"}]}).encode()
    status, _c, body = dashboard.handle_post("/api/upload/views", ok)
    assert status == 200 and _json.loads(body)["views"]["present"] is True
    assert (tmp_path / "views.json").exists()
    bad = b"not json"
    status, _c, body = dashboard.handle_post("/api/upload/views", bad)
    assert status == 400


def test_upload_catfile_magic(monkeypatch, tmp_path):
    monkeypatch.setattr(dashboard, "CAT_FILE", str(tmp_path / "cat.xlsx"))
    monkeypatch.setattr(dashboard, "VIEWS_FILE", tmp_path / "views.json")
    status, _c, _b = dashboard.handle_post("/api/upload/catfile", b"NOTAZIP")
    assert status == 400
    status, _c, body = dashboard.handle_post("/api/upload/catfile", b"PK\x03\x04fakezip")
    assert status == 200 and _json.loads(body)["catfile"]["present"] is True


def test_catfile_download(monkeypatch, tmp_path):
    monkeypatch.setattr(dashboard, "CAT_FILE", str(tmp_path / "cat.xlsx"))
    status, _c, _b = dashboard.handle_get("/api/catfile")
    assert status == 404
    (tmp_path / "cat.xlsx").write_bytes(b"PK\x03\x04data")
    status, ctype, body = dashboard.handle_get("/api/catfile")
    assert status == 200 and body.startswith(b"PK")
    assert "spreadsheetml" in ctype


def test_upload_catfile_size_and_empty(monkeypatch, tmp_path):
    """Test catfile upload validates empty body and size limits."""
    monkeypatch.setattr(dashboard, "CAT_FILE", str(tmp_path / "cat.xlsx"))
    monkeypatch.setattr(dashboard, "VIEWS_FILE", tmp_path / "views.json")
    # Empty body -> 400
    status, _c, body = dashboard.handle_post("/api/upload/catfile", b"")
    assert status == 400
    assert "empty" in _json.loads(body)["error"].lower()
    # Over 10 MB -> 400
    oversized = b"PK" + b"0" * (10 * 1024 * 1024)
    status, _c, body = dashboard.handle_post("/api/upload/catfile", oversized)
    assert status == 400
    assert "10" in _json.loads(body)["error"]


def test_catfile_download_disposition(monkeypatch, tmp_path):
    """Test that /api/catfile GET response includes Content-Disposition attachment header."""
    monkeypatch.setattr(dashboard, "CAT_FILE", str(tmp_path / "cat.xlsx"))
    # Verify _extra_header logic for /api/catfile at 200
    extra = dashboard._extra_header("/api/catfile", 200)
    assert extra is not None
    assert extra[0] == "Content-Disposition"
    assert "attachment" in extra[1]
    assert "Supplier_Category_List.xlsx" in extra[1]
    # Verify it returns None for 404
    extra_404 = dashboard._extra_header("/api/catfile", 404)
    assert extra_404 is None
    # Verify /api/ppt also gets disposition at 200
    extra_ppt = dashboard._extra_header("/api/ppt", 200)
    assert extra_ppt is not None
    assert "risk-insights.pptx" in extra_ppt[1]
    # Verify other paths return None
    extra_other = dashboard._extra_header("/api/config", 200)
    assert extra_other is None
