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
