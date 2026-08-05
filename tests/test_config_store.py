import json
import datetime as dt
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

def test_default_config_not_mutated(tmp_path):
    cfg = cs.load_config(str(tmp_path / "nope.json"))
    cfg["schedule"]["mode"] = "daily"
    assert cs.DEFAULT_CONFIG["schedule"]["mode"] == "off"

def test_save_overwrites_secret_persisted_to_disk(tmp_path):
    p = str(tmp_path / "config.json")
    existing = {"hostname": "h", "client_id": "c", "client_secret": "OLD",
                "schedule": {"mode": "off", "time": "06:00", "interval_hours": 6}}
    merged = cs.save_config(p, {"client_secret": "NEW"}, existing)
    assert merged["client_secret"] == "NEW"
    assert json.loads(open(p, encoding="utf-8").read())["client_secret"] == "NEW"

def test_is_configured_all_fields():
    base = {"hostname": "h", "client_id": "c", "client_secret": "s"}
    assert cs.is_configured({**base, "hostname": ""}) is False
    assert cs.is_configured({**base, "client_id": ""}) is False
    assert cs.is_configured(base) is True

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
