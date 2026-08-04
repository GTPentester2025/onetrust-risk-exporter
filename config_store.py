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
    m = copy.deepcopy({k: v for k, v in cfg.items() if k != "client_secret"})
    m["has_secret"] = bool((cfg.get("client_secret") or "").strip())
    return m


def is_configured(cfg):
    return all(bool((cfg.get(k) or "").strip())
               for k in ("hostname", "client_id", "client_secret"))
