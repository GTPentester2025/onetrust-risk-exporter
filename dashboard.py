#!/usr/bin/env python3
"""Local dashboard server for OneTrust zone risk insights."""
import datetime
import json
import subprocess
import sys
from pathlib import Path

from zone_reports import stats
from zone_reports import narrative

HERE = Path(__file__).resolve().parent


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
        "cat_order": list(stats.CAT_ORDER),
        "zones": zones,
    }
    if source is not None:
        payload["source"] = source
    return payload


RAW = str(HERE / "raw_export.csv")
DATA = str(HERE / "refined.csv")
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
        msg = "\n".join(tail[-5:]) if tail else f"exit {r.returncode}"
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
