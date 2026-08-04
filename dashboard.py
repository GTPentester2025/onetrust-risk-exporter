#!/usr/bin/env python3
"""Local dashboard server for OneTrust zone risk insights."""
import argparse
import datetime
import json
import subprocess
import sys
import webbrowser
from http.server import BaseHTTPRequestHandler, HTTPServer
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
