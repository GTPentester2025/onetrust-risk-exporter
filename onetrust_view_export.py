#!/usr/bin/env python3
"""
OneTrust Risk Register -- export a saved view's risks to CSV.

Your OAuth2 client-credentials token is NOT authorized for OneTrust's internal
risk-views / reports-export endpoints (those are gated to logged-in user
sessions and return 403 to API tokens). This tool therefore uses the
OAuth-accessible grid endpoint, which returns risk rows as JSON:

  POST /api/risk/v2/risks/views/pages   (filters -> rows; paginated)

Each saved view's filters live in views.json (captured from the UI once, since
the token cannot read view definitions). The script posts a view's filters,
pages through all rows, and writes a flattened CSV.

Auth: OAuth2 client credentials (id + secret) -> token, or ONETRUST_TOKEN.

Usage:
    python onetrust_view_export.py                 # pick a view, export
    python onetrust_view_export.py --list
    python onetrust_view_export.py --view "TPRM Global View" --out tprm.csv
    python onetrust_view_export.py --dump          # print first raw response, no CSV
    python onetrust_view_export.py --endpoint /api/risk/v2/risks/pages   # alt endpoint

Read/write: this only reads (POST to a search/list endpoint). It does not
create, edit, or delete anything, and does not trigger a server-side export job.
"""

import argparse
import csv
import getpass
import json
import os
import sys
from pathlib import Path

import requests  # pip install requests

VIEWS_FILE = Path(__file__).with_name("views.json")
DEFAULT_ENDPOINT = "/api/risk/v2/risks/views/pages"
PAGE_SIZE = 100


# --------------------------------------------------------------------------- #
def load_views():
    if not VIEWS_FILE.exists():
        sys.exit(f"Missing {VIEWS_FILE}. Copy views.example.json to views.json "
                 f"and fill in your captured view filters.")
    data = json.loads(VIEWS_FILE.read_text(encoding="utf-8"))
    views = data.get("views", [])
    if not views:
        sys.exit("No views defined in views.json.")
    return views


def pick_view(views, name_arg):
    if name_arg:
        for v in views:
            if v["name"].lower() == name_arg.lower():
                return v
        sys.exit(f"View '{name_arg}' not found. Use --list.")
    print("\nAvailable views:\n")
    for i, v in enumerate(views, 1):
        print(f"  {i}. {v['name']}")
    print()
    while True:
        c = input(f"Select a view [1-{len(views)}]: ").strip()
        if c.isdigit() and 1 <= int(c) <= len(views):
            return views[int(c) - 1]
        print("Invalid selection.")


def get_token(hostname, cid, sec):
    url = f"https://{hostname}/api/access/v1/oauth/token"
    r = requests.post(url, data={"grant_type": "client_credentials",
                                 "client_id": cid, "client_secret": sec},
                      headers={"Accept": "application/json"}, timeout=60)
    if r.status_code == 401:
        sys.exit("401 - invalid client id / secret.")
    r.raise_for_status()
    tok = r.json().get("access_token")
    if not tok:
        sys.exit(f"No access_token: {r.text}")
    print("Got access token.")
    return tok


def resolve_auth(args):
    hostname = args.hostname or os.getenv("ONETRUST_HOSTNAME")
    if not hostname:
        hostname = input("OneTrust hostname (e.g. app-de.onetrust.com): ").strip()
    hostname = hostname.replace("https://", "").replace("http://", "").rstrip("/")
    if not hostname:
        sys.exit("Hostname is required.")
    token = args.token or os.getenv("ONETRUST_TOKEN")
    if not token:
        cid = args.client_id or os.getenv("ONETRUST_CLIENT_ID") \
            or input("Client id: ").strip()
        sec = args.client_secret or os.getenv("ONETRUST_CLIENT_SECRET") \
            or getpass.getpass("Client secret (hidden): ").strip()
        if not cid or not sec:
            sys.exit("Need a token, or both client id and secret.")
        token = get_token(hostname, cid, sec)
    return hostname, token


# --------------------------------------------------------------------------- #
def build_body(view):
    """Documented /risks/pages schema: filters = map, predicates = array."""
    predicates = []
    for f in view.get("filters", []):
        val = f.get("value")
        # documented predicate value is a single scalar; expand multi-value arrays
        if isinstance(val, list):
            for v in val:
                predicates.append({"field": f["field"],
                                   "operator": f.get("operator", "EQUAL_TO"),
                                   "value": v})
        else:
            predicates.append({"field": f["field"],
                               "operator": f.get("operator", "EQUAL_TO"),
                               "value": val})
    cols = [c for c in view.get("visibleColumns", []) if c]
    return {"filters": {}, "predicates": predicates, "visibleColumns": cols}


def sort_param(view):
    si = view.get("sortInfo") or {}
    col = si.get("columnName", "createdDate")
    asc = si.get("ascending", False)
    return f"{col},{'asc' if asc else 'desc'}"


def fetch_page(hostname, endpoint, headers, view, page, size):
    url = f"https://{hostname}{endpoint}"
    body = build_body(view)
    params = {"page": page, "size": size, "sort": sort_param(view)}
    r = requests.post(url, headers=headers, params=params, json=body, timeout=60)
    if r.status_code == 403:
        sys.exit(f"403 at {endpoint} - token not authorized for this endpoint.\n"
                 f"{r.text[:300]}")
    if r.status_code >= 400:
        sys.exit(f"{r.status_code} at {endpoint}: {r.text[:400]}")
    return r.json()


def rows_from(data):
    if isinstance(data, dict):
        for k in ("content", "data", "items", "results"):
            if isinstance(data.get(k), list):
                return data[k], data
    if isinstance(data, list):
        return data, {}
    return [], data


def fetch_all(hostname, endpoint, headers, view):
    all_rows, page = [], 0
    while True:
        data = fetch_page(hostname, endpoint, headers, view, page, PAGE_SIZE)
        rows, meta = rows_from(data)
        all_rows.extend(rows)
        total_pages = meta.get("totalPages")
        is_last = meta.get("last")
        print(f"  page {page + 1}"
              + (f"/{total_pages}" if total_pages else "")
              + f": {len(rows)} rows")
        if is_last is True or not rows:
            break
        if total_pages is not None and page + 1 >= total_pages:
            break
        page += 1
    return all_rows


def flatten(row, prefix=""):
    flat = {}
    for k, v in row.items():
        key = f"{prefix}{k}"
        if isinstance(v, dict):
            flat.update(flatten(v, key + "."))
        elif isinstance(v, list):
            flat[key] = "; ".join(
                json.dumps(x, ensure_ascii=False) if isinstance(x, (dict, list)) else str(x)
                for x in v)
        else:
            flat[key] = v
    return flat


def write_csv(rows, out_path):
    if not rows:
        print("No rows returned. Nothing written.")
        return
    flat = [flatten(r) for r in rows]
    cols = []
    for r in flat:
        for k in r:
            if k not in cols:
                cols.append(k)
    with open(out_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        w.writerows(flat)
    print(f"\nSaved: {out_path}  ({len(rows)} rows)")


# --------------------------------------------------------------------------- #
def main():
    ap = argparse.ArgumentParser(description="Export a OneTrust risk view to CSV.")
    ap.add_argument("--view", help="View name (skips picker).")
    ap.add_argument("--list", action="store_true", help="List views and exit.")
    ap.add_argument("--out", help="Output CSV path.")
    ap.add_argument("--endpoint", default=DEFAULT_ENDPOINT,
                    help=f"Grid endpoint (default {DEFAULT_ENDPOINT}).")
    ap.add_argument("--dump", action="store_true",
                    help="Print the first raw response and exit (no CSV).")
    ap.add_argument("--hostname")
    ap.add_argument("--token")
    ap.add_argument("--client-id")
    ap.add_argument("--client-secret")
    args = ap.parse_args()

    views = load_views()
    if args.list:
        for v in views:
            print(f"- {v['name']}")
        return

    view = pick_view(views, args.view)
    hostname, token = resolve_auth(args)
    headers = {"Authorization": f"Bearer {token}",
               "Accept": "application/json", "Content-Type": "application/json"}
    print(f"\nSelected view: {view['name']}")

    if args.dump:
        print("\n--- request body ---")
        print(json.dumps(build_body(view), ensure_ascii=False, indent=2))
        print(f"--- sort param: {sort_param(view)} ---\n")
        data = fetch_page(hostname, args.endpoint, headers, view, 0, 2)
        print("--- response ---")
        print(json.dumps(data, ensure_ascii=False, indent=2)[:4000])
        return

    rows = fetch_all(hostname, args.endpoint, headers, view)
    out = args.out or f"risk_{view['name'].replace(' ', '_').lower()}.csv"
    write_csv(rows, out)


if __name__ == "__main__":
    main()
