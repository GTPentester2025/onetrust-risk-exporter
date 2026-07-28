#!/usr/bin/env python3
"""
v3 test: does POST /api/risk/v2/risks/views/pages filter server-side by viewId?

If passing a saved view's id narrows the result set to that view's rows, we get
exact parity (right rows + can use the view's columns) with no 25k client pull.
This posts several viewId shapes (body + query) and prints the TOTAL count for
each vs an unfiltered baseline. A shape whose TOTAL is much smaller = it filters.

Pass the view id explicitly (no internal ids are baked into this file):
    python viewid_test_v3.py --view-id <uuid>
    (env: ONETRUST_HOSTNAME / ONETRUST_CLIENT_ID / ONETRUST_CLIENT_SECRET)

Read-in-effect: only POSTs to the /pages search endpoint with an id-only body,
plus a baseline. No writes, no /reports/export, no PUT/PATCH/DELETE.
"""

import argparse
import getpass
import json
import os
import sys

import requests  # pip install requests

ENDPOINT = "/api/risk/v2/risks/views/pages"
OUTFILE = "v3_results.txt"

# body shapes to try (value {id} is substituted)
BODY_SHAPES = [
    ("baseline (no view)", {}),
    ("viewId", {"viewId": "{id}"}),
    ("savedViewId", {"savedViewId": "{id}"}),
    ("riskViewId", {"riskViewId": "{id}"}),
    ("view", {"view": "{id}"}),
    ("viewInfo.id", {"viewInfo": {"id": "{id}"}}),
    ("id", {"id": "{id}"}),
]
QUERY_SHAPES = [("?viewId=", "viewId"), ("?savedViewId=", "savedViewId"),
                ("?riskViewId=", "riskViewId")]


class Tee:
    def __init__(self, path):
        self._c = sys.__stdout__
        self._f = open(path, "w", encoding="utf-8")

    def write(self, s):
        self._c.write(s); self._f.write(s); self._f.flush()

    def flush(self):
        self._c.flush(); self._f.flush()


def get_token(hostname, cid, sec):
    url = f"https://{hostname}/api/access/v1/oauth/token"
    r = requests.post(url, data={"grant_type": "client_credentials",
                                 "client_id": cid, "client_secret": sec},
                      headers={"Accept": "application/json"}, timeout=60)
    if r.status_code != 200:
        sys.exit(f"Token request failed ({r.status_code}): {r.text}")
    t = r.json().get("access_token")
    if not t:
        sys.exit(f"No access_token: {r.text}")
    print("Got access token.\n")
    return t


def summarize(resp):
    try:
        d = resp.json()
    except ValueError:
        return f"status={resp.status_code} (non-JSON) {resp.text[:150]}"
    if not isinstance(d, dict):
        return f"status={resp.status_code} type={type(d).__name__}"
    total = next((d[k] for k in ("totalElements", "totalCount", "total") if k in d), None)
    content = d.get("content")
    n = len(content) if isinstance(content, list) else "?"
    # peek at first row's org for a sanity signal
    org = ""
    if isinstance(content, list) and content and isinstance(content[0], dict):
        og = content[0].get("orgGroup") or {}
        org = f" firstOrg={og.get('name')}"
    return f"status={resp.status_code} TOTAL={total} rows_in_page={n}{org}"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--view-id", required=True, help="Saved view UUID to test.")
    args = ap.parse_args()
    vid = args.view_id

    hostname = os.getenv("ONETRUST_HOSTNAME")
    if not hostname:
        hostname = input("OneTrust hostname (e.g. app-de.onetrust.com): ").strip()
    hostname = hostname.replace("https://", "").replace("http://", "").rstrip("/")
    if not hostname:
        sys.exit("Hostname is required.")

    token = os.getenv("ONETRUST_TOKEN")
    if not token:
        cid = os.getenv("ONETRUST_CLIENT_ID") or input("Client id: ").strip()
        sec = os.getenv("ONETRUST_CLIENT_SECRET") or \
            getpass.getpass("Client secret (hidden): ").strip()
        if not cid or not sec:
            sys.exit("Need token, or both client id and secret.")
        token = get_token(hostname, cid, sec)

    headers = {"Authorization": f"Bearer {token}",
               "Accept": "application/json", "Content-Type": "application/json"}
    base = f"https://{hostname}{ENDPOINT}"

    sys.stdout = Tee(OUTFILE)
    print("=" * 72)
    print("V3 VIEWID TEST  (paste v3_results.txt)")
    print("=" * 72)
    print(f"Endpoint: POST {ENDPOINT}")
    print("A shape with TOTAL far below baseline = server-side viewId filtering works.\n")

    print("### BODY shapes ###")
    for label, shape in BODY_SHAPES:
        body = json.loads(json.dumps(shape).replace("{id}", vid))
        try:
            r = requests.post(base, headers=headers, params={"size": 1}, json=body, timeout=40)
            print(f"\n[{label}] body={json.dumps(body)}")
            print(f"      {summarize(r)}")
        except requests.RequestException as e:
            print(f"\n[{label}] ERROR {e}")

    print("\n### QUERY shapes ###")
    for label, param in QUERY_SHAPES:
        try:
            r = requests.post(base, headers=headers,
                              params={param: vid, "size": 1}, json={}, timeout=40)
            print(f"\n[{label}{vid[:8]}...]")
            print(f"      {summarize(r)}")
        except requests.RequestException as e:
            print(f"\n[{label}] ERROR {e}")

    print("\n" + "=" * 72)
    print("END V3 VIEWID TEST")
    print("=" * 72)
    sys.__stdout__.write(f"\n[written to {OUTFILE}]\n")


if __name__ == "__main__":
    main()
