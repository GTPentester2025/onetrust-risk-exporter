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
import time
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
# Server-side filtering on this endpoint is unreliable (predicate field names
# differ from the UI), so we fetch all risks and filter client-side against the
# response fields. Each view filter field maps to a way of pulling the matching
# id(s) out of a risk row. Returns None if the row has no usable value for that
# field (then the filter is skipped for that row, with a warning printed once).
# Each view filter field -> how to pull the row's comparable value(s), and
# whether we compare against the filter values' UUIDs ("value") or labels
# ("label"). Level fields have a null levelGuid in the grid response, so they
# must be matched by label against the level string ("HIGH", etc).
def _level(row, key):
    lvl = (row.get(key) or {}).get("level")
    return [lvl.upper()] if lvl else []


FIELD_RULES = {
    "organization":      (lambda r: [(r.get("orgGroup") or {}).get("id")], "value"),
    "riskapprover":      (lambda r: r.get("riskApproversId") or [], "value"),
    "approvers":         (lambda r: r.get("riskApproversId") or [], "value"),
    "riskowner":         (lambda r: r.get("riskOwnersId") or [], "value"),
    "owners":            (lambda r: r.get("riskOwnersId") or [], "value"),
    "inherentriskscore": (lambda r: _level(r, "inherentRiskLevel"), "label"),
    "inherentlevel":     (lambda r: _level(r, "inherentRiskLevel"), "label"),
    "residualriskscore": (lambda r: _level(r, "residualRiskLevel"), "label"),
    "residuallevel":     (lambda r: _level(r, "residualRiskLevel"), "label"),
    "stage":             (lambda r: [r.get("state")] if r.get("state") else [], "value"),
}


def _filter_values(f):
    """Return the filter's allowed values in both id and label form."""
    ids, labels = set(), set()
    for v in f.get("value", []):
        if isinstance(v, dict):
            if v.get("value"):
                ids.add(v["value"])
            if v.get("label"):
                labels.add(v["label"].upper())
        else:
            ids.add(v)
    return ids, labels


def sort_param(view):
    si = view.get("sortInfo") or {}
    col = si.get("columnName", "createdDate")
    asc = si.get("ascending", False)
    return f"{col},{'asc' if asc else 'desc'}"


def get_filters(view):
    return view.get("activeFilters") or view.get("filters") or []


def row_matches(row, view_filters, skipped):
    """AND across filters; within a filter the row's value(s) must intersect."""
    for f in view_filters:
        field = (f.get("field") or "").lower()
        rule = FIELD_RULES.get(field)
        if not rule:
            skipped.add(f.get("field"))   # unmappable -> don't exclude
            continue
        accessor, mode = rule
        row_vals = [x for x in accessor(row) if x not in (None, "")]
        ids, labels = _filter_values(f)
        allowed = labels if mode == "label" else ids
        if not (set(row_vals) & allowed):
            return False
    return True


# --- column projection: view column name -> (CSV header, value accessor) ---- #
def _names(lst):
    if not isinstance(lst, list):
        return ""
    out = []
    for x in lst:
        if isinstance(x, dict):
            out.append(x.get("name") or x.get("label") or x.get("value") or "")
        else:
            out.append(str(x))
    return "; ".join(o for o in out if o)


COLUMN_MAP = {
    "id":                ("ID", lambda r: r.get("id")),
    "riskname":          ("Risk name", lambda r: r.get("name") or r.get("riskName")
                          or r.get("title")),
    "source":            ("Source", lambda r: (r.get("source") or {}).get("name")),
    "riskowner":         ("Risk owners", lambda r: r.get("riskOwnersName")
                          or _names(r.get("riskOwners"))),
    "riskowners":        ("Risk owners", lambda r: r.get("riskOwnersName")
                          or _names(r.get("riskOwners"))),
    "description":       ("Description", lambda r: r.get("description")),
    "treatmentplan":     ("Treatment plan", lambda r: r.get("treatment")
                          or r.get("treatmentPlan")),
    "organization":      ("Organization", lambda r: (r.get("orgGroup") or {}).get("name")),
    "stage":             ("Stage", lambda r: r.get("stage") or r.get("state")),
    "inherentriskscore": ("Inherent risk score",
                          lambda r: (r.get("inherentRiskLevel") or {}).get("level")),
    "residualriskscore": ("Residual risk score",
                          lambda r: (r.get("residualRiskLevel")
                                     or r.get("targetRiskLevel") or {}).get("level")),
    "createddate":       ("Date created", lambda r: r.get("createdUTCDateTime")
                          or r.get("createdDate")),
    "category":          ("Category", lambda r: _names(r.get("categories"))),
    "result":            ("Result", lambda r: r.get("result")),
    "dateclosed":        ("Date closed", lambda r: r.get("dateClosed")
                          or r.get("closedDate")),
    "riskapprover":      ("Risk approver", lambda r: r.get("riskApprovers")),
    "number":            ("Number", lambda r: r.get("number")),
}


def project_rows(rows, active_columns):
    """Return (list-of-ordered-dicts, headers) using only the view's columns."""
    headers, accessors = [], []
    for col in active_columns:
        header, acc = COLUMN_MAP.get(col.lower(), (col, lambda r, c=col: r.get(c)))
        headers.append(header)
        accessors.append(acc)
    out = []
    for r in rows:
        out.append({h: a(r) for h, a in zip(headers, accessors)})
    return out, headers


def rows_from(data):
    if isinstance(data, dict):
        for k in ("content", "data", "items", "results"):
            if isinstance(data.get(k), list):
                return data[k], data
    if isinstance(data, list):
        return data, {}
    return [], data


# ---- one page, single attempt, classified outcome ------------------------- #
def fetch_page(hostname, endpoint, headers, view, page, size, timeout=120):
    """One request. Returns a dict describing the outcome (never raises for
    HTTP/network issues the controller should react to)."""
    url = f"https://{hostname}{endpoint}"
    params = {"page": page, "size": size, "sort": sort_param(view)}
    try:
        r = requests.post(url, headers=headers, params=params, json={}, timeout=timeout)
    except (requests.Timeout, requests.ConnectionError) as e:
        return {"page": page, "kind": "timeout", "detail": type(e).__name__}
    if r.status_code == 429:
        ra = r.headers.get("Retry-After")
        try:
            ra = float(ra)
        except (TypeError, ValueError):
            ra = None
        return {"page": page, "kind": "429", "retry_after": ra}
    if r.status_code in (401, 403):
        return {"page": page, "kind": "fatal",
                "detail": f"{r.status_code} not authorized: {r.text[:200]}"}
    if r.status_code >= 500:
        return {"page": page, "kind": "server", "detail": f"{r.status_code}"}
    if r.status_code >= 400:
        return {"page": page, "kind": "fatal", "detail": f"{r.status_code}: {r.text[:200]}"}
    rows, meta = rows_from(r.json())
    return {"page": page, "kind": "ok", "rows": rows, "meta": meta}


def _backoff(attempt):
    """Exponential backoff with full jitter, capped at 30s."""
    import random
    return min(30.0, 2.0 ** attempt) * (0.5 + random.random() / 2)


# ---- adaptive parallel fetch (AIMD + Retry-After + backoff/jitter) --------- #
def fetch_all(hostname, endpoint, headers, view,
              start=4, cap=12, increase_every=8, max_attempts=6):
    import concurrent.futures as cf
    import heapq

    # bootstrap: page 0 tells us total pages
    first = fetch_page(hostname, endpoint, headers, view, 0, PAGE_SIZE)
    if first["kind"] == "fatal":
        sys.exit(f"{endpoint}: {first['detail']}")
    if first["kind"] != "ok":
        # transient on first call - one plain retry
        time.sleep(3)
        first = fetch_page(hostname, endpoint, headers, view, 0, PAGE_SIZE)
        if first["kind"] != "ok":
            sys.exit(f"Could not fetch page 1: {first.get('detail', first['kind'])}")
    meta = first["meta"]
    total_pages = meta.get("totalPages")
    results = {0: first["rows"]}
    if not total_pages or total_pages <= 1:
        return first["rows"]

    print(f"  total pages: {total_pages}. Fetching with adaptive concurrency...")
    MIN = 1
    target = start
    streak = 0
    attempts = {}
    ready = list(range(1, total_pages))
    scheduled = []          # min-heap of (not_before, page)
    paused_until = 0.0
    failed = []
    done_count = 1

    ex = cf.ThreadPoolExecutor(max_workers=cap)
    futures = {}
    try:
        while ready or scheduled or futures:
            now = time.monotonic()
            # move due scheduled pages back to ready
            while scheduled and scheduled[0][0] <= now:
                _, p = heapq.heappop(scheduled)
                ready.append(p)
            # submit while under target and not globally paused
            while ready and len(futures) < target and now >= paused_until:
                p = ready.pop()
                fut = ex.submit(fetch_page, hostname, endpoint, headers, view, p, PAGE_SIZE)
                futures[fut] = p
            if not futures:
                # nothing in flight; sleep until next scheduled/pause
                waits = [paused_until - now] if paused_until > now else []
                if scheduled:
                    waits.append(scheduled[0][0] - now)
                time.sleep(max(0.05, min([w for w in waits if w > 0] or [0.05])))
                continue
            done, _ = cf.wait(futures, timeout=0.5,
                              return_when=cf.FIRST_COMPLETED)
            for fut in done:
                p = futures.pop(fut)
                o = fut.result()
                kind = o["kind"]
                if kind == "ok":
                    results[p] = o["rows"]
                    done_count += 1
                    streak += 1
                    if streak >= increase_every and target < cap:
                        target += 1
                        streak = 0
                    if done_count % 10 == 0 or done_count == total_pages:
                        print(f"    {done_count}/{total_pages} pages  (concurrency={target})")
                elif kind == "429":
                    target = max(MIN, target // 2)
                    streak = 0
                    ra = o.get("retry_after")
                    delay = ra if ra else _backoff(attempts.get(p, 0) + 1)
                    paused_until = max(paused_until, time.monotonic() + (ra or 0))
                    attempts[p] = attempts.get(p, 0) + 1
                    print(f"    429 on page {p + 1} -> concurrency={target}, "
                          f"pausing {delay:.1f}s")
                    if attempts[p] > max_attempts:
                        failed.append(p)
                    else:
                        heapq.heappush(scheduled, (time.monotonic() + delay, p))
                elif kind in ("timeout", "server"):
                    # transient, NOT a rate-limit -> retry, don't cut concurrency
                    attempts[p] = attempts.get(p, 0) + 1
                    if attempts[p] > max_attempts:
                        failed.append(p)
                    else:
                        heapq.heappush(scheduled,
                                       (time.monotonic() + _backoff(attempts[p]), p))
                else:  # fatal
                    sys.exit(f"Page {p + 1}: {o.get('detail')}")
    finally:
        ex.shutdown(wait=False)

    if failed:
        print(f"*** WARNING: {len(failed)} page(s) failed after retries: "
              f"{sorted(pp + 1 for pp in failed)}. CSV will be missing those rows.")

    ordered = []
    for i in range(total_pages):
        ordered.extend(results.get(i, []))
    return ordered


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


def write_csv(rows, out_path, headers=None):
    """rows: list of flat dicts. If headers given, use that exact column order."""
    if not rows:
        print("No rows returned. Nothing written.")
        return
    if headers is None:
        headers = []
        for r in rows:
            for k in r:
                if k not in headers:
                    headers.append(k)
    with open(out_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=headers, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)
    print(f"\nSaved: {out_path}  ({len(rows)} rows, {len(headers)} columns)")


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
    ap.add_argument("--sample", action="store_true",
                    help="Write one full risk row to sample_row.json and exit.")
    ap.add_argument("--all-columns", action="store_true",
                    help="Output every field (flattened) instead of the view's columns.")
    ap.add_argument("--start-concurrency", type=int, default=4,
                    help="Initial parallel page count (AIMD floor probe).")
    ap.add_argument("--max-concurrency", type=int, default=12,
                    help="Max parallel pages (AIMD cap).")
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

    if args.dump or args.sample:
        o = fetch_page(hostname, args.endpoint, headers, view, 0, 2)
        if o["kind"] != "ok":
            print(json.dumps(o, ensure_ascii=False, indent=2)[:1000])
            return
        if args.sample:
            with open("sample_row.json", "w", encoding="utf-8") as f:
                json.dump(o["rows"][0] if o["rows"] else {}, f,
                          ensure_ascii=False, indent=2)
            print("Wrote sample_row.json (one full risk row). Paste it back.")
        else:
            print(f"totalPages={o['meta'].get('totalPages')} "
                  f"totalElements={o['meta'].get('totalElements')}")
            print(json.dumps(o["rows"][:1], ensure_ascii=False, indent=2)[:4000])
        return

    print("Fetching all risks (filtering happens locally)...")
    all_rows = fetch_all(hostname, args.endpoint, headers, view,
                         start=args.start_concurrency, cap=args.max_concurrency)

    view_filters = get_filters(view)
    skipped = set()
    matched = [r for r in all_rows if row_matches(r, view_filters, skipped)]
    print(f"\nFetched {len(all_rows)} risks; {len(matched)} match the view filters.")
    if skipped:
        print("*** WARNING: could not evaluate filter field(s): "
              + ", ".join(sorted(skipped)) + " (not applied).")

    out = args.out or f"risk_{view['name'].replace(' ', '_').lower()}.csv"
    active_columns = view.get("activeColumns")
    if active_columns and not args.all_columns:
        projected, hdrs = project_rows(matched, active_columns)
        write_csv(projected, out, headers=hdrs)
    else:
        write_csv([flatten(r) for r in matched], out)


if __name__ == "__main__":
    main()
