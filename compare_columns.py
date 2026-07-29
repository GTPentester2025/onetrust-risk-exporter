#!/usr/bin/env python3
"""
compare_columns.py -- figure out which CLI-export columns match the GUI export.

You have two files for the SAME risks:
  * the GUI export (few columns, the ones you actually want)
  * the CLI/script export (~200 flattened columns, the raw grid)

This script joins the two on a shared key (the risk GUID by default, auto-
detected if not given), then for every GUI column finds the CLI column whose
values line up best. Output: a mapping you can paste into COLUMN_MAP /
activeColumns, plus a list of GUI columns with no good match and sample
mismatches so you can eyeball why.

It only READS the two files. Nothing is written unless you pass --out.

Usage:
    python compare_columns.py --gui gui.csv --cli cli.csv
    python compare_columns.py --gui gui.csv --cli cli.csv --threshold 0.95
    python compare_columns.py --gui gui.xlsx --cli cli.csv --out mapping.csv
    python compare_columns.py --gui gui.csv --cli cli.csv \
        --key-gui "Risk GUID" --key-cli id
"""

import argparse
import csv
import re
import sys
from pathlib import Path


# --------------------------------------------------------------------------- #
def load_table(path):
    """Return (headers, list-of-row-dicts). Supports .csv and .xlsx."""
    p = Path(path)
    if not p.exists():
        sys.exit(f"Not found: {path}")
    if p.suffix.lower() in (".xlsx", ".xlsm"):
        return _load_xlsx(p)
    # try utf-8-sig first (Excel/our exporter write a BOM), fall back to latin-1
    for enc in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            with open(p, newline="", encoding=enc) as f:
                r = csv.DictReader(f)
                rows = list(r)
                return (r.fieldnames or []), rows
        except UnicodeDecodeError:
            continue
    sys.exit(f"Could not decode {path} as CSV.")


def _load_xlsx(p):
    try:
        from openpyxl import load_workbook          # pip install openpyxl
    except ImportError:
        sys.exit("Reading .xlsx needs openpyxl (pip install openpyxl), "
                 "or just save the sheet as CSV and pass that.")
    wb = load_workbook(p, read_only=True, data_only=True)
    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)
    headers = [str(h) if h is not None else "" for h in next(rows_iter)]
    rows = []
    for raw in rows_iter:
        rows.append({headers[i]: ("" if v is None else v)
                     for i, v in enumerate(raw) if i < len(headers)})
    return headers, rows


# --------------------------------------------------------------------------- #
_ws = re.compile(r"\s+")
_guid = re.compile(r"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-"
                   r"[0-9a-f]{4}-[0-9a-f]{12}$", re.I)


def norm(v):
    """Normalize a cell for comparison: str, trimmed, collapsed ws, lowercase.
    Dates like '2026-07-29T00:00:00Z' and '2026-07-29' compare equal on prefix
    handled by norm_date below; here we just canonicalize text."""
    if v is None:
        return ""
    s = str(v).strip()
    s = _ws.sub(" ", s)
    return s.lower()


_DATE_FMTS = ("%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d",
              "%m/%d/%Y", "%d/%m/%Y", "%m/%d/%y", "%d-%m-%Y",
              "%b %d, %Y", "%B %d, %Y", "%d %b %Y", "%d %B %Y",
              "%d-%b-%Y", "%d-%b-%y", "%Y/%m/%d")


def date_candidates(s):
    """Parse a date/datetime string into the SET of ISO dates it could mean.
    Ambiguous slash dates (03/06/2026) yield both {2026-03-06, 2026-06-03} so a
    DD/MM export still matches an MM/DD or ISO counterpart. Empty set if not a
    date."""
    from datetime import datetime
    s = s.strip()
    if not s:
        return frozenset()
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})[t ]", s)      # ISO datetime prefix
    if m:
        return frozenset({f"{m.group(1)}-{m.group(2)}-{m.group(3)}"})
    core = s.rstrip("zZ").strip()
    core = re.sub(r"[.+]\d+(:\d+)?$", "", core)           # drop ms / tz tail
    out = set()
    for fmt in _DATE_FMTS:
        try:
            out.add(datetime.strptime(core, fmt).strftime("%Y-%m-%d"))
        except ValueError:
            continue
    return frozenset(out)


def to_date(s):
    """Single canonical YYYY-MM-DD (first candidate) -- used for join keys."""
    c = date_candidates(s)
    return min(c) if c else None


def to_float(s):
    """Return float if s is a plain number (allows commas/%), else None."""
    t = s.strip().rstrip("%").replace(",", "")
    if not re.match(r"^[+-]?\d+(\.\d+)?$", t):
        return None
    try:
        return float(t)
    except ValueError:
        return None


_split = re.compile(r"\s*[;,/|]\s*")


def split_multi(s):
    """Split a multi-value cell into a normalized set (owners, categories)."""
    return {p for p in (x.strip() for x in _split.split(s)) if p}


def norm_date(s):
    d = to_date(s)
    return d if d else s


def cmp_key(v):
    """Canonical string used for GUID/key joins (dates collapsed to date part)."""
    return norm_date(norm(v))


_maybe_date = re.compile(r"\d.*[-/]|\b\w{3,9}\s+\d")   # cheap pre-filter
_TOKEN_CACHE = {}


def typed(v):
    """Parse a cell ONCE into a comparable token, memoized on the raw string.
    Token is (kind, payload):
      ("", )                     -> empty (skip)
      ("num", float)
      ("date", "YYYY-MM-DD")
      ("str", normstr, frozenset(parts))
    """
    key = v if isinstance(v, str) else str(v)
    t = _TOKEN_CACHE.get(key)
    if t is not None:
        return t
    s = norm(v)
    if s == "":
        t = ("",)
    else:
        f = to_float(s)
        if f is not None:
            t = ("num", f)
        else:
            ds = date_candidates(s) if _maybe_date.search(s) else frozenset()
            if ds:
                t = ("date", ds)
            else:
                t = ("str", s, frozenset(split_multi(s)))
    _TOKEN_CACHE[key] = t
    return t


def match_tok(g, c):
    """Equivalence score in [0,1] between two pre-parsed tokens. Assumes the
    GUI token is non-empty (caller skips empties)."""
    if c[0] == "":
        return 0.0
    if g[0] == c[0]:
        k = g[0]
        if k == "num":
            return 1.0 if g[1] == c[1] else 0.0
        if k == "date":
            return 1.0 if (g[1] & c[1]) else 0.0
        # both str
        gs, cs = g[1], c[1]
        if gs == cs:
            return 1.0
        gp, cp = g[2], c[2]
        if len(gp) > 1 or len(cp) > 1:
            inter = gp & cp
            return len(inter) / max(len(gp), len(cp)) if inter else 0.0
        if len(gs) >= 4 and len(cs) >= 4 and (gs in cs or cs in gs):
            return 0.5
        return 0.0
    # cross-kind: compare on their string forms as a weak fallback
    gs = g[1] if g[0] == "str" else str(g[1])
    cs = c[1] if c[0] == "str" else str(c[1])
    if gs == cs:
        return 1.0
    if len(gs) >= 4 and len(cs) >= 4 and (gs in cs or cs in gs):
        return 0.5
    return 0.0


def values_match(gv, cv):
    """Convenience wrapper used for sample display."""
    g = typed(gv)
    return 0.0 if g[0] == "" else match_tok(g, typed(cv))


def looks_guid_col(values):
    hits = sum(1 for v in values if _guid.match(str(v).strip()))
    return values and hits / len(values) > 0.8


# --------------------------------------------------------------------------- #
def detect_key(gui_hdr, gui_rows, cli_hdr, cli_rows, key_gui, key_cli):
    """Return (key_gui, key_cli). If not supplied, find a GUI column and a CLI
    column that both hold GUIDs and whose value sets overlap the most."""
    if key_gui and key_cli:
        return key_gui, key_cli

    def guid_cols(hdr, rows):
        out = []
        for h in hdr:
            vals = [r.get(h, "") for r in rows]
            if looks_guid_col(vals):
                out.append(h)
        return out

    g_cands = guid_cols(gui_hdr, gui_rows) or gui_hdr
    c_cands = guid_cols(cli_hdr, cli_rows) or cli_hdr

    best, best_ov = None, -1
    for g in g_cands:
        gset = {cmp_key(r.get(g, "")) for r in gui_rows} - {""}
        if not gset:
            continue
        for c in c_cands:
            cset = {cmp_key(r.get(c, "")) for r in cli_rows} - {""}
            ov = len(gset & cset)
            if ov > best_ov:
                best, best_ov = (g, c), ov
    if not best or best_ov <= 0:
        sys.exit("Could not auto-detect a shared key. Pass --key-gui and "
                 "--key-cli explicitly (a column present in both files, e.g. "
                 "the risk GUID or number).")
    return best


def index_by_key(rows, key):
    idx = {}
    dupes = 0
    for r in rows:
        k = cmp_key(r.get(key, ""))
        if not k:
            continue
        if k in idx:
            dupes += 1
            continue
        idx[k] = r
    return idx, dupes


# --------------------------------------------------------------------------- #
def compare(gui_hdr, gui_rows, cli_hdr, cli_rows, key_gui, key_cli, threshold):
    cli_idx, cli_dupes = index_by_key(cli_rows, key_cli)
    joined = []               # (gui_row, cli_row)
    unmatched = 0
    for gr in gui_rows:
        k = cmp_key(gr.get(key_gui, ""))
        cr = cli_idx.get(k)
        if cr is None:
            unmatched += 1
            continue
        joined.append((gr, cr))

    print(f"Key: GUI '{key_gui}'  <->  CLI '{key_cli}'")
    print(f"GUI rows: {len(gui_rows)} | CLI rows: {len(cli_rows)} | "
          f"joined: {len(joined)} | GUI rows with no CLI match: {unmatched}"
          + (f" | CLI dup keys skipped: {cli_dupes}" if cli_dupes else ""))
    if not joined:
        sys.exit("No rows joined - key values don't overlap. Check --key-*.")
    print()

    # Parse every cell ONCE into a typed token. CLI column vectors are built a
    # single time and reused across all GUI columns (the expensive part).
    gui_j = [gr for gr, _ in joined]
    cli_j = [cr for _, cr in joined]
    cli_typed = {c: [typed(cr.get(c, "")) for cr in cli_j] for c in cli_hdr}

    results = []              # (gui_col, best_cli_col, score, n, samples)
    for gcol in gui_hdr:
        gvec = [typed(gr.get(gcol, "")) for gr in gui_j]
        idxs = [i for i, g in enumerate(gvec) if g[0] != ""]   # non-empty GUI
        n = len(idxs)
        best_col, best_score, best_n = None, -1.0, n
        ranked = []
        for ccol in cli_hdr:
            cvec = cli_typed[ccol]
            if n == 0:
                sc = 0.0
            else:
                hit = 0.0
                for i in idxs:
                    hit += match_tok(gvec[i], cvec[i])
                sc = hit / n
            ranked.append((sc, n, ccol))
            if sc > best_score:
                best_col, best_score, best_n = ccol, sc, n
        ranked.sort(reverse=True)
        samples = []
        if best_col is not None and best_score < 1.0:
            for gr, cr in joined:
                gv, cv = gr.get(gcol, ""), cr.get(best_col, "")
                if norm(gv) != "" and values_match(gv, cv) < 1.0:
                    samples.append((str(gv), str(cv)))
                if len(samples) >= 3:
                    break
        runners = [f"{c}({s:.2f})" for s, n, c in ranked[1:4] if s > 0]
        results.append((gcol, best_col, best_score, best_n, samples, runners))

    # ---- report ---------------------------------------------------------- #
    print(f"{'GUI column':28} {'best CLI column':34} {'score':>6} {'n':>5}")
    print("-" * 80)
    good, weak = [], []
    for gcol, ccol, sc, n, samples, runners in results:
        flag = "OK " if sc >= threshold else ("~~ " if sc >= 0.5 else "XX ")
        print(f"{flag}{gcol[:27]:27} {str(ccol)[:33]:33} {sc:6.2f} {n:5d}")
        if runners:
            print(f"      alt: {', '.join(runners)}")
        for gv, cv in samples:
            print(f"      mismatch  GUI={gv[:34]!r}  CLI={cv[:34]!r}")
        (good if sc >= threshold else weak).append((gcol, ccol, sc))

    print("\n" + "=" * 80)
    print(f"Matched >= {threshold:.2f}: {len(good)} / {len(gui_hdr)} GUI columns")
    if weak:
        print("Needs attention (score below threshold):")
        for gcol, ccol, sc in weak:
            print(f"  - {gcol}: best guess '{ccol}' @ {sc:.2f}")

    # suggested activeColumns / COLUMN_MAP snippet from the confident matches
    print("\nSuggested CLI columns to keep (best match per GUI column):")
    for gcol, ccol, sc in [(g, c, s) for g, c, s, *_ in
                           [(r[0], r[1], r[2]) for r in results]]:
        mark = "" if sc >= threshold else "   # <-- LOW, verify"
        print(f"  {gcol!r:30} -> {str(ccol)!r}{mark}")

    return results, joined


def write_out(path, results):
    rows = [{"gui_column": g, "best_cli_column": c, "score": f"{s:.3f}",
             "n_compared": n, "runners_up": " | ".join(runners)}
            for g, c, s, n, _samples, runners in results]
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["gui_column", "best_cli_column",
                                          "score", "n_compared", "runners_up"])
        w.writeheader()
        w.writerows(rows)
    print(f"\nWrote mapping: {path}")


# --------------------------------------------------------------------------- #
def main():
    ap = argparse.ArgumentParser(
        description="Match GUI-export columns to CLI-export columns by value.")
    ap.add_argument("--gui", required=True, help="GUI export (.csv or .xlsx).")
    ap.add_argument("--cli", required=True, help="CLI/script export (.csv or .xlsx).")
    ap.add_argument("--key-gui", help="Join column in GUI file (auto if omitted).")
    ap.add_argument("--key-cli", help="Join column in CLI file (auto if omitted).")
    ap.add_argument("--threshold", type=float, default=0.9,
                    help="Match ratio to count a column as confidently mapped.")
    ap.add_argument("--out", help="Write the mapping to this CSV.")
    args = ap.parse_args()

    gui_hdr, gui_rows = load_table(args.gui)
    cli_hdr, cli_rows = load_table(args.cli)
    if not gui_rows:
        sys.exit("GUI file has no data rows.")
    if not cli_rows:
        sys.exit("CLI file has no data rows.")

    key_gui, key_cli = detect_key(gui_hdr, gui_rows, cli_hdr, cli_rows,
                                  args.key_gui, args.key_cli)
    results, _ = compare(gui_hdr, gui_rows, cli_hdr, cli_rows,
                         key_gui, key_cli, args.threshold)
    if args.out:
        write_out(args.out, results)


if __name__ == "__main__":
    main()
