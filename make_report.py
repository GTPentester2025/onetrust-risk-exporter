#!/usr/bin/env python3
"""
make_report.py -- one command: OneTrust API -> refined data -> full Excel report.

Chains the three stages:
  1. onetrust_view_export.py  --all-columns   -> raw_export.csv   (API pull)
  2. refine_columns.py                         -> refined.csv      (+ Cat lookup)
  3. zone_reports (build_zone_reports)         -> <out>.xlsx       (zone sheets)

Every zone sheet (GHQ, AFR, SAZ, MAZ, NAZ, APAC, EUR, BEES, BEES-FT) plus an
Overall sheet is built from scratch with a live PivotTable, pie + stacked-bar
charts, gold banner title, and narrative.

Auth for stage 1 flows through the exporter: pass --hostname / --client-id /
--client-secret, or set ONETRUST_HOSTNAME / ONETRUST_CLIENT_ID /
ONETRUST_CLIENT_SECRET / ONETRUST_TOKEN in the environment.

Usage:
    python make_report.py --view "TPRM Global View" --out Risk_Reports.xlsx
    python make_report.py --view "TPRM Global View" --cat-file Supplier_Category_List.xlsx
    python make_report.py --skip-export --data refined.csv        # reuse existing data
"""
import argparse
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent


def run(cmd):
    print(f"\n>>> {' '.join(cmd)}")
    r = subprocess.run(cmd)
    if r.returncode != 0:
        sys.exit(f"Step failed (exit {r.returncode}): {' '.join(cmd)}")


def main():
    ap = argparse.ArgumentParser(description="API -> refined -> full Excel report.")
    # stage 1 (export) pass-through
    ap.add_argument("--view", help="View name to export (stage 1).")
    ap.add_argument("--hostname")
    ap.add_argument("--client-id")
    ap.add_argument("--client-secret")
    ap.add_argument("--token")
    ap.add_argument("--raw", default="raw_export.csv", help="Raw export path (stage 1 output).")
    # stage 2 (refine)
    ap.add_argument("--cat-file", default="Supplier_Category_List.xlsx",
                    help="Supplier->Cat lookup for stage 2.")
    ap.add_argument("--data", default="refined.csv", help="Refined CSV (stage 2 output / stage 3 input).")
    # stage 3 (report)
    ap.add_argument("--out", default="Risk_Reports.xlsx", help="Final workbook.")
    # skips for reruns
    ap.add_argument("--skip-export", action="store_true", help="Reuse existing --raw; skip stage 1.")
    ap.add_argument("--skip-refine", action="store_true", help="Reuse existing --data; skip stages 1-2.")
    args = ap.parse_args()

    py = sys.executable

    # --- stage 1: export ---------------------------------------------------
    if not (args.skip_export or args.skip_refine):
        if not args.view:
            sys.exit("--view is required (or use --skip-export / --skip-refine).")
        cmd = [py, str(HERE / "onetrust_view_export.py"),
               "--view", args.view, "--all-columns", "--out", args.raw]
        for flag, val in (("--hostname", args.hostname), ("--client-id", args.client_id),
                          ("--client-secret", args.client_secret), ("--token", args.token)):
            if val:
                cmd += [flag, val]
        run(cmd)

    # --- stage 2: refine ---------------------------------------------------
    if not args.skip_refine:
        cmd = [py, str(HERE / "refine_columns.py"),
               "--in", args.raw, "--out", args.data, "--cat-file", args.cat_file]
        run(cmd)

    # --- stage 3: build the workbook --------------------------------------
    run([py, str(HERE / "build_zone_reports.py"), "--data", args.data, "--out", args.out])

    print(f"\nDone. Report: {args.out}")


if __name__ == "__main__":
    main()
