#!/usr/bin/env python3
"""Build per-zone + Overall Excel risk reports from refined.csv.

Pure computation and --dry-run need no Excel. A full run drives Excel via COM
(pywin32, auto-installed) to clone the MAZ template per zone. See
docs/superpowers/specs/2026-07-29-zone-risk-reports-design.md.
"""
import argparse
import sys
from datetime import date

from zone_reports.stats import load_rows, compute_zone_stats, ZONES
from zone_reports.narrative import build_title, build_narrative, format_date_ordinal


def zone_display_name(sheet, org):
    return "Overall" if org is None else org


def iter_zone_reports(rows, date_str):
    reports = []
    for sheet, org in ZONES.items():
        stats = compute_zone_stats(rows, org)
        name = zone_display_name(sheet, org)
        reports.append({
            "sheet": sheet, "org": org, "zone_name": name, "stats": stats,
            "title": build_title(name, date_str),
            "narrative": build_narrative(stats, date_str),
        })
    return reports


def _print_dry_run(reports):
    for r in reports:
        print("=" * 70)
        print(r["title"])
        print(f"  sheet={r['sheet']}  org={r['org']}  total={r['stats'].total}")
        print("  by Cat:", r["stats"].by_cat)
        print("  " + "\n  ".join(r["narrative"]))


def main(argv=None):
    ap = argparse.ArgumentParser(description="Build zone risk report sheets.")
    ap.add_argument("--data", default="refined.csv")
    ap.add_argument("--out", default="Risk_Reports.xlsx",
                    help="Output workbook to CREATE (default: Risk_Reports.xlsx).")
    ap.add_argument("--dry-run", action="store_true",
                    help="Compute + print stats/narrative for every zone; no Excel.")
    args = ap.parse_args(argv)

    rows = load_rows(args.data)
    date_str = format_date_ordinal(date.today())
    reports = iter_zone_reports(rows, date_str)

    if args.dry_run:
        _print_dry_run(reports)
        return 0

    from zone_reports.excel import render_workbook
    render_workbook(args.out, args.data, reports)       # builds a new workbook
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
