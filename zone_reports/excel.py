# zone_reports/excel.py
"""Excel COM (pywin32) rendering for zone reports. Windows + Excel required."""
import os
import shutil
import subprocess
import sys

from zone_reports.stats import CAT_ORDER

DATA_SHEET = "Data"
TABLE_NAME = "RiskData"


def _ensure_pywin32():
    try:
        import win32com.client  # noqa: F401
        return
    except ImportError:
        print("pywin32 not found -> installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pywin32"])
        import win32com.client  # noqa: F401


def _write_data_sheet(wb, data_csv):
    """Load refined.csv into a fresh Data sheet as a ListObject Table; return the
    table's range address (e.g. 'Data!$A$1:$S$4491')."""
    import csv
    with open(data_csv, newline="", encoding="utf-8-sig") as f:
        rows = list(csv.reader(f))
    if not rows:
        raise SystemExit("refined.csv is empty.")

    # remove an existing Data sheet, then recreate
    for ws in list(wb.Worksheets):
        if ws.Name == DATA_SHEET:
            ws.Delete()
    ws = wb.Worksheets.Add()
    ws.Name = DATA_SHEET

    nrows, ncols = len(rows), len(rows[0])
    # bulk write via a 2D array assignment (fast)
    top_left = ws.Cells(1, 1)
    bottom_right = ws.Cells(nrows, ncols)
    ws.Range(top_left, bottom_right).Value = rows

    # define the Table (ListObject) over the used range
    xlSrcRange = 1
    rng = ws.Range(top_left, bottom_right)
    lo = ws.ListObjects.Add(xlSrcRange, rng, None, 1)  # 1 = xlYes (has headers)
    lo.Name = TABLE_NAME
    return f"{DATA_SHEET}!{rng.Address}"


def _find_pivot(ws):
    """Return the first PivotTable on a sheet, or None."""
    try:
        if ws.PivotTables().Count >= 1:
            return ws.PivotTables(1)
    except Exception:
        pass
    return None


def _repoint_pivot_source(wb, pt, table_range):
    """Point a pivot's cache at the RiskData table range and refresh."""
    xlDatabase = 1
    cache = wb.PivotCaches().Create(SourceType=xlDatabase, SourceData=table_range)
    pt.ChangePivotCache(cache)
    pt.RefreshTable()


def _configure_pivot(pt, org):
    """Set the Organization page-filter (org=None -> All), Stage -> All, and
    enforce the fixed Cat column order. Field names come from the template
    pivot, not hard-coded captions."""
    # Organization page filter
    try:
        f = pt.PivotFields("Organization")
        f.ClearAllFilters()
        if org is None:
            f.CurrentPage = "(All)"
        else:
            f.CurrentPage = org
    except Exception as e:
        print(f"  WARN: Organization filter: {e}")

    # Stage -> all
    try:
        s = pt.PivotFields("Stage")
        s.ClearAllFilters()
        s.CurrentPage = "(All)"
    except Exception:
        pass

    # Fixed Cat order
    try:
        cf = pt.PivotFields("Cat")
        pos = 1
        for cat in CAT_ORDER:
            try:
                cf.PivotItems(cat).Position = pos
                pos += 1
            except Exception:
                continue  # cat not present in this zone
    except Exception as e:
        print(f"  WARN: Cat order: {e}")

    pt.RefreshTable()


def _find_title_cell(ws):
    """Find the cell whose text contains 'Risk Analysis' (the MAZ title)."""
    try:
        found = ws.UsedRange.Find("Risk Analysis")
        if found is not None:
            return found.Row, found.Column
    except Exception:
        pass
    return None


def _write_title_and_narrative(ws, title_rc, title, lines):
    """Overwrite the title cell and the narrative block below it."""
    if title_rc is None:
        print("  WARN: no title cell found; skipping title/narrative for this sheet.")
        return
    r, c = title_rc
    ws.Cells(r, c).Value = title
    # narrative goes into the first column of the block, a couple rows below
    start = r + 1
    for i, line in enumerate(lines):
        ws.Cells(start + i, c).Value = line


def render_workbook(workbook_path, out_path, template_sheet, data_csv, reports):
    _ensure_pywin32()
    import win32com.client as win32

    backup = os.path.splitext(workbook_path)[0] + "_backup.xlsx"
    shutil.copyfile(workbook_path, backup)
    print(f"Backup written: {backup}")

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(os.path.abspath(workbook_path))

        if template_sheet not in [ws.Name for ws in wb.Worksheets]:
            raise SystemExit(f"Template sheet '{template_sheet}' not found in workbook.")

        table_range = _write_data_sheet(wb, data_csv)

        tmpl = wb.Worksheets(template_sheet)

        for rep in reports:
            sheet = rep["sheet"]
            # start each zone from a clean clone of the template
            if sheet != template_sheet and sheet in [ws.Name for ws in wb.Worksheets]:
                wb.Worksheets(sheet).Delete()
            if sheet == template_sheet:
                ws = tmpl
            else:
                tmpl.Copy(After=wb.Worksheets(wb.Worksheets.Count))
                ws = wb.Worksheets(wb.Worksheets.Count)
                ws.Name = sheet

            pt = _find_pivot(ws)
            if pt is not None:
                _repoint_pivot_source(wb, pt, table_range)
                _configure_pivot(pt, rep["org"])
            else:
                print(f"  WARN: no pivot on sheet {sheet}.")

            _write_title_and_narrative(ws, _find_title_cell(ws),
                                       rep["title"], rep["narrative"])
            print(f"  built {sheet}: total={rep['stats'].total}")

        wb.RefreshAll()
        excel.CalculateUntilAsyncQueriesDone()
        wb.SaveAs(os.path.abspath(out_path))
        print(f"Saved: {out_path}")
    finally:
        excel.Quit()
