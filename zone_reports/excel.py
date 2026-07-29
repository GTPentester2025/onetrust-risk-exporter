"""Excel COM (pywin32) rendering for zone reports — builds every sheet FROM
SCRATCH into a brand-new workbook. Windows + Excel required.

Each zone sheet gets: a live PivotTable (Category rows x Cat columns, Count of
ID, filtered by Organization + Stage), a pie chart (risks by domain), a stacked
bar chart (Cat-wise risks by domain), a gold banner title, and the narrative.
Styling mimics the hand-built MAZ reference (gold/amber theme).

Cannot be exercised headlessly; validate on a Windows+Excel machine.
"""
import csv
import os
import subprocess
import sys

from zone_reports.stats import CAT_ORDER

DATA_SHEET = "Data"
TABLE_NAME = "RiskData"

# --- Excel enum constants (COM) ------------------------------------------- #
xlDatabase = 1
xlSrcRange = 1
xlRowField, xlColumnField, xlPageField = 1, 2, 3
xlCount = -4112
xlPie = 5
xlColumnStacked = 52
xlRows = 1
xlOpenXMLWorkbook = 51
xlHAlignCenter = -4108


# --- theme colours (Excel Interior.Color wants R + G*256 + B*65536) ------- #
def _rgb(r, g, b):
    return r + g * 256 + b * 65536


GOLD_BANNER = _rgb(255, 192, 0)     # strong amber for the title banner
GOLD_LIGHT = _rgb(255, 217, 102)    # lighter gold for pivot header
GOLD_PALE = _rgb(255, 242, 204)     # pale gold fill for narrative

# Chart anchors (points): to the right of the pivot/narrative column.
CHART_LEFT = 470
PIE_TOP, PIE_H = 10, 250
BAR_TOP, BAR_H = 270, 250
CHART_W = 430

BANNER_ROW = 13      # fallback banner row (zero-row zones / pivot build failure)
HELPER_COL = 27      # hidden helper block for chart data (column AA)


def _ensure_pywin32():
    try:
        import win32com.client  # noqa: F401
        return
    except ImportError:
        print("pywin32 not found -> installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pywin32"])
        import win32com.client  # noqa: F401


# --------------------------------------------------------------------------- #
def _write_data_sheet(wb, data_csv):
    """Rename the workbook's default sheet to Data, load refined.csv into it as a
    ListObject Table named RiskData. Pivots then reference TABLE_NAME."""
    with open(data_csv, newline="", encoding="utf-8-sig") as f:
        rows = list(csv.reader(f))
    if not rows:
        raise SystemExit("data CSV is empty.")

    ws = wb.Worksheets(1)
    ws.Name = DATA_SHEET
    nrows, ncols = len(rows), len(rows[0])
    top_left, bottom_right = ws.Cells(1, 1), ws.Cells(nrows, ncols)
    ws.Range(top_left, bottom_right).Value = tuple(tuple(r) for r in rows)

    rng = ws.Range(top_left, bottom_right)
    lo = ws.ListObjects.Add(xlSrcRange, rng, None, 1)   # 1 = xlYes (has headers)
    lo.Name = TABLE_NAME


# --------------------------------------------------------------------------- #
def _build_pivot(cache, ws, org, name):
    """Create a live PivotTable at A4: Category rows, Cat cols, Count of ID,
    Organization + Stage page filters. Returns the PivotTable."""
    pt = cache.CreatePivotTable(TableDestination=ws.Cells(4, 1), TableName=name)

    pt.PivotFields("Category").Orientation = xlRowField
    pt.PivotFields("Cat").Orientation = xlColumnField

    org_field = pt.PivotFields("Organization")
    org_field.Orientation = xlPageField
    stage_field = pt.PivotFields("Stage")
    stage_field.Orientation = xlPageField

    pt.AddDataField(pt.PivotFields("ID"), "Count of ID", xlCount)

    # Organization filter (org=None -> all items visible)
    try:
        org_field.ClearAllFilters()
        if org is not None:
            org_field.CurrentPage = org
        else:
            org_field.EnableMultiplePageItems = False
            org_field.CurrentPage = "(All)"
    except Exception as e:
        print(f"    WARN [{name}] Organization filter: {e}")
    try:
        stage_field.ClearAllFilters()
        stage_field.CurrentPage = "(All)"
    except Exception:
        pass

    # fixed Cat column order
    try:
        cf = pt.PivotFields("Cat")
        pos = 1
        for cat in CAT_ORDER:
            try:
                cf.PivotItems(cat).Position = pos
                pos += 1
            except Exception:
                continue   # cat not present in this zone
    except Exception as e:
        print(f"    WARN [{name}] Cat order: {e}")

    pt.RefreshTable()
    return pt


def _style_pivot(pt, ws):
    """Gold pivot style + bold gold header rows."""
    try:
        pt.TableStyle2 = "PivotStyleMedium3"
    except Exception:
        pass
    try:
        rng = pt.TableRange1
        rng.Font.Size = 10
        hdr = ws.Range(ws.Cells(4, 1), ws.Cells(5, rng.Columns.Count))
        hdr.Interior.Color = GOLD_LIGHT
        hdr.Font.Bold = True
    except Exception:
        pass


# --------------------------------------------------------------------------- #
def _write_banner_title(ws, title, row):
    """Big gold banner merged across the pivot width at the given row."""
    rng = ws.Range(ws.Cells(row, 1), ws.Cells(row, 9))
    rng.Merge()
    rng.Value = title
    rng.Interior.Color = GOLD_BANNER
    rng.Font.Bold = True
    rng.Font.Size = 20
    rng.HorizontalAlignment = xlHAlignCenter
    ws.Rows(row).RowHeight = 42


def _write_narrative(ws, lines, start):
    """Narrative text down column A on a pale-gold background; domain headers
    (lines containing '%') bold. Starts at row `start`."""
    for i, line in enumerate(lines):
        cell = ws.Cells(start + i, 1)
        cell.Value = line
        if line and not line.startswith("    ") and ("%" in line or line.startswith("Key Insights")):
            cell.Font.Bold = True
    block = ws.Range(ws.Cells(start, 1), ws.Cells(start + max(len(lines), 1), 6))
    block.Interior.Color = GOLD_PALE


# --------------------------------------------------------------------------- #
def _write_helper(ws, stats):
    """Write chart source into hidden helper columns. Returns (pie_range,
    bar_matrix_range) as COM Range objects."""
    c0 = HELPER_COL
    ws.Cells(1, c0).Value = "Domain"
    ws.Cells(1, c0 + 1).Value = "Count"
    for i, (dom, n) in enumerate(stats.by_domain, start=2):
        ws.Cells(i, c0).Value = dom
        ws.Cells(i, c0 + 1).Value = n
    pie_rng = ws.Range(ws.Cells(1, c0), ws.Cells(1 + len(stats.by_domain), c0 + 1))

    cats = [c for c, _ in stats.grand_by_cat]
    m0 = c0 + 3
    ws.Cells(1, m0).Value = ""                       # corner
    for j, cat in enumerate(cats, start=1):
        ws.Cells(1, m0 + j).Value = cat
    for i, (dom, _) in enumerate(stats.by_domain, start=2):
        ws.Cells(i, m0).Value = dom
        row = stats.crosstab.get(dom, {})
        for j, cat in enumerate(cats, start=1):
            ws.Cells(i, m0 + j).Value = row.get(cat, 0)
    bar_rng = ws.Range(ws.Cells(1, m0),
                       ws.Cells(1 + len(stats.by_domain), m0 + len(cats)))

    try:
        ws.Range(ws.Columns(c0), ws.Columns(m0 + len(cats))).EntireColumn.Hidden = True
    except Exception:
        pass
    return pie_rng, bar_rng


def _add_charts(ws, pie_rng, bar_rng):
    pie = ws.ChartObjects().Add(CHART_LEFT, PIE_TOP, CHART_W, PIE_H).Chart
    pie.ChartType = xlPie
    pie.SetSourceData(Source=pie_rng)
    pie.HasTitle = True
    pie.ChartTitle.Text = "RISKS IDENTIFIED"
    try:
        pie.ApplyDataLabels()
    except Exception:
        pass

    bar = ws.ChartObjects().Add(CHART_LEFT, BAR_TOP, CHART_W, BAR_H).Chart
    bar.ChartType = xlColumnStacked
    bar.SetSourceData(Source=bar_rng, PlotBy=xlRows)   # each domain a stacked series
    bar.HasTitle = True
    bar.ChartTitle.Text = "Category wise Risks"


# --------------------------------------------------------------------------- #
def render_workbook(out_path, data_csv, reports):
    """Create a NEW workbook: Data sheet + one sheet per report, each built from
    scratch. `reports` are the dicts from iter_zone_reports."""
    _ensure_pywin32()
    import win32com.client as win32

    excel = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Add()
        while wb.Worksheets.Count > 1:                # trim default extra sheets
            wb.Worksheets(wb.Worksheets.Count).Delete()

        _write_data_sheet(wb, data_csv)
        cache = wb.PivotCaches().Create(SourceType=xlDatabase, SourceData=TABLE_NAME)

        for rep in reports:
            ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
            ws.Name = rep["sheet"]
            pt_name = "pt_" + rep["sheet"].replace("-", "_")
            total = rep["stats"].total

            # Banner sits below the pivot body; default row when there's no
            # pivot (zero-row zone) or if the pivot build fails.
            banner_row = BANNER_ROW
            if total > 0:
                try:
                    pt = _build_pivot(cache, ws, rep["org"], pt_name)
                    _style_pivot(pt, ws)
                    tr = pt.TableRange1
                    banner_row = tr.Row + tr.Rows.Count + 1   # clear of the pivot
                except Exception as e:
                    print(f"  WARN: pivot build failed on {rep['sheet']}: {e}")

            # Title + narrative never write into the pivot body now; still guard
            # so one sheet's failure can't abort the whole SaveAs.
            try:
                _write_banner_title(ws, rep["title"], banner_row)
                _write_narrative(ws, rep["narrative"], banner_row + 2)
            except Exception as e:
                print(f"  WARN: title/narrative failed on {rep['sheet']}: {e}")

            if total > 0:
                try:
                    pie_rng, bar_rng = _write_helper(ws, rep["stats"])
                    _add_charts(ws, pie_rng, bar_rng)
                except Exception as e:
                    print(f"  WARN: charts failed on {rep['sheet']}: {e}")
            print(f"  built {rep['sheet']}: total={total}")

        wb.Worksheets(DATA_SHEET).Move(After=wb.Worksheets(wb.Worksheets.Count))
        wb.SaveAs(os.path.abspath(out_path), FileFormat=xlOpenXMLWorkbook)
        print(f"Saved: {out_path}")
    finally:
        if excel is not None:
            excel.Quit()
