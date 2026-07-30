"""Excel COM (pywin32) rendering for zone reports — builds every sheet FROM
SCRATCH into a brand-new workbook. Windows + Excel required.

Each zone sheet gets: a live PivotTable (Category rows x Cat columns, Count of
ID, filtered by Organization + Stage), a pie chart (risks by domain), a stacked
column chart (Cat-wise risks by domain), a gold banner title, and the
narrative. Chart source data lives on a hidden ChartData sheet (hidden CELLS
don't plot; a hidden SHEET's ranges do), and series are built manually via
NewSeries — the most reliable COM path. Styling mimics the hand-built MAZ
reference (gold theme, domain colours, % labels on the pie).

Cannot be exercised headlessly; validate on a Windows+Excel machine.
"""
import csv
import os
import subprocess
import sys
import traceback

from zone_reports.stats import CAT_ORDER

DATA_SHEET = "Data"
CHART_SHEET = "ChartData"
TABLE_NAME = "RiskData"

# --- Excel enum constants (COM) ------------------------------------------- #
xlDatabase = 1
xlSrcRange = 1
xlRowField, xlColumnField, xlPageField = 1, 2, 3
xlCount = -4112
xlPie = 5
xlColumnStacked = 52
xlOpenXMLWorkbook = 51
xlHAlignCenter = -4108
xlLegendPositionBottom = -4107
xlSheetHidden = 0


# --- theme colours (Excel Interior.Color wants R + G*256 + B*65536) ------- #
def _rgb(r, g, b):
    return r + g * 256 + b * 65536


GOLD_BANNER = _rgb(255, 192, 0)     # strong amber for the title banner
GOLD_LIGHT = _rgb(255, 217, 102)    # lighter gold for pivot header
GOLD_PALE = _rgb(255, 242, 204)     # pale gold fill for narrative

# Domain colours matched to the reference charts.
DOMAIN_COLORS = {
    "Business Continuity Management": _rgb(31, 86, 103),    # dark teal
    "Information Security":           _rgb(237, 125, 49),   # orange
    "Privacy":                        _rgb(46, 125, 50),    # dark green
    "Security":                       _rgb(68, 184, 213),   # light blue
}

CHART_W, CHART_H = 360, 230         # points
CHART_COL = 8                       # charts anchored at column H
BANNER_ROW = 13                     # fallback banner row (zero-row zones)


# --------------------------------------------------------------------------- #
def _register_pywin32_dlls():
    """pywin32 ships pythoncomXX.dll / pywintypesXX.dll in a `pywin32_system32`
    folder that isn't on the DLL search path after a plain `pip install`
    (esp. Python 3.12+). Put it on the path so win32com can import."""
    import glob
    import site

    dirs = []
    try:
        dirs += list(site.getsitepackages())
    except Exception:
        pass
    try:
        dirs.append(site.getusersitepackages())
    except Exception:
        pass
    dirs.append(os.path.join(os.path.dirname(sys.executable), "Lib", "site-packages"))

    for sp in dirs:
        d = os.path.join(sp, "pywin32_system32")
        if os.path.isdir(d):
            try:
                os.add_dll_directory(d)
            except Exception:
                pass
            os.environ["PATH"] = d + os.pathsep + os.environ.get("PATH", "")
            # some setups also keep the DLLs under win32/ and win32/lib
            for extra in glob.glob(os.path.join(sp, "win32")):
                try:
                    os.add_dll_directory(extra)
                except Exception:
                    pass


def _run_pywin32_postinstall():
    """Official pywin32 post-install: copies the DLLs into place. No-op if the
    script is missing; ignores failure (may need admin)."""
    script = os.path.join(os.path.dirname(sys.executable), "Scripts",
                          "pywin32_postinstall.py")
    if os.path.isfile(script):
        try:
            subprocess.check_call([sys.executable, script, "-install"])
        except Exception as e:
            print(f"  pywin32 post-install step skipped: {e}")


def _ensure_pywin32():
    try:
        import win32com.client  # noqa: F401
        return
    except ImportError:
        pass

    # Install the package only if it's genuinely absent (import may have failed
    # purely because the DLLs weren't on the path).
    try:
        import win32com  # noqa: F401
    except ImportError:
        print("pywin32 not found -> installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pywin32"])

    _register_pywin32_dlls()
    try:
        import win32com.client  # noqa: F401
        return
    except ImportError:
        pass

    # Last resort: run the official post-install, then register + import again.
    _run_pywin32_postinstall()
    _register_pywin32_dlls()
    import win32com.client  # noqa: F401  (raise if still broken)


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
    """Gold pivot style, bold gold header rows, reference caption, autofit."""
    try:
        pt.TableStyle2 = "PivotStyleMedium3"
    except Exception:
        pass
    try:
        # the reference workbook captions the row header "Risk Names"
        pt.PivotFields("Category").Caption = "Risk Names"
    except Exception:
        pass
    try:
        rng = pt.TableRange1
        rng.Font.Size = 10
        hdr = ws.Range(ws.Cells(4, 1), ws.Cells(5, rng.Columns.Count))
        hdr.Interior.Color = GOLD_LIGHT
        hdr.Font.Bold = True
        pt.TableRange2.EntireColumn.AutoFit()
    except Exception:
        pass


# --------------------------------------------------------------------------- #
def _write_banner_title(ws, title, row):
    """Big gold banner merged across the report width at the given row."""
    rng = ws.Range(ws.Cells(row, 1), ws.Cells(row, 14))
    rng.Merge()
    rng.Value = title
    rng.Interior.Color = GOLD_BANNER
    rng.Font.Bold = True
    rng.Font.Size = 20
    rng.HorizontalAlignment = xlHAlignCenter
    ws.Rows(row).RowHeight = 42


def _write_narrative(ws, lines, start):
    """Narrative text down column A on a pale-gold background; headline lines
    (containing '%' or the Key Insights header) bold. Starts at row `start`."""
    for i, line in enumerate(lines):
        cell = ws.Cells(start + i, 1)
        cell.Value = line
        cell.Font.Size = 11
        if line and not line.startswith("    ") and ("%" in line or line.startswith("Key Insights")):
            cell.Font.Bold = True
    block = ws.Range(ws.Cells(start, 1), ws.Cells(start + max(len(lines), 1), 6))
    block.Interior.Color = GOLD_PALE


# --------------------------------------------------------------------------- #
def _write_chart_block(cws, r0, stats):
    """Write one zone's chart source data on the ChartData sheet starting at
    row r0. Returns (pie_x, pie_v, cats_rng, dom_rows) as COM Ranges, where
    dom_rows is [(domain, values_range), ...] in by_domain order."""
    doms = stats.by_domain
    cats = [c for c, _ in stats.grand_by_cat]

    for i, (dom, n) in enumerate(doms, start=1):
        cws.Cells(r0 + i, 1).Value = dom
        cws.Cells(r0 + i, 2).Value = n
    pie_x = cws.Range(cws.Cells(r0 + 1, 1), cws.Cells(r0 + len(doms), 1))
    pie_v = cws.Range(cws.Cells(r0 + 1, 2), cws.Cells(r0 + len(doms), 2))

    c0 = 4
    for j, cat in enumerate(cats):
        cws.Cells(r0, c0 + 1 + j).Value = cat
    cats_rng = cws.Range(cws.Cells(r0, c0 + 1), cws.Cells(r0, c0 + len(cats)))

    dom_rows = []
    for i, (dom, _n) in enumerate(doms, start=1):
        cws.Cells(r0 + i, c0).Value = dom
        row = stats.crosstab.get(dom, {})
        for j, cat in enumerate(cats):
            cws.Cells(r0 + i, c0 + 1 + j).Value = row.get(cat, 0)
        vals = cws.Range(cws.Cells(r0 + i, c0 + 1), cws.Cells(r0 + i, c0 + len(cats)))
        dom_rows.append((dom, vals))
    return pie_x, pie_v, cats_rng, dom_rows


def _new_chart(ws, ctype, left, top, w, h):
    """Create an embedded chart; AddChart2 first, ChartObjects fallback."""
    try:
        return ws.Shapes.AddChart2(-1, ctype, left, top, w, h).Chart
    except Exception:
        ch = ws.ChartObjects().Add(left, top, w, h).Chart
        ch.ChartType = ctype
        return ch


def _add_charts(ws, anchor_row, pie_x, pie_v, cats_rng, dom_rows):
    """Pie (domain share) + stacked column (Cat-wise by domain), anchored at
    column H below the banner, series built manually."""
    left = ws.Cells(anchor_row, CHART_COL).Left
    top = ws.Cells(anchor_row, CHART_COL).Top

    pie = _new_chart(ws, xlPie, left, top, CHART_W, CHART_H)
    s = pie.SeriesCollection().NewSeries()
    s.XValues = pie_x
    s.Values = pie_v
    pie.HasTitle = True
    try:
        pie.ChartTitle.Text = "RISKS IDENTIFIED"
    except Exception:
        pass
    try:
        for i, (dom, _r) in enumerate(dom_rows, start=1):
            col = DOMAIN_COLORS.get(dom)
            if col is not None:
                s.Points(i).Format.Fill.ForeColor.RGB = col
    except Exception:
        pass
    try:
        s.ApplyDataLabels()
        dls = s.DataLabels()
        dls.ShowPercentage = True
        dls.ShowValue = False
        dls.Font.Bold = True
    except Exception:
        pass
    try:
        pie.HasLegend = True
        pie.Legend.Position = xlLegendPositionBottom
    except Exception:
        pass

    bar = _new_chart(ws, xlColumnStacked, left, top + CHART_H + 15, CHART_W, CHART_H)
    for dom, vals in dom_rows:
        sr = bar.SeriesCollection().NewSeries()
        sr.Name = dom
        sr.Values = vals
        sr.XValues = cats_rng
        col = DOMAIN_COLORS.get(dom)
        if col is not None:
            try:
                sr.Format.Fill.ForeColor.RGB = col
            except Exception:
                pass
        try:
            sr.ApplyDataLabels()
        except Exception:
            pass
    bar.HasTitle = True
    try:
        bar.ChartTitle.Text = "Category wise Risks"
    except Exception:
        pass
    try:
        bar.HasLegend = True
        bar.Legend.Position = xlLegendPositionBottom
    except Exception:
        pass


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

        cws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        cws.Name = CHART_SHEET
        chart_row = 1

        for rep in reports:
            ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
            ws.Name = rep["sheet"]
            pt_name = "pt_" + rep["sheet"].replace("-", "_")
            total = rep["stats"].total

            banner_row = BANNER_ROW
            if total > 0:
                try:
                    pt = _build_pivot(cache, ws, rep["org"], pt_name)
                    _style_pivot(pt, ws)
                    tr = pt.TableRange1
                    banner_row = tr.Row + tr.Rows.Count + 2   # clear of the pivot
                except Exception as e:
                    print(f"  WARN: pivot build failed on {rep['sheet']}: {e}")

            try:
                _write_banner_title(ws, rep["title"], banner_row)
                _write_narrative(ws, rep["narrative"], banner_row + 1)
            except Exception as e:
                print(f"  WARN: title/narrative failed on {rep['sheet']}: {e}")

            if total > 0:
                try:
                    pie_x, pie_v, cats_rng, dom_rows = _write_chart_block(
                        cws, chart_row, rep["stats"])
                    _add_charts(ws, banner_row + 1, pie_x, pie_v, cats_rng, dom_rows)
                    chart_row += len(rep["stats"].by_domain) + 3
                except Exception:
                    print(f"  WARN: charts failed on {rep['sheet']}:")
                    traceback.print_exc()

            try:
                ws.Activate()
                excel.ActiveWindow.DisplayGridlines = False
            except Exception:
                pass
            print(f"  built {rep['sheet']}: total={total}")

        # tab order: zone sheets first, then Data; ChartData hidden at the end
        wb.Worksheets(DATA_SHEET).Move(After=wb.Worksheets(wb.Worksheets.Count))
        try:
            cws.Move(After=wb.Worksheets(wb.Worksheets.Count))
        except Exception:
            pass
        try:
            cws.Visible = xlSheetHidden
        except Exception:
            pass
        try:
            wb.Worksheets(reports[0]["sheet"]).Activate()
        except Exception:
            pass

        wb.SaveAs(os.path.abspath(out_path), FileFormat=xlOpenXMLWorkbook)
        print(f"Saved: {out_path}")
    finally:
        if excel is not None:
            excel.Quit()
