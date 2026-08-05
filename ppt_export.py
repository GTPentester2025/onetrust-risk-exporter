"""Build an editable PowerPoint deck (native charts/tables) from a dashboard payload.

Dark+gold theme (Task 2): fixed-fit layout, labeled zone slides.
"""
import io
import subprocess
import sys


def ensure_pptx():
    try:
        import pptx  # noqa: F401
    except ImportError:
        subprocess.run([sys.executable, "-m", "pip", "install", "python-pptx"], check=True)


def _label(cat):
    return "Uncategorized" if cat == "(blank)" else cat


def _active_cats(payload, z):
    known = [c for c in payload["cat_order"] if c in z["by_cat"]]
    extras = [c for c in z["by_cat"] if c not in payload["cat_order"]]
    return known + extras


# ---------------------------------------------------------------------------
# Theme constants
# ---------------------------------------------------------------------------
ensure_pptx()
from pptx.dml.color import RGBColor  # noqa: E402

THEME = {
    "bg":    RGBColor(0x1B, 0x1A, 0x19),
    "gold":  RGBColor(0xE8, 0xC8, 0x10),
    "cream": RGBColor(0xF3, 0xF2, 0xF1),
    "muted": RGBColor(0xC8, 0xC6, 0xC4),
    "tile":  RGBColor(0x25, 0x24, 0x23),
    "tile2": RGBColor(0x2F, 0x2E, 0x2D),
}

PALETTE = [
    "E8C810", "4FC3F7", "F0736A", "34D399",
    "CE93D8", "FFB74D", "90CAF9", "F48FB1",
    "AED581", "FFD54F", "4DD0E1", "BCAAA4",
]
BLANK_COLOR = "797775"


def _series_color(idx, cat=None):
    if cat == "(blank)":
        return RGBColor.from_string(BLANK_COLOR)
    return RGBColor.from_string(PALETTE[idx % len(PALETTE)])


def _fit_domains(doms, limit=8):
    """Return at most *limit* domain rows; excess rows collapsed into 'Other (K)'."""
    if len(doms) <= limit:
        return list(doms)
    head = list(doms[:limit - 1])
    other_count = len(doms) - (limit - 1)
    other_total = sum(n for _, n in doms[limit - 1:])
    head.append((f"Other ({other_count})", other_total))
    return head


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _set_bg(slide):
    """Paint slide background with the dark theme colour."""
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = THEME["bg"]


def _set_font(run_or_para, size_pt, color, bold=False):
    from pptx.util import Pt
    f = run_or_para.font
    f.size = Pt(size_pt)
    f.color.rgb = color
    f.bold = bold


def _cell_fill(cell, color):
    cell.fill.solid()
    cell.fill.fore_color.rgb = color


def _style_chart_base(chart):
    """Apply dark-theme chrome to any chart."""
    from pptx.enum.chart import XL_LEGEND_POSITION
    from pptx.util import Pt
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    chart.font.size = Pt(9)
    chart.font.color.rgb = THEME["cream"]
    # Transparent plot area / chart area backgrounds
    try:
        chart.plot_area.format.fill.background()
    except Exception:
        pass
    try:
        chart.chart_area.format.fill.background()
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Build deck
# ---------------------------------------------------------------------------

def build_deck(payload):
    ensure_pptx()
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.enum.text import PP_ALIGN
    from pptx.oxml.ns import qn
    from lxml import etree
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # -----------------------------------------------------------------------
    # Title slide
    # -----------------------------------------------------------------------
    s = prs.slides.add_slide(blank)
    _set_bg(s)

    # Gold rule bar
    rule = s.shapes.add_shape(
        1,  # MSO_SHAPE.RECTANGLE = 1
        Inches(0.7), Inches(2.4),
        Inches(11.9), Inches(0.03),
    )
    rule.line.fill.background()
    rule.fill.solid()
    rule.fill.fore_color.rgb = THEME["gold"]

    tb = s.shapes.add_textbox(Inches(0.7), Inches(2.5), Inches(12.0), Inches(1.5))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.text = "OneTrust Risk Insights"
    run = tf.paragraphs[0].runs[0]
    _set_font(run, 40, THEME["gold"], bold=True)

    p2 = tf.add_paragraph()
    p2.text = f'{payload.get("view", "")}  ·  generated {payload.get("generated_at", "")}'
    p2.runs[0].font.size = Pt(16)
    p2.runs[0].font.color.rgb = THEME["muted"]

    # -----------------------------------------------------------------------
    # Zone slides
    # -----------------------------------------------------------------------
    slide_idx = 1  # 0 = title
    for code, z in payload["zones"].items():
        s = prs.slides.add_slide(blank)
        _set_bg(s)
        slide_idx += 1

        total = z.get("total", 0)
        doms_raw = z.get("by_domain", [])
        doms = _fit_domains(doms_raw)
        cats = _active_cats(payload, z)

        # -------------------------------------------------------------------
        # Header band  (top=0.0 → height=0.9)
        # -------------------------------------------------------------------
        hdr = s.shapes.add_textbox(Inches(0.5), Inches(0.08), Inches(12.3), Inches(0.9))
        hdr.text_frame.word_wrap = False
        tf_h = hdr.text_frame

        # Line 1: zone title in gold
        tf_h.text = f"{code} — Risk Insights"
        run_h1 = tf_h.paragraphs[0].runs[0]
        _set_font(run_h1, 22, THEME["gold"], bold=True)

        # Line 2: count + date in muted
        generated = payload.get("generated_at", "")
        date_str = generated[:10] if generated else ""
        p_sub = tf_h.add_paragraph()
        p_sub.text = f"{total} risks · till {date_str}"
        _set_font(p_sub, 10, THEME["muted"])

        # -------------------------------------------------------------------
        # KPI strip  (top=1.0)
        # -------------------------------------------------------------------
        kpi = s.shapes.add_textbox(Inches(0.5), Inches(1.0), Inches(12.3), Inches(0.5))
        kpi.text_frame.word_wrap = False
        top2 = ", ".join(
            f"{d} ({z['domain_pct'].get(d, 0)}%)" for d, _ in doms_raw[:2]
        ) or "—"
        topcat = (
            max(z["by_cat"].items(), key=lambda kv: kv[1])[0]
            if z["by_cat"] else "—"
        )
        kpi.text_frame.text = (
            f"Total {total}   |   Top domains: {top2}"
            f"   |   Top category: {_label(topcat)}"
        )
        _set_font(kpi.text_frame.paragraphs[0], 10, THEME["cream"])

        # -------------------------------------------------------------------
        # Pie chart  left=0.5, top=1.6, w=5.9, h=3.1
        # -------------------------------------------------------------------
        if doms:
            from pptx.util import Pt as Pt_
            cd = CategoryChartData()
            cd.categories = [d for d, _ in doms]
            cd.add_series("Risks", [n for _, n in doms])
            gr = s.shapes.add_chart(
                XL_CHART_TYPE.PIE,
                Inches(0.5), Inches(1.6),
                Inches(5.9), Inches(3.1),
                cd,
            )
            chart = gr.chart
            _style_chart_base(chart)

            # Per-point colours
            series = chart.series[0]
            for i in range(len(doms)):
                pt = series.points[i]
                pt.format.fill.solid()
                pt.format.fill.fore_color.rgb = _series_color(i)

            # Percentage data labels
            plot = chart.plots[0]
            plot.has_data_labels = True
            dl = plot.data_labels
            dl.show_percentage = True
            dl.show_value = False
            dl.number_format = "0%"
            dl.font.size = Pt(9)
            dl.font.color.rgb = THEME["cream"]

        # -------------------------------------------------------------------
        # Stacked bar  left=6.6, top=1.6, w=6.2, h=3.1
        # -------------------------------------------------------------------
        if doms and cats:
            from pptx.enum.chart import XL_LABEL_POSITION
            cd = CategoryChartData()
            cd.categories = [d for d, _ in doms]
            for j, c in enumerate(cats):
                cd.add_series(
                    _label(c),
                    [z["crosstab"].get(d, {}).get(c, 0) for d, _ in doms],
                )
            gr2 = s.shapes.add_chart(
                XL_CHART_TYPE.COLUMN_STACKED,
                Inches(6.6), Inches(1.6),
                Inches(6.2), Inches(3.1),
                cd,
            )
            chart2 = gr2.chart
            _style_chart_base(chart2)

            # Per-series colours + value labels
            for j, (c, series) in enumerate(zip(cats, chart2.series)):
                series.format.fill.solid()
                series.format.fill.fore_color.rgb = _series_color(j, c)
                series.data_labels.show_value = True
                series.data_labels.show_percentage = False
                series.data_labels.font.size = Pt(8)
                series.data_labels.font.color.rgb = THEME["cream"]

        # -------------------------------------------------------------------
        # Table  left=0.5, top=4.9, w=7.6, h=2.3
        # -------------------------------------------------------------------
        colcats = cats
        ncol = 2 + len(colcats)
        nrow = len(doms) + 1  # header + domain rows
        if nrow < 2:
            nrow = 2  # pptx minimum

        tbl_shape = s.shapes.add_table(
            nrow, ncol,
            Inches(0.5), Inches(4.9),
            Inches(7.6), Inches(2.3),
        )
        tbl = tbl_shape.table

        # Column widths: first col 1.8in, rest even across remaining 5.8in
        tbl.columns[0].width = Inches(1.8)
        if ncol > 1:
            rest_w = Inches(7.6) - Inches(1.8)
            each_w = int(rest_w / (ncol - 1))
            for ci in range(1, ncol):
                tbl.columns[ci].width = each_w

        font_sz = 10 if ncol <= 5 else 8

        # Header row
        _cell_fill(tbl.cell(0, 0), THEME["tile2"])
        tbl.cell(0, 0).text = "Domain"
        _set_font(tbl.cell(0, 0).text_frame.paragraphs[0], 9, THEME["gold"], bold=True)

        for j, c in enumerate(colcats):
            _cell_fill(tbl.cell(0, j + 1), THEME["tile2"])
            tbl.cell(0, j + 1).text = _label(c)
            _set_font(tbl.cell(0, j + 1).text_frame.paragraphs[0], 9, THEME["gold"], bold=True)

        _cell_fill(tbl.cell(0, ncol - 1), THEME["tile2"])
        tbl.cell(0, ncol - 1).text = "Total"
        _set_font(tbl.cell(0, ncol - 1).text_frame.paragraphs[0], 9, THEME["gold"], bold=True)

        # Body rows
        for i, (d, dn) in enumerate(doms, start=1):
            row_color = THEME["tile"] if i % 2 == 1 else THEME["tile2"]
            _cell_fill(tbl.cell(i, 0), row_color)
            tbl.cell(i, 0).text = d
            _set_font(tbl.cell(i, 0).text_frame.paragraphs[0], font_sz, THEME["cream"])

            for j, c in enumerate(colcats):
                _cell_fill(tbl.cell(i, j + 1), row_color)
                tbl.cell(i, j + 1).text = str(z["crosstab"].get(d, {}).get(c, 0))
                _set_font(tbl.cell(i, j + 1).text_frame.paragraphs[0], font_sz, THEME["cream"])

            _cell_fill(tbl.cell(i, ncol - 1), row_color)
            tbl.cell(i, ncol - 1).text = str(dn)
            _set_font(tbl.cell(i, ncol - 1).text_frame.paragraphs[0], font_sz, THEME["cream"])

        # Fill remaining rows if nrow was padded to 2
        if not doms and nrow == 2:
            for ci in range(ncol):
                _cell_fill(tbl.cell(1, ci), THEME["tile"])

        # -------------------------------------------------------------------
        # Key Insights box  left=8.3, top=4.9, w=4.5, h=2.3
        # -------------------------------------------------------------------
        nb = s.shapes.add_textbox(Inches(8.3), Inches(4.9), Inches(4.5), Inches(2.3))
        nb.text_frame.word_wrap = True
        tf_n = nb.text_frame

        tf_n.text = "Key Insights"
        _set_font(tf_n.paragraphs[0].runs[0], 12, THEME["gold"], bold=True)

        narrative = z.get("narrative", [])
        lines = []
        for line in narrative:
            lines.extend(line.split("\n"))

        if len(lines) > 16:
            lines = lines[:16]
            lines.append("…")

        for line in lines:
            p_n = tf_n.add_paragraph()
            p_n.text = line or " "
            _set_font(p_n, 9, THEME["cream"])

        # -------------------------------------------------------------------
        # Footer  bottom-left, muted 8pt
        # -------------------------------------------------------------------
        ft = s.shapes.add_textbox(
            Inches(0.5), Inches(7.2),
            Inches(8.0), Inches(0.28),
        )
        ft.text_frame.text = f"{code} · slide {slide_idx - 1}"
        _set_font(ft.text_frame.paragraphs[0], 8, THEME["muted"])

    return prs


def deck_to_bytes(payload):
    prs = build_deck(payload)
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()
