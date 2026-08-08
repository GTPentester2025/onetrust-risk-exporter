"""Build an editable PowerPoint deck (native charts/tables) from a dashboard payload.

Dark+gold theme: two slides per zone (Overview + Treatment & Aging), 19 slides total.
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

# Muted palette matching UI
PALETTE = [
    "E8C810", "C9A227", "5FA8A0", "7D93A8",
    "A08BA8", "8A8578", "D6B56A", "4E8AC9",
    "B5766B", "6BAF8D", "9C8FBF", "77746C",
]
BLANK_COLOR = "3E3B35"

# Treated donut: Treated = green, Open = muted gold
TREATED_COLOR = RGBColor(0x34, 0xD3, 0x99)
OPEN_COLOR    = RGBColor(0xC9, 0xA2, 0x27)

# Aging bucket bar colors (green → amber → clay → rust)
BUCKET_COLORS = ["6BAF8D", "C9A227", "C98B4C", "BF6E5E"]

# Aging-by-domain bar ramp (younger = amber, older = rust)
AGING_RAMP = ["C9A227", "C98B4C", "BF6E5E"]


def _series_color(idx, cat=None):
    if cat == "(blank)":
        return RGBColor.from_string(BLANK_COLOR)
    return RGBColor.from_string(PALETTE[idx % len(PALETTE)])


def _aging_ramp_color(idx, total):
    """Map index (0=youngest) to a clay ramp color; clamp to AGING_RAMP."""
    if total <= 1:
        return RGBColor.from_string(AGING_RAMP[-1])
    # fraction 0..1, but age_by_domain is desc (index 0 = oldest)
    frac = idx / max(total - 1, 1)
    # oldest (idx=0) → last ramp color; youngest (idx=total-1) → first ramp color
    ramp_idx = round(frac * (len(AGING_RAMP) - 1))
    return RGBColor.from_string(AGING_RAMP[ramp_idx])


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


def _style_chart_base(chart, legend=True):
    """Apply dark-theme chrome to any chart."""
    from pptx.enum.chart import XL_LEGEND_POSITION
    from pptx.util import Pt
    chart.has_legend = legend
    if legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
    chart.font.size = Pt(9)
    chart.font.color.rgb = THEME["cream"]
    try:
        chart.plot_area.format.fill.background()
    except Exception:
        pass
    try:
        chart.chart_area.format.fill.background()
    except Exception:
        pass


def _add_header(slide, title_text, sub_text, footer_text):
    """Add a two-line gold header band + footer to a slide."""
    from pptx.util import Inches, Pt

    hdr = slide.shapes.add_textbox(Inches(0.4), Inches(0.08), Inches(12.5), Inches(0.9))
    hdr.text_frame.word_wrap = False
    tf = hdr.text_frame
    tf.text = title_text
    _set_font(tf.paragraphs[0].runs[0], 22, THEME["gold"], bold=True)

    p_sub = tf.add_paragraph()
    p_sub.text = sub_text
    _set_font(p_sub, 10, THEME["muted"])

    ft = slide.shapes.add_textbox(
        Inches(0.4), Inches(7.22), Inches(10.0), Inches(0.28),
    )
    ft.text_frame.text = footer_text
    _set_font(ft.text_frame.paragraphs[0], 8, THEME["muted"])


def _add_kpi_strip(slide, z, doms_raw):
    """Add the KPI strip row below the header."""
    from pptx.util import Inches

    kpi = slide.shapes.add_textbox(Inches(0.4), Inches(1.05), Inches(12.5), Inches(0.5))
    kpi.text_frame.word_wrap = False

    total = z.get("total", 0)
    top2 = ", ".join(
        f"{d} ({z['domain_pct'].get(d, 0)}%)" for d, _ in doms_raw[:2]
    ) or "—"
    topcat = (
        max(z["by_cat"].items(), key=lambda kv: kv[1])[0]
        if z["by_cat"] else "—"
    )
    treated = z.get("treated", 0)
    treated_pct = z.get("treated_pct", 0)
    kpi_text = (
        f"Total {total}   |   Top domains: {top2}"
        f"   |   Top category: {_label(topcat)}"
        f"  ·  Treated {treated} ({treated_pct}%)"
    )
    if z.get("untreated"):
        kpi_text += f"  ·  Avg untreated age {z.get('avg_age', 0)}d"

    kpi.text_frame.text = kpi_text
    _set_font(kpi.text_frame.paragraphs[0], 10, THEME["cream"])


# ---------------------------------------------------------------------------
# Chart builders
# ---------------------------------------------------------------------------

def _add_domain_pie(slide, doms, x, y, w, h):
    """Native PIE chart with per-point palette colours and % labels."""
    from pptx.util import Inches, Pt
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE

    cd = CategoryChartData()
    cd.categories = [d for d, _ in doms]
    cd.add_series("Risks", [n for _, n in doms])
    gr = slide.shapes.add_chart(
        XL_CHART_TYPE.PIE,
        Inches(x), Inches(y), Inches(w), Inches(h),
        cd,
    )
    chart = gr.chart
    chart.has_title = True
    chart.chart_title.text_frame.text = "Risks by Domain"
    _set_font(chart.chart_title.text_frame.paragraphs[0], 11, THEME["cream"])
    _style_chart_base(chart)

    series = chart.series[0]
    for i in range(len(doms)):
        pt = series.points[i]
        pt.format.fill.solid()
        col = BLANK_COLOR if doms[i][0].startswith("Other (") else None
        pt.format.fill.fore_color.rgb = (
            RGBColor.from_string(col) if col else _series_color(i)
        )

    plot = chart.plots[0]
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.show_percentage = True
    dl.show_value = False
    dl.number_format = "0%"
    dl.font.size = Pt(9)
    dl.font.color.rgb = THEME["cream"]
    return gr


def _add_stacked_bar(slide, doms, cats, z, x, y, w, h):
    """Stacked column chart (categories = domains, series = risk categories)."""
    from pptx.util import Inches, Pt
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE

    cd = CategoryChartData()
    cd.categories = [d for d, _ in doms]
    for j, c in enumerate(cats):
        cd.add_series(
            _label(c),
            [z["crosstab"].get(d, {}).get(c, 0) for d, _ in doms],
        )
    gr = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_STACKED,
        Inches(x), Inches(y), Inches(w), Inches(h),
        cd,
    )
    chart = gr.chart
    chart.has_title = True
    chart.chart_title.text_frame.text = "Risks by Domain & Category"
    _set_font(chart.chart_title.text_frame.paragraphs[0], 11, THEME["cream"])
    _style_chart_base(chart)

    for j, (c, series) in enumerate(zip(cats, chart.series)):
        series.format.fill.solid()
        series.format.fill.fore_color.rgb = _series_color(j, c)
        series.data_labels.show_value = True
        series.data_labels.show_percentage = False
        series.data_labels.font.size = Pt(8)
        series.data_labels.font.color.rgb = THEME["cream"]
    return gr


def _add_treated_donut(slide, total, treated, x, y, w, h):
    """Treated/Open doughnut chart."""
    from pptx.util import Inches, Pt
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION

    openn = max(0, total - treated)
    cd = CategoryChartData()
    cd.categories = ["Treated", "Open"]
    cd.add_series("Risks", (treated, openn))
    gr = slide.shapes.add_chart(
        XL_CHART_TYPE.DOUGHNUT,
        Inches(x), Inches(y), Inches(w), Inches(h),
        cd,
    )
    ch = gr.chart
    ch.has_title = True
    ch.chart_title.text_frame.text = "Risk Treatment"
    _set_font(ch.chart_title.text_frame.paragraphs[0], 11, THEME["cream"])

    pts = ch.plots[0].series[0].points
    pts[0].format.fill.solid()
    pts[0].format.fill.fore_color.rgb = TREATED_COLOR
    pts[1].format.fill.solid()
    pts[1].format.fill.fore_color.rgb = OPEN_COLOR

    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.BOTTOM
    ch.legend.include_in_layout = False
    ch.font.color.rgb = THEME["cream"]
    ch.font.size = Pt(9)

    try:
        ch.plot_area.format.fill.background()
    except Exception:
        pass
    try:
        ch.chart_area.format.fill.background()
    except Exception:
        pass

    plot = ch.plots[0]
    plot.has_data_labels = True
    plot.data_labels.number_format = '0'
    plot.data_labels.font.size = Pt(9)
    plot.data_labels.font.color.rgb = THEME["cream"]
    return gr


def _add_aging_bar(slide, age_by_domain, x, y, w, h):
    """Horizontal BAR_CLUSTERED chart of avg days by domain (or skip + textbox)."""
    from pptx.util import Inches, Pt
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE

    if not age_by_domain:
        tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tb.text_frame.word_wrap = True
        tb.text_frame.text = "No untreated aging data"
        _set_font(tb.text_frame.paragraphs[0], 10, THEME["muted"])
        return None

    domains = [_label(row[0]) for row in age_by_domain]
    avgs    = [row[1] for row in age_by_domain]
    total_n = len(age_by_domain)

    cd = CategoryChartData()
    cd.categories = domains
    cd.add_series("Avg days", avgs)

    gr = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(x), Inches(y), Inches(w), Inches(h),
        cd,
    )
    chart = gr.chart
    chart.has_title = True
    chart.chart_title.text_frame.text = "Risk Aging by Domain"
    _set_font(chart.chart_title.text_frame.paragraphs[0], 11, THEME["cream"])
    _style_chart_base(chart, legend=False)

    series = chart.series[0]
    for i in range(total_n):
        pt = series.points[i]
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = _aging_ramp_color(i, total_n)

    series.data_labels.show_value = True
    series.data_labels.number_format = '0"d"'
    series.data_labels.font.size = Pt(9)
    series.data_labels.font.color.rgb = THEME["cream"]
    return gr


def _add_bucket_chart(slide, age_buckets, x, y, w, h):
    """COLUMN_CLUSTERED chart of aging bucket counts."""
    from pptx.util import Inches, Pt
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE

    labels = [b[0] for b in age_buckets]
    counts = [b[1] for b in age_buckets]

    cd = CategoryChartData()
    cd.categories = labels
    cd.add_series("Count", counts)

    gr = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(x), Inches(y), Inches(w), Inches(h),
        cd,
    )
    chart = gr.chart
    chart.has_title = True
    chart.chart_title.text_frame.text = "Aging Buckets"
    _set_font(chart.chart_title.text_frame.paragraphs[0], 11, THEME["cream"])
    _style_chart_base(chart, legend=False)

    series = chart.series[0]
    for i in range(len(age_buckets)):
        pt = series.points[i]
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = RGBColor.from_string(
            BUCKET_COLORS[i] if i < len(BUCKET_COLORS) else BUCKET_COLORS[-1]
        )

    series.data_labels.show_value = True
    series.data_labels.font.size = Pt(9)
    series.data_labels.font.color.rgb = THEME["cream"]
    return gr


def _add_crosstab_table(slide, doms, cats, z, x, y, w, h):
    """Domain × category crosstab table."""
    from pptx.util import Inches

    colcats = cats
    ncol = 2 + len(colcats)
    nrow = len(doms) + 1
    if nrow < 2:
        nrow = 2

    tbl_shape = slide.shapes.add_table(
        nrow, ncol,
        Inches(x), Inches(y), Inches(w), Inches(h),
    )
    tbl = tbl_shape.table

    tbl.columns[0].width = Inches(1.8)
    if ncol > 1:
        rest_w = Inches(w) - Inches(1.8)
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

    if not doms and nrow == 2:
        for ci in range(ncol):
            _cell_fill(tbl.cell(1, ci), THEME["tile"])

    return tbl_shape


def _add_key_insights(slide, z, x, y, w, h, max_lines=22):
    """Key Insights text box with full narrative + aging lines."""
    from pptx.util import Inches

    nb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    nb.text_frame.word_wrap = True
    tf = nb.text_frame

    tf.text = "Key Insights"
    _set_font(tf.paragraphs[0].runs[0], 12, THEME["gold"], bold=True)

    narrative = z.get("narrative", [])
    lines = []
    for line in narrative:
        lines.extend(line.split("\n"))

    aging_lines = []
    if z.get("untreated") and z.get("aged_count"):
        aging_lines.append(
            f"Untreated avg {z['avg_age']}d · median {z['median_age']}d"
            f" · oldest {z['oldest_age']}d"
        )
        for d, avg, c in (z.get("age_by_domain") or [])[:3]:
            aging_lines.append(f"{_label(d)}: {avg} d avg ({c})")

    if aging_lines:
        available = max_lines - len(aging_lines)
        if available < 0:
            available = 0
        if len(lines) > available:
            lines = lines[:available]
            lines.append("…")
        lines = lines + aging_lines
    else:
        if len(lines) > max_lines:
            lines = lines[:max_lines]
            lines.append("…")

    for line in lines:
        p_n = tf.add_paragraph()
        p_n.text = line or " "
        _set_font(p_n, 9, THEME["cream"])

    return nb


# ---------------------------------------------------------------------------
# Build deck
# ---------------------------------------------------------------------------

def build_deck(payload):
    ensure_pptx()
    from pptx import Presentation
    from pptx.util import Inches, Pt

    prs = Presentation()
    prs.slide_width  = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # -----------------------------------------------------------------------
    # Title slide
    # -----------------------------------------------------------------------
    s = prs.slides.add_slide(blank)
    _set_bg(s)

    rule = s.shapes.add_shape(
        1,
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
    _set_font(tf.paragraphs[0].runs[0], 40, THEME["gold"], bold=True)

    p2 = tf.add_paragraph()
    p2.text = f'{payload.get("view", "")}  ·  generated {payload.get("generated_at", "")}'
    p2.runs[0].font.size = Pt(16)
    p2.runs[0].font.color.rgb = THEME["muted"]

    # -----------------------------------------------------------------------
    # Zone slides: two per zone (A=Overview, B=Treatment & Aging)
    # -----------------------------------------------------------------------
    generated = payload.get("generated_at", "")
    date_str  = generated[:10] if generated else ""

    for code, z in payload["zones"].items():
        total    = z.get("total", 0)
        doms_raw = z.get("by_domain", [])
        doms     = _fit_domains(doms_raw)
        cats     = _active_cats(payload, z)

        # ===================================================================
        # SLIDE A — Overview
        # ===================================================================
        sA = prs.slides.add_slide(blank)
        _set_bg(sA)

        _add_header(
            sA,
            f"{code} — Risk Insights (1/2)",
            f"{total} risks · till {date_str}",
            f"{code} · overview",
        )
        _add_kpi_strip(sA, z, doms_raw)

        # Domain PIE  left=0.4, top=1.7, w=4.6, h=3.4
        if doms:
            _add_domain_pie(sA, doms, 0.4, 1.7, 4.6, 3.4)

        # Stacked bar  left=5.3, top=1.7, w=7.6, h=3.4
        if doms and cats:
            _add_stacked_bar(sA, doms, cats, z, 5.3, 1.7, 7.6, 3.4)

        # Crosstab table  left=0.4, top=5.3, w=12.5, h=1.9
        _add_crosstab_table(sA, doms, cats, z, 0.4, 5.3, 12.5, 1.9)

        # ===================================================================
        # SLIDE B — Treatment & Aging
        # ===================================================================
        sB = prs.slides.add_slide(blank)
        _set_bg(sB)

        _add_header(
            sB,
            f"{code} — Treatment & Aging (2/2)",
            f"{total} risks · till {date_str}",
            f"{code} · treatment",
        )

        # Treated doughnut  left=0.4, top=1.7, w=3.6, h=3.0
        if total:
            _add_treated_donut(sB, total, z.get("treated", 0), 0.4, 1.7, 3.6, 3.0)

        # Aging bar  left=4.2, top=1.7, w=4.4, h=3.0
        age_by_domain = z.get("age_by_domain") or []
        _add_aging_bar(sB, age_by_domain, 4.2, 1.7, 4.4, 3.0)

        # Bucket column chart  left=8.8, top=1.7, w=4.1, h=3.0
        age_buckets = z.get("age_buckets") or []
        if age_buckets and any(b[1] for b in age_buckets):
            _add_bucket_chart(sB, age_buckets, 8.8, 1.7, 4.1, 3.0)

        # Key Insights  left=0.4, top=5.0, w=12.5, h=2.2 (22-line cap)
        _add_key_insights(sB, z, 0.4, 5.0, 12.5, 2.2, max_lines=22)

    return prs


def deck_to_bytes(payload):
    prs = build_deck(payload)
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()
