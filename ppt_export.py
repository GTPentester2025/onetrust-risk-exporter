"""Build an editable PowerPoint deck (native charts/tables) from a dashboard payload."""
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


def build_deck(payload):
    ensure_pptx()
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # Title slide
    s = prs.slides.add_slide(blank)
    tb = s.shapes.add_textbox(Inches(0.7), Inches(2.6), Inches(12), Inches(2))
    tf = tb.text_frame
    tf.text = "OneTrust Risk Insights"
    tf.paragraphs[0].runs[0].font.size = Pt(40)
    p = tf.add_paragraph()
    p.text = f'{payload.get("view", "")}  ·  generated {payload.get("generated_at", "")}'
    p.font.size = Pt(16)

    for code, z in payload["zones"].items():
        s = prs.slides.add_slide(blank)
        head = s.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.3), Inches(0.8))
        head.text_frame.text = f"{code} — {z['total']} risks"
        head.text_frame.paragraphs[0].runs[0].font.size = Pt(26)

        # KPI text
        kpi = s.shapes.add_textbox(Inches(0.5), Inches(1.1), Inches(12.3), Inches(0.6))
        doms = z["by_domain"]
        top2 = ", ".join(f"{d} ({z['domain_pct'].get(d,0)}%)" for d, _ in doms[:2]) or "—"
        topcat = max(z["by_cat"].items(), key=lambda kv: kv[1])[0] if z["by_cat"] else "—"
        kpi.text_frame.text = (f"Total {z['total']}   |   Top domains: {top2}"
                               f"   |   Top category: {_label(topcat)}")
        kpi.text_frame.paragraphs[0].font.size = Pt(12)

        # Pie: risks by domain (native, editable)
        if doms:
            cd = CategoryChartData()
            cd.categories = [d for d, _ in doms]
            cd.add_series("Risks", [n for _, n in doms])
            s.shapes.add_chart(XL_CHART_TYPE.PIE, Inches(0.5), Inches(1.9),
                               Inches(6), Inches(2.6), cd)

        # Stacked bar: domain x Cat (native, editable)
        cats = _active_cats(payload, z)
        if doms and cats:
            cd = CategoryChartData()
            cd.categories = [d for d, _ in doms]
            for c in cats:
                cd.add_series(_label(c), [z["crosstab"].get(d, {}).get(c, 0) for d, _ in doms])
            ch = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, Inches(6.7), Inches(1.9),
                                    Inches(6.1), Inches(2.6), cd)
            ch.chart.has_legend = True

        # Crosstab table (native)
        rows = len(doms) + 1
        colcats = cats
        ncol = 2 + len(colcats)
        tbl = s.shapes.add_table(rows if rows > 1 else 2, ncol,
                                 Inches(0.5), Inches(4.7), Inches(8.3), Inches(2.4)).table
        tbl.cell(0, 0).text = "Domain"
        for j, c in enumerate(colcats):
            tbl.cell(0, 1 + j).text = _label(c)
        tbl.cell(0, ncol - 1).text = "Total"
        for i, (d, dn) in enumerate(doms, start=1):
            tbl.cell(i, 0).text = d
            for j, c in enumerate(colcats):
                tbl.cell(i, 1 + j).text = str(z["crosstab"].get(d, {}).get(c, 0))
            tbl.cell(i, ncol - 1).text = str(dn)

        # Narrative text box
        nb = s.shapes.add_textbox(Inches(9.0), Inches(4.7), Inches(3.8), Inches(2.4))
        nb.text_frame.word_wrap = True
        nb.text_frame.text = "\n".join(z["narrative"]) or "No data"
        for para in nb.text_frame.paragraphs:
            para.font.size = Pt(9)

    return prs


def deck_to_bytes(payload):
    prs = build_deck(payload)
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()
