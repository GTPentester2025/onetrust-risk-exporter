import io
import pytest

pptx = pytest.importorskip("pptx")  # skip cleanly if python-pptx not installed
from pptx import Presentation
from pptx.util import Emu
import ppt_export

def _payload():
    zones = {}
    for i, code in enumerate(["GHQ", "AFR", "SAZ", "MAZ", "NAZ", "APAC",
                              "EUR", "GRO", "Overall"]):
        zones[code] = {
            "organization": code, "total": 3,
            "treated": 1, "treated_pct": 33,
            "by_stage": [["Monitoring", 1], ["Assessment", 2]],
            "by_cat": {"COMMERCIAL": 2, "NCI": 1},
            "grand_by_cat": [["COMMERCIAL", 2], ["NCI", 1]],
            "by_domain": [["Security", 2], ["Privacy", 1]],
            "crosstab": {"Security": {"COMMERCIAL": 2}, "Privacy": {"NCI": 1}},
            "domain_pct": {"Security": 67, "Privacy": 33},
            "top_cats": {"Security": [["COMMERCIAL", 2]]},
            "narrative": ["A total of 3 risks identified across business units"],
            "untreated": 2, "aged_count": 2, "avg_age": 90, "median_age": 85,
            "oldest_age": 150,
            "age_by_domain": [["Security", 120, 1], ["Privacy", 60, 1]],
            "age_buckets": [["0-30", 0], ["31-90", 1], ["91-180", 1], ["180+", 0]],
        }
    return {"view": "TPRM Global View", "generated_at": "2026-08-05T06:00:00",
            "cat_order": ["COMMERCIAL", "LOGISTICS", "NCI", "PACKAGING",
                          "TECHNOLOGY", "RAU", "Fees"], "zones": zones}

def test_deck_has_title_plus_one_slide_per_zone():
    b = ppt_export.deck_to_bytes(_payload())
    prs = Presentation(io.BytesIO(b))
    assert len(prs.slides) == 10  # title + 9 zones

def test_zone_slide_has_chart_and_table():
    prs = Presentation(io.BytesIO(ppt_export.deck_to_bytes(_payload())))
    zone_slide = prs.slides[1]
    has_chart = any(sh.has_chart for sh in zone_slide.shapes)
    has_table = any(sh.has_table for sh in zone_slide.shapes)
    assert has_chart and has_table

def test_zone_slide_has_three_charts():
    prs = Presentation(io.BytesIO(ppt_export.deck_to_bytes(_payload())))
    slide = prs.slides[1]
    charts = [sh for sh in slide.shapes if sh.has_chart]
    assert len(charts) >= 3      # pie, stacked bar, treated donut
    assert any(sh.has_table for sh in slide.shapes)

def test_zero_total_zone_no_crash():
    """Regression: zero-total zone should not crash and should have no chart."""
    payload = _payload()
    # Override GHQ zone to have zero total
    payload["zones"]["GHQ"] = {
        "organization": "GHQ",
        "total": 0,
        "treated": 0, "treated_pct": 0,
        "by_stage": [],
        "by_cat": {},
        "grand_by_cat": [],
        "by_domain": [],
        "crosstab": {},
        "domain_pct": {},
        "top_cats": {},
        "narrative": ["A total of 0 risks identified across business units"],
    }

    # Should not crash on deck_to_bytes
    b = ppt_export.deck_to_bytes(payload)
    prs = Presentation(io.BytesIO(b))

    # Assert slide count is 10 (title + 9 zones)
    assert len(prs.slides) == 10

    # GHQ is the first zone slide (slides[1] since slides[0] is title)
    ghq_slide = prs.slides[1]

    # Zero-total zone should have no chart
    has_chart = any(sh.has_chart for sh in ghq_slide.shapes)
    assert not has_chart, "Zero-total zone should not have a chart"

def test_table_dimensions():
    """Regression: table should have correct rows and columns based on by_domain and active_cats."""
    payload = _payload()
    b = ppt_export.deck_to_bytes(payload)
    prs = Presentation(io.BytesIO(b))

    # Get first zone slide (slides[1])
    first_zone_slide = prs.slides[1]

    # Find the table on the slide
    tbl = next(sh.table for sh in first_zone_slide.shapes if sh.has_table)

    # For the default payload:
    # - by_domain has ["Security", "Privacy"] (2 domains)
    # - by_cat has {"COMMERCIAL": 2, "NCI": 1} (2 active cats)
    # - cat_order includes both COMMERCIAL and NCI
    # Rows: len(by_domain) + 1 header = 2 + 1 = 3
    # Cols: Domain (1) + active_cats (2) + Total (1) = 4

    expected_rows = 3  # header + 2 domains
    expected_cols = 4  # Domain + COMMERCIAL + NCI + Total

    assert len(tbl.rows) == expected_rows, f"Expected {expected_rows} rows, got {len(tbl.rows)}"
    assert len(tbl.columns) == expected_cols, f"Expected {expected_cols} cols, got {len(tbl.columns)}"


def test_many_domains_table_capped():
    p = _payload()
    z = p["zones"]["GHQ"]
    doms = [[f"Domain{i}", 10 - i] for i in range(10)]
    z["by_domain"] = doms
    z["domain_pct"] = {d: 10 for d, _ in doms}
    z["crosstab"] = {d: {"COMMERCIAL": n} for d, n in doms}
    z["total"] = sum(n for _, n in doms)
    prs = Presentation(io.BytesIO(ppt_export.deck_to_bytes(p)))
    slide = prs.slides[1]      # GHQ
    tbl = next(sh.table for sh in slide.shapes if sh.has_table)
    assert len(tbl.rows) <= 9      # header + 7 + Other
    texts = [tbl.cell(r, 0).text for r in range(len(tbl.rows))]
    assert any(t.startswith("Other") for t in texts)


def test_zone_slide_labeled_and_fits():
    prs = Presentation(io.BytesIO(ppt_export.deck_to_bytes(_payload())))
    slide = prs.slides[1]
    texts = " ".join(sh.text_frame.text for sh in slide.shapes if sh.has_text_frame)
    assert "GHQ" in texts and "Risk Insights" in texts
    W, H = Emu(int(12192000)), Emu(int(6858000))   # 13.333x7.5in
    for sh in slide.shapes:
        assert sh.left >= 0 and sh.top >= 0
        assert sh.left + sh.width <= W + Emu(1)
        assert sh.top + sh.height <= H + Emu(1)


def test_dark_theme_background():
    from pptx.enum.dml import MSO_FILL
    prs = Presentation(io.BytesIO(ppt_export.deck_to_bytes(_payload())))
    for slide in prs.slides:
        fill = slide.background.fill
        assert fill.type == MSO_FILL.SOLID
        assert str(fill.fore_color.rgb) == "1B1A19"


def test_treated_donut_colors_and_values():
    """Regression: donut chart should show correct colors and values for Treated/Open."""
    from pptx.enum.chart import XL_CHART_TYPE

    payload = _payload()
    prs = Presentation(io.BytesIO(ppt_export.deck_to_bytes(payload)))

    # First zone slide (index 1, since index 0 is title)
    zone_slide = prs.slides[1]
    zone_data = payload["zones"]["GHQ"]

    # Find the doughnut chart among the slide's charts
    donut_chart = None
    for shape in zone_slide.shapes:
        if shape.has_chart:
            if shape.chart.chart_type == XL_CHART_TYPE.DOUGHNUT:
                donut_chart = shape.chart
                break

    assert donut_chart is not None, "Doughnut chart not found on slide"

    # Check categories
    categories = list(donut_chart.plots[0].categories)
    assert categories == ["Treated", "Open"], f"Expected ['Treated', 'Open'], got {categories}"

    # Check values: (treated, total-treated)
    values = donut_chart.plots[0].series[0].values
    treated = zone_data["treated"]
    total = zone_data["total"]
    expected_values = (treated, total - treated)
    assert values == expected_values, f"Expected {expected_values}, got {values}"

    # Check point colors
    points = donut_chart.plots[0].series[0].points

    # Point 0: Treated (green #34D399)
    color_0 = str(points[0].format.fill.fore_color.rgb)
    assert color_0 == "34D399", f"Point 0 color: expected '34D399', got '{color_0}'"

    # Point 1: Open (gold #E8C810)
    color_1 = str(points[1].format.fill.fore_color.rgb)
    assert color_1 == "E8C810", f"Point 1 color: expected 'E8C810', got '{color_1}'"


def test_slide_has_aging_text():
    prs = Presentation(io.BytesIO(ppt_export.deck_to_bytes(_payload())))
    slide = prs.slides[1]
    texts = " ".join(sh.text_frame.text for sh in slide.shapes if sh.has_text_frame)
    assert "Avg untreated age 90d" in texts
    assert "days old" in texts or "d avg" in texts   # aging appears in insights
