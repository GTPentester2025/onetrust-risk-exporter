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


# ---------------------------------------------------------------------------
# Slide count: title + 9 zones x 2 = 19
# ---------------------------------------------------------------------------
def test_deck_has_title_plus_two_slides_per_zone():
    b = ppt_export.deck_to_bytes(_payload())
    prs = Presentation(io.BytesIO(b))
    assert len(prs.slides) == 19   # title + 9 zones × 2


# ---------------------------------------------------------------------------
# Slide A (overview) has chart(s) and a table; slide B exists
# ---------------------------------------------------------------------------
def test_zone_slide_has_chart_and_table():
    prs = Presentation(io.BytesIO(ppt_export.deck_to_bytes(_payload())))
    slide_a = prs.slides[1]   # first zone, slide A
    has_chart = any(sh.has_chart for sh in slide_a.shapes)
    has_table = any(sh.has_table for sh in slide_a.shapes)
    assert has_chart and has_table


# ---------------------------------------------------------------------------
# Slide A: >=2 charts + table; Slide B: >=3 charts
# ---------------------------------------------------------------------------
def test_zone_slide_a_has_two_charts_and_table():
    prs = Presentation(io.BytesIO(ppt_export.deck_to_bytes(_payload())))
    slide_a = prs.slides[1]
    charts = [sh for sh in slide_a.shapes if sh.has_chart]
    assert len(charts) >= 2, f"Slide A expected >=2 charts, got {len(charts)}"
    assert any(sh.has_table for sh in slide_a.shapes), "Slide A missing table"


def test_zone_slide_b_has_three_charts():
    prs = Presentation(io.BytesIO(ppt_export.deck_to_bytes(_payload())))
    slide_b = prs.slides[2]   # first zone, slide B
    charts = [sh for sh in slide_b.shapes if sh.has_chart]
    assert len(charts) >= 3, f"Slide B expected >=3 charts, got {len(charts)}"


# ---------------------------------------------------------------------------
# Slide B chart types: BAR_CLUSTERED + COLUMN_CLUSTERED
# ---------------------------------------------------------------------------
def test_slide_b_has_bar_and_column_charts():
    from pptx.enum.chart import XL_CHART_TYPE
    prs = Presentation(io.BytesIO(ppt_export.deck_to_bytes(_payload())))
    slide_b = prs.slides[2]
    types = [sh.chart.chart_type for sh in slide_b.shapes if sh.has_chart]
    assert XL_CHART_TYPE.BAR_CLUSTERED in types, "Slide B missing BAR_CLUSTERED"
    assert XL_CHART_TYPE.COLUMN_CLUSTERED in types, "Slide B missing COLUMN_CLUSTERED"


# ---------------------------------------------------------------------------
# Aging bar categories match age_by_domain domains
# ---------------------------------------------------------------------------
def test_aging_bar_categories_match_age_by_domain():
    from pptx.enum.chart import XL_CHART_TYPE
    payload = _payload()
    prs = Presentation(io.BytesIO(ppt_export.deck_to_bytes(payload)))
    slide_b = prs.slides[2]   # GHQ slide B
    bar = next(
        sh.chart for sh in slide_b.shapes
        if sh.has_chart and sh.chart.chart_type == XL_CHART_TYPE.BAR_CLUSTERED
    )
    cats = list(bar.plots[0].categories)
    expected = [ppt_export._label(row[0]) for row in payload["zones"]["GHQ"]["age_by_domain"]]
    assert cats == expected, f"Aging bar cats: expected {expected}, got {cats}"


# ---------------------------------------------------------------------------
# Donut lives on Slide B (index 2)
# ---------------------------------------------------------------------------
def test_treated_donut_colors_and_values():
    """Regression: donut chart on slide B shows correct colors and values."""
    from pptx.enum.chart import XL_CHART_TYPE

    payload = _payload()
    prs = Presentation(io.BytesIO(ppt_export.deck_to_bytes(payload)))

    # Donut now on slide B (index 2)
    slide_b = prs.slides[2]
    zone_data = payload["zones"]["GHQ"]

    donut_chart = None
    for shape in slide_b.shapes:
        if shape.has_chart and shape.chart.chart_type == XL_CHART_TYPE.DOUGHNUT:
            donut_chart = shape.chart
            break

    assert donut_chart is not None, "Doughnut chart not found on slide B"

    categories = list(donut_chart.plots[0].categories)
    assert categories == ["Treated", "Open"], f"Expected ['Treated', 'Open'], got {categories}"

    values = donut_chart.plots[0].series[0].values
    treated = zone_data["treated"]
    total   = zone_data["total"]
    expected_values = (treated, total - treated)
    assert values == expected_values, f"Expected {expected_values}, got {values}"

    points = donut_chart.plots[0].series[0].points
    color_0 = str(points[0].format.fill.fore_color.rgb)
    assert color_0 == "34D399", f"Treated color: expected '34D399', got '{color_0}'"

    # Open color now muted gold C9A227
    color_1 = str(points[1].format.fill.fore_color.rgb)
    assert color_1 == "C9A227", f"Open color: expected 'C9A227', got '{color_1}'"


# ---------------------------------------------------------------------------
# Zero-total zone: both slides render without crash, no charts on slide A
# ---------------------------------------------------------------------------
def test_zero_total_zone_no_crash():
    """Regression: zero-total zone should not crash; both slides still created."""
    payload = _payload()
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
        "untreated": 0, "aged_count": 0, "avg_age": 0, "median_age": 0,
        "oldest_age": 0, "age_by_domain": [], "age_buckets": [],
    }

    b = ppt_export.deck_to_bytes(payload)
    prs = Presentation(io.BytesIO(b))

    assert len(prs.slides) == 19   # title + 9 zones × 2

    # GHQ slide A: no chart (no domains)
    ghq_a = prs.slides[1]
    assert not any(sh.has_chart for sh in ghq_a.shapes), \
        "Zero-total slide A should have no chart"

    # GHQ slide B: no donut (total=0), aging bar replaced by textbox — no crash
    # (just ensure it rendered)
    ghq_b = prs.slides[2]
    assert any(sh.has_text_frame for sh in ghq_b.shapes), \
        "Slide B should have at least one text frame"


# ---------------------------------------------------------------------------
# Table dimensions
# ---------------------------------------------------------------------------
def test_table_dimensions():
    payload = _payload()
    b = ppt_export.deck_to_bytes(payload)
    prs = Presentation(io.BytesIO(b))

    # Table is on slide A (index 1)
    first_zone_slide = prs.slides[1]
    tbl = next(sh.table for sh in first_zone_slide.shapes if sh.has_table)

    expected_rows = 3   # header + 2 domains
    expected_cols = 4   # Domain + COMMERCIAL + NCI + Total

    assert len(tbl.rows) == expected_rows, f"Expected {expected_rows} rows, got {len(tbl.rows)}"
    assert len(tbl.columns) == expected_cols, f"Expected {expected_cols} cols, got {len(tbl.columns)}"


# ---------------------------------------------------------------------------
# _fit_domains cap
# ---------------------------------------------------------------------------
def test_many_domains_table_capped():
    p = _payload()
    z = p["zones"]["GHQ"]
    doms = [[f"Domain{i}", 10 - i] for i in range(10)]
    z["by_domain"] = doms
    z["domain_pct"] = {d: 10 for d, _ in doms}
    z["crosstab"] = {d: {"COMMERCIAL": n} for d, n in doms}
    z["total"] = sum(n for _, n in doms)
    prs = Presentation(io.BytesIO(ppt_export.deck_to_bytes(p)))
    slide = prs.slides[1]   # GHQ slide A
    tbl = next(sh.table for sh in slide.shapes if sh.has_table)
    assert len(tbl.rows) <= 9
    texts = [tbl.cell(r, 0).text for r in range(len(tbl.rows))]
    assert any(t.startswith("Other") for t in texts)


# ---------------------------------------------------------------------------
# Labels and bounds (both zone slides)
# ---------------------------------------------------------------------------
def test_zone_slides_labeled_and_fit():
    prs = Presentation(io.BytesIO(ppt_export.deck_to_bytes(_payload())))
    W = Emu(int(12192000))   # 13.333 in
    H = Emu(int(6858000))    # 7.5 in

    for idx in [1, 2]:   # slide A and slide B for GHQ
        slide = prs.slides[idx]
        texts = " ".join(
            sh.text_frame.text for sh in slide.shapes if sh.has_text_frame
        )
        assert "GHQ" in texts, f"Slide {idx} missing 'GHQ'"
        for sh in slide.shapes:
            assert sh.left >= 0 and sh.top >= 0, \
                f"Slide {idx}: shape off top/left edge"
            assert sh.left + sh.width <= W + Emu(1), \
                f"Slide {idx}: shape too wide"
            assert sh.top + sh.height <= H + Emu(1), \
                f"Slide {idx}: shape too tall"


# ---------------------------------------------------------------------------
# Dark background on all slides
# ---------------------------------------------------------------------------
def test_dark_theme_background():
    from pptx.enum.dml import MSO_FILL
    prs = Presentation(io.BytesIO(ppt_export.deck_to_bytes(_payload())))
    for slide in prs.slides:
        fill = slide.background.fill
        assert fill.type == MSO_FILL.SOLID
        assert str(fill.fore_color.rgb) == "1B1A19"


# ---------------------------------------------------------------------------
# Aging text: KPI strip on slide A; aging detail lines on slide B
# ---------------------------------------------------------------------------
def test_slide_a_has_avg_untreated_kpi():
    prs = Presentation(io.BytesIO(ppt_export.deck_to_bytes(_payload())))
    slide_a = prs.slides[1]   # GHQ slide A
    texts = " ".join(sh.text_frame.text for sh in slide_a.shapes if sh.has_text_frame)
    assert "Avg untreated age 90d" in texts


def test_slide_b_has_aging_detail_text():
    prs = Presentation(io.BytesIO(ppt_export.deck_to_bytes(_payload())))
    slide_b = prs.slides[2]   # GHQ slide B
    texts = " ".join(sh.text_frame.text for sh in slide_b.shapes if sh.has_text_frame)
    # Key Insights on slide B must contain the aging summary line
    assert "Untreated avg" in texts and "90d" in texts
    assert "d avg" in texts   # per-domain aging lines
