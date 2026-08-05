import io
import pytest

pptx = pytest.importorskip("pptx")  # skip cleanly if python-pptx not installed
from pptx import Presentation
import ppt_export

def _payload():
    zones = {}
    for i, code in enumerate(["GHQ", "AFR", "SAZ", "MAZ", "NAZ", "APAC",
                              "EUR", "BEES", "BEES-FT", "Overall"]):
        zones[code] = {
            "organization": code, "total": 3,
            "by_cat": {"COMMERCIAL": 2, "NCI": 1},
            "grand_by_cat": [["COMMERCIAL", 2], ["NCI", 1]],
            "by_domain": [["Security", 2], ["Privacy", 1]],
            "crosstab": {"Security": {"COMMERCIAL": 2}, "Privacy": {"NCI": 1}},
            "domain_pct": {"Security": 67, "Privacy": 33},
            "top_cats": {"Security": [["COMMERCIAL", 2]]},
            "narrative": ["A total of 3 risks identified across business units"],
        }
    return {"view": "TPRM Global View", "generated_at": "2026-08-05T06:00:00",
            "cat_order": ["COMMERCIAL", "LOGISTICS", "NCI", "PACKAGING",
                          "TECHNOLOGY", "RAU", "Fees"], "zones": zones}

def test_deck_has_title_plus_one_slide_per_zone():
    b = ppt_export.deck_to_bytes(_payload())
    prs = Presentation(io.BytesIO(b))
    assert len(prs.slides) == 11  # title + 10 zones

def test_zone_slide_has_chart_and_table():
    prs = Presentation(io.BytesIO(ppt_export.deck_to_bytes(_payload())))
    zone_slide = prs.slides[1]
    has_chart = any(sh.has_chart for sh in zone_slide.shapes)
    has_table = any(sh.has_table for sh in zone_slide.shapes)
    assert has_chart and has_table

def test_zero_total_zone_no_crash():
    """Regression: zero-total zone should not crash and should have no chart."""
    payload = _payload()
    # Override GHQ zone to have zero total
    payload["zones"]["GHQ"] = {
        "organization": "GHQ",
        "total": 0,
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

    # Assert slide count is 11 (title + 10 zones)
    assert len(prs.slides) == 11

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
