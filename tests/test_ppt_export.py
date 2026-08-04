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
