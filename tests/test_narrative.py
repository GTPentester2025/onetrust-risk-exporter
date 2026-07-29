from datetime import date
from zone_reports.stats import compute_zone_stats
from zone_reports.narrative import build_narrative, build_title, format_date_ordinal
from tests.test_stats import _maz_rows

def test_format_date_ordinal():
    assert format_date_ordinal(date(2026, 5, 14)) == "14th May 2026"
    assert format_date_ordinal(date(2026, 5, 1)) == "1st May 2026"
    assert format_date_ordinal(date(2026, 5, 22)) == "22nd May 2026"
    assert format_date_ordinal(date(2026, 5, 3)) == "3rd May 2026"

def test_title():
    assert build_title("Middle America Zone", "14th May 2026") == \
        "Middle America Zone Risk Analysis (Till 14th May 2026)"

def test_narrative_key_lines():
    s = compute_zone_stats(_maz_rows(), "Middle America Zone")
    text = "\n".join(build_narrative(s, "14th May 2026"))
    assert "A total of 66 risks identified across business units" in text
    assert "Commercial (27 risks) and NCI (18 risks) together account for 68% of total exposure" in text
    assert "Business Continuity Management (48% of total risk" in text
    assert "Security (24% of total risk)" in text
    # top cats appear with title-cased names and counts
    assert "Commercial (14)" in text
    assert "NCI (9)" in text
