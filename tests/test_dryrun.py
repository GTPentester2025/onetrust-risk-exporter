from build_zone_reports import iter_zone_reports, zone_display_name
from tests.test_stats import _maz_rows

def test_zone_display_name():
    assert zone_display_name("MAZ", "Middle America Zone") == "Middle America Zone"
    assert zone_display_name("Overall", None) == "Overall"

def test_iter_zone_reports_covers_all_zones():
    reports = iter_zone_reports(_maz_rows(), "14th May 2026")
    sheets = {r["sheet"] for r in reports}
    assert {"GHQ","AFR","SAZ","MAZ","NAZ","APAC","EUR","BEES","BEES-FT","Overall"} <= sheets
    maz = next(r for r in reports if r["sheet"] == "MAZ")
    assert maz["stats"].total == 66
    assert maz["title"] == "Middle America Zone Risk Analysis (Till 14th May 2026)"
    overall = next(r for r in reports if r["sheet"] == "Overall")
    assert overall["stats"].total == 67
