import datetime
from dashboard import build_payload

def _rows():
    rows = []
    n = 0
    for cat in ("COMMERCIAL", "NCI", "COMMERCIAL"):
        n += 1
        rows.append({"ID": f"R{n}", "Organization": "GHQ",
                     "Category": "Security", "Cat": cat, "Stage": "Open"})
    rows.append({"ID": "A1", "Organization": "Africa",
                 "Category": "Privacy", "Cat": "NCI", "Stage": "Open"})
    return rows

def test_payload_shape_and_zone_totals():
    p = build_payload(_rows(), view="V",
                      generated_at=datetime.datetime(2026, 8, 4, 12, 0, 0))
    assert p["view"] == "V"
    assert p["generated_at"] == "2026-08-04T12:00:00"
    assert p["cat_order"] == ["COMMERCIAL", "LOGISTICS", "NCI",
                              "PACKAGING", "TECHNOLOGY", "RAU", "Fees"]
    # every zone code present, incl. Overall
    assert set(p["zones"]) == {"GHQ", "AFR", "SAZ", "MAZ", "NAZ",
                               "APAC", "EUR", "BEES", "BEES-FT", "Overall"}
    ghq = p["zones"]["GHQ"]
    assert ghq["total"] == 3
    assert ghq["by_cat"] == {"COMMERCIAL": 2, "NCI": 1}
    assert p["zones"]["Overall"]["total"] == 4      # includes Africa row
    assert p["zones"]["AFR"]["total"] == 1
    assert ghq["narrative"][0] == "A total of 3 risks identified across business units"

def test_zero_row_zone_is_present():
    p = build_payload(_rows(), view="V",
                      generated_at=datetime.datetime(2026, 8, 4, 12, 0, 0))
    naz = p["zones"]["NAZ"]
    assert naz["total"] == 0
    assert naz["narrative"] == ["A total of 0 risks identified across business units"]
