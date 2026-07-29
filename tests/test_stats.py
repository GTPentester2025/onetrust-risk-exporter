# tests/test_stats.py
import csv
from zone_reports.stats import (
    compute_zone_stats, select_top_cats, CAT_ORDER, ZONES,
)

# The MAZ cross-tab from the reference workbook (rows=domain, cols=Cat).
MAZ = {
    "Business Continuity Management": {"COMMERCIAL":14,"LOGISTICS":1,"NCI":9,"TECHNOLOGY":2,"RAU":5,"Fees":1},
    "Information Security":           {"COMMERCIAL":4,"LOGISTICS":4,"PACKAGING":3},
    "Privacy":                        {"COMMERCIAL":2,"LOGISTICS":2,"NCI":3},
    "Security":                       {"COMMERCIAL":7,"LOGISTICS":1,"NCI":6,"PACKAGING":2},
}

def _maz_rows():
    rows, n = [], 0
    for domain, cats in MAZ.items():
        for cat, count in cats.items():
            for _ in range(count):
                n += 1
                rows.append({"ID": f"R{n}", "Organization": "Middle America Zone",
                             "Category": domain, "Cat": cat, "Stage": "Open"})
    # add noise from another zone that must be filtered out
    rows.append({"ID":"X1","Organization":"Africa","Category":"Privacy","Cat":"NCI","Stage":"Open"})
    return rows

def test_total_and_grand_by_cat():
    s = compute_zone_stats(_maz_rows(), "Middle America Zone")
    assert s.total == 66
    assert s.by_cat == {"COMMERCIAL":27,"LOGISTICS":8,"NCI":18,"PACKAGING":5,
                        "TECHNOLOGY":2,"RAU":5,"Fees":1}

def test_domain_percentages():
    s = compute_zone_stats(_maz_rows(), "Middle America Zone")
    assert s.domain_pct["Business Continuity Management"] == 48
    assert s.domain_pct["Security"] == 24
    assert s.domain_pct["Information Security"] == 17
    assert s.domain_pct["Privacy"] == 11
    # domains sorted descending by count
    assert [d for d, _ in s.by_domain][0] == "Business Continuity Management"

def test_top_cats_match_reference_narrative():
    s = compute_zone_stats(_maz_rows(), "Middle America Zone")
    assert s.top_cats["Business Continuity Management"] == [("COMMERCIAL",14),("NCI",9),("RAU",5)]
    assert s.top_cats["Security"] == [("COMMERCIAL",7),("NCI",6)]
    assert s.top_cats["Information Security"] == [("COMMERCIAL",4),("LOGISTICS",4),("PACKAGING",3)]
    assert s.top_cats["Privacy"] == [("NCI",3),("COMMERCIAL",2),("LOGISTICS",2)]

def test_overall_includes_all_orgs():
    s = compute_zone_stats(_maz_rows(), None)
    assert s.total == 67  # 66 MAZ + 1 Africa noise row

def test_select_top_cats_rule():
    # accumulate desc until cumulative >= 80% of total, min 2, max 4
    assert select_top_cats([("A",7),("B",6),("C",2),("D",1)], 16) == [("A",7),("B",6)]
    assert select_top_cats([("A",14),("B",9),("C",5),("D",2)], 32) == [("A",14),("B",9),("C",5)]
