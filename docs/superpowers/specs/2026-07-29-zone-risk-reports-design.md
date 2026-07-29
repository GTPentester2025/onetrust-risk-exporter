# Zone Risk Reports — Design

**Date:** 2026-07-29
**Status:** Approved (pending spec review)

## Purpose

From the refined 19-column risk data, generate a multi-sheet Excel report: one
sheet per business zone plus one "Overall" sheet, each containing a **live
PivotTable** (Category rows × Cat columns, Count of ID), the native pie and
stacked-bar charts, a title, and an auto-generated narrative — matching the
existing hand-built **MAZ** sheet exactly.

## Pipeline position

```
onetrust_view_export.py   ->  raw risk rows (API)
refine_columns.py         ->  refined.csv (19 cols, incl. Cat)
build_zone_reports.py     ->  workbook with per-zone + Overall sheets   <-- THIS
```

`build_zone_reports.py` does no API calls. It consumes `refined.csv` and an
existing workbook whose **MAZ** sheet is the gold template.

## Runtime & dependencies

- Windows with Excel installed (drives Excel via COM).
- `pywin32`. The script checks `import win32com.client` at startup and, if
  missing, runs `pip install pywin32` automatically before continuing.
- Fails fast with a clear message if Excel is not installed / not COM-registerable.

## Inputs

| Arg | Default | Meaning |
|--|--|--|
| `--workbook` | (required) | `.xlsx` containing the MAZ gold template sheet |
| `--data` | `refined.csv` | the 19-column refined export |
| `--template-sheet` | `MAZ` | sheet to clone for every zone |
| `--out` | overwrite `--workbook` | output path (a `*_backup.xlsx` copy is always written first) |

## Zone map

Sheet code → Organization value (the pivot's Organization page-filter):

| Sheet | Organization |
|--|--|
| GHQ | GHQ |
| AFR | Africa |
| SAZ | South America Zone |
| MAZ | Middle America Zone |
| NAZ | North America Zone |
| APAC | APAC |
| EUR | Europe |
| BEES | BEES |
| BEES-FT | BEES \| FINTECH |

Plus an **Overall** sheet: same layout, **no** Organization filter (all zones).

The map is a module-level dict, easy to edit.

## Pivot definition (all sheets identical except the Organization filter)

- **Source:** an Excel Table `RiskData` on a `Data` sheet (auto-expands).
- **Rows:** `Category` field (values are risk domains: Business Continuity
  Management, Information Security, Privacy, Security, …). Note: in MAZ the row
  header is captioned **"Risk Names"**, but the underlying pivot field is
  `Category` — clones read the field name from the template pivot, not the caption.
- **Columns:** `Cat` (supplier category).
- **Values:** Count of `ID`.
- **Page filters:** `Organization` (set per zone; blank/all on Overall),
  `Stage` (all).
- **Cat column order:** fixed — `COMMERCIAL, LOGISTICS, NCI, PACKAGING,
  TECHNOLOGY, RAU, Fees`. Applied via the pivot field's manual sort/order; any
  Cat value not in the list is appended after, before Grand Total.

## Components

### 1. `stats.py` — pure computation (no Excel, unit-testable)

`compute_zone_stats(rows, organization=None) -> ZoneStats` where rows are the
refined dicts. Filters to the organization (or all), then computes:

- `total` — risk count
- `by_cat` — {Cat: count}, in the fixed Cat order
- `by_domain` — {Category: count}, sorted desc
- `crosstab` — {Category: {Cat: count}} with row/col/grand totals
- `domain_pct` — {Category: percent of total}
- `top_cats_by_domain` — {Category: [(Cat, count), …] desc}

This module reproduces the MAZ numbers (66 total, Commercial 27, NCI 18, BCM
48%, …) and is verified against them in tests **without Excel**.

### 2. `narrative.py` — placeholder templates

A template string with `{placeholders}` mirroring the MAZ text block, filled
from `ZoneStats`. Sketch:

```
{zone} Risk Analysis (Till {date})

A total of {total} risks identified across business units
Risk concentration is uneven, with strong clustering in a few key areas
{top_cat_1} ({top_cat_1_n} risks) and {top_cat_2} ({top_cat_2_n} risks)
together account for {top2_pct}% of total exposure

Key Insights by Risk Domain
{domain_blocks}
```

`{domain_blocks}` is generated per domain (desc):

```
{domain} ({pct}% of total risk{superlative})
{n} risks{concentration_clause}
    {cat}: {cat_n}
    ...
```

Returns plain text lines placed into the narrative cells. Templates live at the
top of the module so wording is editable in one place.

### 3. `build_zone_reports.py` — Excel COM orchestration

1. Ensure pywin32 (auto-install if absent).
2. Load `refined.csv`.
3. Copy `--workbook` to `*_backup.xlsx`.
4. Open Excel (`Visible=False`, `DisplayAlerts=False`), open the workbook.
5. Write `refined.csv` into a `Data` sheet; define/replace Table `RiskData`.
6. Locate the pivot on MAZ; repoint its cache to `RiskData`; refresh; read its
   field names so the clones are configured generically (RowField=Category,
   ColumnField=Cat, PageFields=Organization/Stage, DataField=Count of ID).
7. For each zone (incl. re-doing MAZ for consistency): clone the template sheet,
   rename to the sheet code, set the Organization page-filter, refresh the pivot
   (charts follow automatically).
8. Build the **Overall** sheet the same way with no Organization filter.
9. For every sheet, overwrite the title cell and narrative cells with the
   regenerated text from `narrative.py`.
10. Refresh all, save (`--out`), quit Excel.

Title/narrative cell anchors are discovered by scanning the MAZ sheet once (cell
containing "Risk Analysis" = title; the contiguous text block below = narrative
region) and reused for every clone, so they stay aligned with your layout.

## Error handling

- No Excel / COM error → clear message, exit non-zero, leave workbook untouched
  (backup already made).
- MAZ template sheet missing → abort before any change.
- A zone with zero matching rows → sheet still created; pivot empty; narrative
  states "0 risks".
- Excel process always quit in a `finally` (no orphaned EXCEL.EXE).

## Testing

- `stats.py` and `narrative.py`: unit tests with a small fixture reproducing the
  MAZ figures — run anywhere, no Excel.
- COM orchestration: validated interactively on this machine against a **copy**
  of the real workbook (backup guarantees safety).

## Out of scope

- No API calls (upstream scripts own that).
- No live-refresh scheduling; the pivots are set to refresh-on-open so opening
  the file recomputes if the Data sheet is later edited by hand.
