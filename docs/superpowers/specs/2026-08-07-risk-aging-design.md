# Untreated Risk Aging Analytics — Design

**Date:** 2026-08-07
**Status:** Approved

## Purpose

Add data-analyst-grade aging analytics for **untreated** risks (Stage ≠
Monitoring): average age, median (outlier-resistant), oldest, per-domain
average aging, and SLA aging buckets. Age = `today − Date created` in days.
Surfaced on the dashboard (KPI + two charts) and the PPT (compact text), fed
from the existing `refined.csv` → `stats` → payload pipeline so all numbers
agree.

## Component 1 — stats.py

- `parse_date(s) -> date | None`: strip; take the first 10 chars (handles both
  `YYYY-MM-DD` and full ISO `2026-06-15T08:30:00Z`); `date.fromisoformat`;
  return None on blank/garbage.
- `AGE_BUCKETS` order: `["0-30", "31-90", "91-180", "180+"]`.
- `compute_zone_stats(rows, organization=None, today=None)`: `today` defaults
  to `datetime.date.today()`. Over the selected rows, for each **untreated**
  row (`not is_treated(Stage)`) with a parseable `Date created`, compute
  `age = (today - created).days`; **drop** ages `< 0` (future/created-error).
  Compute:
  - `untreated: int` — count of untreated rows (regardless of date validity).
  - `aged_count: int` — untreated rows with a valid, non-negative age.
  - `avg_age: int` — mean of ages rounded (0 when `aged_count == 0`).
  - `median_age: int` — `statistics.median` rounded (0 when none).
  - `oldest_age: int` — max age (0 when none).
  - `age_by_domain: list[(domain, avg_age:int, count:int)]` — mean age of the
    untreated+aged rows per domain, sorted by `avg_age` desc then domain.
  - `age_buckets: list[(label, count)]` in `AGE_BUCKETS` order — bucketed by
    age: 0–30, 31–90, 91–180, 180+.
- New `ZoneStats` fields (defaults so positional construction stays safe):
  `untreated=0, aged_count=0, avg_age=0, median_age=0, oldest_age=0,
  age_by_domain=field(default_factory=list), age_buckets=field(default_factory=list)`.

## Component 2 — narrative.py

In `build_narrative`, after the treated line (when `untreated > 0` and
`aged_count > 0`), append:

```
Untreated risks average {avg_age} days old (median {median_age}); oldest {oldest_age} days
```

## Component 3 — dashboard.py

`build_payload` passes `today=generated_at.date()` to `compute_zone_stats` so
aging is measured against the report timestamp (deterministic, matches
`generated_at`). `zone_to_dict` adds `untreated`, `aged_count`, `avg_age`,
`median_age`, `oldest_age`, `age_by_domain`, `age_buckets`.

## Component 4 — index.html (glass tiles, existing theme)

- **KPI card "Avg Untreated Age":** big `{avg_age}d`; sublabel
  `median {median_age}d · oldest {oldest_age}d · {untreated} untreated`.
  Zero untreated → `0d` and sublabel `no untreated risks`.
- **"Risk Aging by Domain" tile:** Chart.js horizontal bar (`indexAxis:'y'`),
  one bar per `age_by_domain` entry (oldest-first), x = avg days, day-count
  data labels, warm ramp (older = redder: interpolate gold `#E8C810` → orange
  `#FFB74D` → red `#F0736A` by the bar's share of the max age), `dispName` on
  domain labels. Empty state when no aged untreated risks.
- **"Aging Buckets" tile:** Chart.js bar of `age_buckets` counts, fixed
  order/colors: 0-30 `#34D399`, 31-90 `#E8C810`, 91-180 `#FFB74D`, 180+
  `#F0736A`; count data labels. Empty state when all zero.
- Both new charts follow existing conventions: inner `--panel` wrapper,
  destroy-before-recreate (add both instances to the destroy logic), per-chart
  datalabels config, `render*` called from `render()`, min-height on the tiles.
- **Key Insights:** add the aging line (mirrors narrative wording) under the
  treated line.

## Component 5 — ppt_export.py

- KPI strip text appends `  ·  Avg untreated age {avg_age}d` (when untreated>0).
- The Key Insights text box appends, after the existing narrative lines: the
  overall aging line and up to the top 3 `age_by_domain` rows as
  `"{domain}: {avg} d avg ({count})"`. No new chart (slide already has 3).

## Data flow

`Date created` + `Stage` (refined.csv) → `compute_zone_stats` aging (relative to
`generated_at.date()`) → payload → dashboard KPI/charts + narrative + PPT text.
One computation; identical numbers everywhere.

## Error handling

| Condition | Result |
|--|--|
| Blank/garbage `Date created` | row excluded from aging (`aged_count` < untreated). |
| `created` in the future (age<0) | dropped from aging. |
| Zone with 0 untreated (all treated / 0 total) | avg/median/oldest 0; charts empty state; narrative aging line omitted. |
| Untreated exist but none have valid dates | `aged_count`=0; treated as "no aging data" (empty state). |

## Security

No new inputs or routes; presentational + pure computation only. Server stays
127.0.0.1.

## Testing

- `parse_date`: `YYYY-MM-DD`, full ISO with `T…Z`, `""`, `"garbage"`, None.
- Aging on a fixture with a fixed `today` and known created dates + mixed
  stages: `avg_age`/`median_age`/`oldest_age` correct; treated rows excluded;
  a future-dated untreated row dropped (age<0); a blank-date untreated row
  raises `untreated` but not `aged_count`; `age_by_domain` sorted desc; each
  `age_buckets` boundary (30, 31, 90, 91, 180, 181) lands in the right band.
- `narrative`: aging line present with correct numbers; absent when
  untreated/aged 0.
- `dashboard`: payload zone dict has all aging fields; `build_payload` uses
  `generated_at.date()` as `today` (test with a fixed `generated_at`).
- `ppt_export`: KPI strip contains "Avg untreated age"; insights text contains
  the aging line; still 10 slides, shapes within bounds.
- UI: `node --check`; headless greps — aging KPI + both chart canvas ids
  present; `/api/data` still serves.

## Out of scope

- Aging of treated/closed risks; time-to-treat metrics.
- Configurable bucket thresholds (fixed 0-30/31-90/91-180/180+).
- Aging trend over time (needs historical snapshots).
