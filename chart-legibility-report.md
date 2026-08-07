# Chart Legibility Polish — Report

## Fix 1: Type Scale Consolidation

Six CSS custom properties added to `:root`:

```css
--fs-xs:   0.72rem;   /* <=0.72 */
--fs-sm:   0.80rem;   /* 0.73–0.85 */
--fs-base: 0.90rem;   /* 0.86–0.95 */
--fs-md:   1.00rem;   /* 0.96–1.10 */
--fs-lg:   1.25rem;   /* 1.11–1.40 (reserved) */
--fs-kpi:  1.75rem;   /* >=1.50 */
```

### Mapping applied (every hardcoded CSS `font-size` replaced):

| Original value(s) | Token used |
|---|---|
| 0.65rem (sort-indicator) | `var(--fs-xs)` |
| 0.68rem (tile-title, modal-section-label) | `var(--fs-xs)` |
| 0.69rem (kpi-label) | `var(--fs-xs)` |
| 0.70rem (table th, field-group label, settings hint text) | `var(--fs-xs)` |
| 0.72rem (#status, ki-pct-badge, ki-chip) | `var(--fs-xs)` |
| 0.77rem (kpi-sub) | `var(--fs-sm)` |
| 0.78rem (zone-tab, ki-count, filesMsg, inline modal small btns) | `var(--fs-sm)` |
| 0.80rem (ki-section, ki-phrase, cfgTestResult) | `var(--fs-sm)` |
| 0.81rem (#table font-size, toast) | `var(--fs-sm)` |
| 0.82rem (bar-btn, banner-open, menu-item) | `var(--fs-sm)` |
| 0.83rem (files-status, modal-btn) | `var(--fs-sm)` |
| 0.85rem (ki-intro, ki-concentration, conn-label, field inputs) | `var(--fs-sm)` |
| 0.87rem (banner-text, treated-empty, aging-empty, narrative, ki-domain-name, ki-zero) | `var(--fs-base)` |
| 0.88rem (#empty p) | `var(--fs-base)` |
| 0.90rem (settings-head h2) | `var(--fs-base)` |
| 0.95rem (product-mark, ki-headline, zone-heading) | `var(--fs-base)` |
| 1.00rem (banner-dismiss) | `var(--fs-md)` |
| 1.05rem (#empty h2) | `var(--fs-md)` |
| 1.75rem (kpi-value) | `var(--fs-kpi)` |

**Intentionally kept as-is:**
- `html { font-size: 14px }` — root base size
- `.ki-num { font-size: 1.2em }` — em-relative, intentionally ~large vs parent

---

## Fix 2: Top-5 + Other Stacked Bar

`renderBar` now:
1. Computes each category's grand total across all domains (`catTotals`).
2. Sorts `activeCats` by grand total descending.
3. Keeps top 5 (`topCats`), folds the rest into a synthetic `"Other"` series.
4. Colors: `topCats[0..4]` use `seriesColor(idx, cat)` (PALETTE[0..4]); `"Other"` uses `'#797775'` (gray). Max 6 legend colors.
5. Data label threshold raised from `> 0.06` to `> 0.08` (segments are larger now with fewer series).
6. Segment label font: size 11, weight 700, color `#1b1a19`.
7. Column total label: size 11, weight 700, color `#f3f2f1`.
8. Legend size updated to 11.
9. Tooltip for `"Other"` correctly shows its per-domain summed value (computed from folded cats).
10. `catLabel` / `dispName` still applied; `"Other"` label is literal.

---

## Fix 3: Data-Label Consistency

Unified across all charts (per-chart configs, no global bleed):

| Chart | Legend size | Datalabels size/weight | Axis ticks |
|---|---|---|---|
| renderPie | 11 | 11 / 700 | n/a (pie) |
| renderBar | 11 | 11 / 700 | 10 |
| renderTreated | 11 | 11 / 700 | n/a (donut) |
| renderAging | n/a (hidden) | 11 / 700 | 10 |
| renderBuckets | n/a (hidden) | 11 / 700 | 10 |

Aging warm color ramp (`ageColor`) unchanged.

---

## Verification (curl output)

```
# --fs- token count in served HTML
curl -s http://127.0.0.1:8085/ | grep -c -- "--fs-"
53

# Remaining hardcoded font-size values (should be only base + em)
curl -s http://127.0.0.1:8085/ | grep -oE "font-size: *[0-9.]+(rem|px|em)" | sort | uniq -c
      1 font-size: 1.2em
      1 font-size: 14px

# "Other" label present in JS
curl -s http://127.0.0.1:8085/ | grep -c "Other"
4
```

**pytest:** 79 passed, 0 failures.
