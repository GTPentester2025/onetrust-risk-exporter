# Legibility Polish Review — index.html
**Diff:** `_legibility.diff` | **Built file:** `index.html`

---

## 1. Type Scale

**Status: PASS with one minor note.**

All six tokens declared correctly in `:root` (lines 37–42):

```
--fs-xs:   0.72rem
--fs-sm:   0.80rem
--fs-base: 0.90rem
--fs-md:   1.00rem
--fs-lg:   1.25rem   (reserved, unused — OK)
--fs-kpi:  1.75rem
```

Every previously-hardcoded CSS `font-size` in the stylesheet is replaced with a token. A grep of the built file for `font-size: [digit]` confirms only two survivors:

- **Line 48** — `html { font-size: 14px; }` — root px anchor. Intentional and correctly exempt per spec.
- **Line 632** — `.ki-num { font-size: 1.2em; }` — a relative `em` multiplier on `.ki-headline`. Also intentional; spec allows this.

No broken `var()` references. All token names used in properties match the declared names exactly.

**Minor:** `--fs-lg` is declared but has zero consumers in the current file (comment says "reserved / future"). Not a bug, but any future rule that uses it should verify the value is appropriate for that context before assuming it fits.

---

## 2. Stacked Bar — Top-5 + Other Folding

### 2a. Folding math correctness

**Status: PASS — no data bug.**

The algorithm (lines 1884–1919):

1. Iterates `activeCats` and sums each cat across all domains via `z.crosstab[dom][cat]` — same source as the old per-cat datasets.
2. Sorts descending by grand total; `slice(0, 5)` = `topCats`, `slice(5)` = `otherCats`.
3. Each top-cat dataset reads `z.crosstab[dom][cat]` per domain — identical lookup to pre-patch code.
4. `otherData[dom]` = sum of `z.crosstab[dom][cat]` over every cat in `otherCats`.

**Column-total preservation:** For any domain `d`, the visible column stack equals the sum of top-5 values for `d` plus the sum of other-cat values for `d`, which equals the sum of ALL cats for `d`. Since `activeCats` is exhaustively partitioned into `topCats ∪ otherCats` with no overlap and no omission, column totals are preserved by construction. The `total` datalabel reads from `colTotals[ctx.dataIndex]` (derived from `z.by_domain`, the authoritative figure) independent of the folding, so it will always be correct even if the folding were wrong.

### 2b. Other-series guard (no spurious empty series)

**Status: PASS.** `if (otherCats.length > 0)` (line 1907) ensures "Other" is pushed only when more than 5 active categories exist. With 5 or fewer cats, no "Other" series appears.

### 2c. Legend color budget

**Status: PASS.** Top-5 use `PALETTE[0..4]` (5 distinct colors); Other is `#797775`. At most 6 legend entries. Matches spec.

### 2d. seriesColor collision when (blank) is top-5

**Status: MINOR — pre-existing, not introduced by this diff.**

`seriesColor(idx, cat)` returns `#797775` when `cat === '(blank)'`, which is the same gray as the "Other" series. If `(blank)` is among the top-5 categories, both it and "Other" will show identical gray swatches in the legend. Not a data bug; users can still read the tooltip. Track for a future fix (assign a distinct fallback color for `(blank)` in the top-5 path, e.g., `PALETTE[5]`).

### 2e. Segment datalabel threshold change

**Status: PASS.** Raised from `> 0.06` to `> 0.08` of `maxTotal`. With fewer series (≤6 vs. potentially many), segments are larger; the higher threshold still covers meaningful-sized slices while reducing label clutter on genuinely small segments. Font bumped to size 11 / weight 700, consistent with all other datalabels. OK.

### 2f. Tooltip — Other shows no category breakdown

**Status: IMPORTANT — UX gap, not a data-accuracy bug.**

The tooltip `label` callback (line 1937):
```js
return ' ' + item.dataset.label + ': ' + item.parsed.y + ' (of ' + tot + ')';
```
When hovering the "Other" bar segment, the user sees `Other: 12 (of 45)` but gets no breakdown of which categories are folded in. This is a loss of discoverability compared to the pre-patch chart where every category was individually visible on hover. Recommend adding an `afterLabel` callback that, when `item.dataset.label === 'Other'`, lists each folded category's per-domain count from `otherCats`. Low priority for this polish pass but should be tracked.

### 2g. barCsvData and renderTable not updated — undocumented divergence

**Status: IMPORTANT — intentional behavior, but needs a code comment.**

`barCsvData` (line 2172) and `renderTable` (line 2310) still iterate the full `activeCats` list (all categories, unfolded). The bar chart shows Top-5+Other visually, but the CSV export and table show all categories. This is the correct behavior — exports should be complete — but the bar legend and the table column headers are now intentionally inconsistent. A user cross-referencing the bar to the table will find more columns in the table than legend entries in the chart. Add a comment near `barCsvData` and `renderBar` stating this divergence is deliberate, so a future maintainer does not "sync" the two by accident.

---

## 3. Data-Label Consistency

**Status: PASS.**

All chart datalabels post-patch:

| Chart | Datalabel size / weight | Legend size | Tick size |
|---|---|---|---|
| Pie | 11 / 700 | 11 | n/a |
| Bar (segment) | 11 / 700 | 11 | x: 10, y: 10 |
| Bar (total) | 11 / 700 | — | — |
| Donut (treated) | 11 / 700 | 11 | n/a |
| Aging (horizontal bar) | 11 / 700 | hidden | x: 10, y: 10 |
| Buckets (bar) | 11 / 700 | hidden | x: 10, y: 10 |

Consistent across all six chart instances. No global `ChartDataLabels` default overriding per-chart config — each chart declares its own `datalabels` plugin options in isolation. The aging warm ramp (`ageColor`, lines 1556–1572) is untouched.

---

## 4. Behavior Regression Check

**Status: PASS.**

- **Zone switch — no fetch:** `selectZone` (line 1712) calls `render()` only when `payload` is non-null; no `fetch` call inside it. Confirmed.
- **Charts destroy-before-recreate:** All five render functions call `.destroy()` and null the variable before `new Chart(...)`. Confirmed for `renderPie`, `renderBar`, `renderTreated`, `renderAging`, `renderBuckets`.
- **JS validity:** The new block is standard ES5-compatible (no arrow functions, no `const`/`let`, no template literals). All `var` scoped correctly within `renderBar`. No syntax issues.
- **No data-flow mutation:** `catTotals`, `sortedCats`, `topCats`, `otherCats` are local to `renderBar`. No writes to `z`, `payload`, or any module-level state.
- **Other tiles intact:** KPI, pie, treated, aging, buckets, narrative, and table render functions are unchanged. All CSS for those tiles is unchanged.

---

## Summary of Findings

| Severity | Location | Issue | Recommendation |
|---|---|---|---|
| **Important** | `renderBar` tooltip `label` callback, line 1937 | Hovering "Other" shows only the summed value; no per-category breakdown | Add `afterLabel` listing folded cat names and counts when `item.dataset.label === 'Other'` |
| **Important** | `barCsvData` line 2172 / `renderBar` line 1869 | Bar chart (Top-5+Other) vs. CSV/table (all cats) divergence is undocumented | Add a comment near both functions stating the divergence is deliberate |
| **Minor** | `seriesColor()` + `renderBar` | `(blank)` in top-5 shares gray `#797775` with "Other" legend swatch | Pre-existing; assign `PALETTE[5]` or another distinct color for `(blank)` in the top-5 path |
| **Minor** | `:root`, line 41 | `--fs-lg` declared but unused | No action needed now; verify value is appropriate before first use |

**Quality verdict:** The tokenization is thorough and correct — all hardcoded CSS font-sizes have been replaced, leaving only the intentional `html` base px and the `ki-num` em multiplier. The Top-5+Other folding is mathematically sound: column totals are preserved, no categories are dropped or double-counted, and the guard against a spurious empty "Other" series is correctly in place. The two Important findings are UX/maintainability gaps, not data bugs. The patch is safe to ship with those noted as follow-up items.

---

## Fix Round 1

**Applied:** Three focused legibility improvements to renderBar area (no data-flow changes):

1. **Distinct color for blank category** (Fix #56 priority):
   - Modified `seriesColor()` to use `PALETTE[i % PALETTE.length]` directly for top-5 datasets.
   - Removed special case returning gray `#797775` for `(blank)` categories.
   - Now if `(blank)` is in top-5, it gets a distinct palette color instead of clashing with the gray "Other" series.
   - Only the folded "Other" series uses gray `#797775`.

2. **Other tooltip breakdown** (Fix #65 priority):
   - Added `afterLabel` callback in tooltip options for renderBar.
   - When hovering the "Other" dataset, tooltip now appends the folded category names + per-domain counts.
   - Preserves existing "cat: n (of total)" and footer Total behavior for other series.
   - Builds breakdown from `otherCats` and `z.crosstab` already in scope.

3. **Chart-vs-CSV comment** (Fix #79 priority):
   - Added one-line code comment near `renderBar()` (line 1871): "Chart shows Top-5 categories + Other; CSV export and the crosstab table intentionally keep ALL categories."
   - Duplicated comment near `tableTsv()` (line 2375) for visibility at export point.
   - Clarifies intentional divergence for future maintainers.

**Verification:**
- `node --check` on extracted script: PASS (no syntax errors).
- `grep "Chart shows Top-5"` confirms both comments present.
- `grep "Other"` in index.html: 9 matches (series still exists, fully functional).
- Syntax preserved: all changes are cosmetic + UX (no data-flow mutation).
- Changes isolated to renderBar and tooltip/table comments (other tiles untouched).
