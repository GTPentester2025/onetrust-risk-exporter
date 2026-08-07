# GRO Merge, Risk-Treatment Analytics, Insight Parity & Glass Polish

**Date:** 2026-08-07
**Status:** Approved

## Purpose

Four connected changes, all fed by the existing `refined.csv` → `stats` → payload
pipeline so the dashboard and PPT stay in agreement:

1. **GRO zone** — merge `BEES` + `BEES | FINTECH` into one zone "GRO".
2. **Risk Treatment** — derive "treated" risks (Stage = Monitoring), surface a
   KPI card + a Treated-vs-Open donut on the dashboard and each PPT zone slide.
3. **Insight parity** — the dashboard Key Insights match/exceed the PPT write-up
   (title-cased names, treated line, full per-domain detail).
4. **Glassmorphism polish** — a classy, professional frosted-glass restyle that
   holds WCAG AA legibility.

## Component 1 — GRO zone (stats.py)

- `ZONES` becomes (order matters — drives tab order and PPT slide order):

  ```python
  ZONES = {
      "GHQ": "GHQ", "AFR": "Africa", "SAZ": "South America Zone",
      "MAZ": "Middle America Zone", "NAZ": "North America Zone",
      "APAC": "APAC", "EUR": "Europe",
      "GRO": ["BEES", "BEES | FINTECH"],
      "Overall": None,
  }
  ```

  `BEES` and `BEES-FT` entries are removed. 9 codes incl. Overall.
- `compute_zone_stats(rows, organization=None)` accepts `str | list | None`:
  - `None` → all rows.
  - `list` → rows whose `Organization` is in the list (set membership).
  - `str` → equality (unchanged behavior).
  The filter is the only change; all downstream computation is identical.

## Component 2 — Treated risks (stats.py, narrative.py, dashboard.py)

`stats.py`:
- `TREATED_STAGES = {"monitoring"}` (lower-cased comparison set).
- `is_treated(stage) -> bool`: `(stage or "").strip().lower() in TREATED_STAGES`.
- `ZoneStats` gains: `treated: int`, `treated_pct: int` (percent of total, 0 when
  total 0), `by_stage: list[(stage, count)]` sorted desc then name.
- `compute_zone_stats` fills them: iterate the selected rows, count treated and
  per-stage; blank stage bins under "(blank)".

`narrative.py` — `build_narrative` adds, right after the concentration line
(only when total > 0):

```
{treated} of {total} risks ({treated_pct}%) are actively monitored (treated)
```

This single source feeds the PPT narrative box, the dashboard text/insights, and
the copy/`.txt` export.

`dashboard.py` — `zone_to_dict` adds `treated`, `treated_pct`, `by_stage` to each
zone entry. Payload shape otherwise unchanged.

## Component 3 — Insight parity (index.html)

The dashboard Key Insights panel renders from the zone stats and matches the PPT
write-up exactly in content:

- Title-case domain and Cat names via a JS port of `_disp` (`dispName`):
  keep acronyms `{NCI, RAU, GHQ, BEES, GRO}` upper, else Title Case.
- Headline (gold total), intro line, concentration line (top-2 by count with
  "N risks"), **treated line** (from Component 2), "Key Insights by Risk Domain"
  heading, then one block per domain (desc): "Domain (P% of total risk –
  highest)", "N risks, <rank phrase>:", and ALL `top_cats` as chips
  "Commercial (14)".
- Zero-total zone → "No risks recorded for this zone."
- Copy/download still export `z.narrative` (now includes the treated line).

## Component 4 — Risk Treatment tiles (index.html + ppt_export.py)

Dashboard:
- **KPI card "Risks Treated"** in the KPI row: big number `treated`, sublabel
  `of {total} · {treated_pct}%`, green accent (`--live #34d399`).
- **"Risk Treatment" tile** with a **Treated-vs-Open doughnut** (Chart.js
  `doughnut`): slices Treated (`#34D399`) and Open (`#E8C810` gold),
  `total-treated` = open. Data labels (count + %), legend bottom, tooltip
  "Treated: n (P%)". Destroy-before-recreate like the other charts. Empty state
  when total 0.

PPT (`ppt_export.py`) — each zone slide gains the treated donut; layout moves to
a **3-chart row** so nothing overflows the 13.333×7.5in slide:
- Header band (top ~0.9in) + KPI strip (~1.0in) — KPI strip text also states
  "Treated N (P%)".
- Charts row at y≈1.6, height≈3.0: **pie** (x0.4, w4.3) · **stacked bar**
  (x4.9, w4.6) · **treated donut** (x9.7, w3.2). Donut = native pie chart with
  two points Treated/Open, palette green/gold, % labels.
- Bottom row at y≈4.9, height≈2.3: **table** (x0.4, w7.6) · **Key Insights**
  (x8.3, w4.6).
- Footer unchanged. `_fit_domains` cap still applied to bar + table.

## Component 5 — Glassmorphism polish (index.html)

Single dark theme, restyled to frosted glass, WCAG AA preserved.

- **Ambient base:** body background = gradient `#14100a`→`#0a0806` plus 2–3
  static blurred radial "blobs" (gold + teal, `opacity ≤0.09`) via fixed
  pseudo-elements; no pure black; `prefers-reduced-motion` safe (blobs static).
- **Glass surfaces** (KPI cards, chart/insight/table tiles, header bar, settings
  modal): `background: rgba(255,255,255,0.05)`, `backdrop-filter: blur(14px)
  saturate(1.2)` (+ `-webkit-`), `border: 1px solid rgba(255,255,255,0.09)`,
  inset top highlight `inset 0 1px 0 rgba(255,255,255,0.06)`, `border-radius:
  16px`, one soft shadow `0 8px 24px rgba(0,0,0,0.45)`.
- **Legibility:** body text cream `#f3f2f1` / muted `#c8c6c4` at ≥4.5:1; each
  chart canvas + the table sit on an inner panel `rgba(20,16,10,0.55)` (not
  transparent) so data stays crisp over blur; a faint scrim behind KPI numbers.
  Gold `#e8c810` used only for accents/active state/numbers.
- **Chrome:** unified glass tile header (title + hover toolbar); zone tab strip
  with a single sliding gold underline for the active tab; `tabular-nums` +
  thousands separators (`toLocaleString`) on all figures; consistent 8px spacing
  rhythm; visible `:focus-visible` gold ring; polished loading (spinner/skeleton)
  / empty / error states; min-heights so tiles don't reflow between zones.
- **Perf guard:** `backdrop-filter` only on top-level tiles/chrome — never nested
  — to avoid frame drops (per 2026 glass guidance).

## Data flow

`refined.csv` (has Stage + Organization) → `stats.compute_zone_stats`
(GRO membership, treated, by_stage) → `build_payload` → `index.html`
(tabs, KPI incl. treated, pie/bar/treated-donut, insights) and
`ppt_export.build_deck` (3-chart slides). One computation, identical numbers.

## Error handling

| Condition | Result |
|--|--|
| Row with blank Stage | counts toward total + Open; `by_stage` "(blank)". |
| Zone with 0 rows (GRO or any) | KPIs 0; donut empty state; narrative "0 risks". |
| Stage never "Monitoring" | treated 0 (0%); donut all-Open; valid. |
| GRO orgs absent in data | GRO total 0; still a tab/slide. |

## Security

No new surfaces; no new inputs. Same 127.0.0.1 bind, masked config, gitignored
artifacts. `backdrop-filter` is presentational only.

## Testing

- `stats`: GRO total == BEES + FINTECH rows (and excludes others);
  `compute_zone_stats` accepts a list; `is_treated` truth table; `treated`/
  `treated_pct`/`by_stage` on a fixture with mixed stages; Overall unaffected.
- `narrative`: treated line present with correct numbers; absent when total 0.
- `dashboard` payload: zone-set == {GHQ,AFR,SAZ,MAZ,NAZ,APAC,EUR,GRO,Overall};
  each zone dict has treated/treated_pct/by_stage; GRO total correct from a
  fixture. Update tests asserting BEES/BEES-FT.
- `ppt_export`: slide count == title + 9 == 10 (update the old 11);
  a zone slide has ≥3 chart shapes (pie, bar, donut) + a table; all shapes
  within slide bounds (extend the existing bounds test); dark bg intact.
- UI: `node --check`; headless smoke — tabs list contains GRO and not BEES/
  BEES-FT; `backdrop-filter` present; "Risks Treated" + treated donut ids
  present; `/api/data` renders. Visual (glass, contrast) verified manually with
  the dummy `refined.csv`.

## Out of scope

- Legacy Excel `build_zone_reports.py` will not gain GRO multi-org pivot-filter
  support (the web dashboard is the active product); it will be guarded so a
  list-valued ZONES entry does not crash its import, but a GRO Excel sheet is
  not produced. Documented, not implemented.
- Treated-stage set stays code-level `{"Monitoring"}` (not Settings-editable).
- No new animation beyond static ambient blobs.
