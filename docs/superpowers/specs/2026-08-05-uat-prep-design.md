# UAT Prep — Config Privacy, File Uploads, Key Insights, PPT & Chart Polish

**Date:** 2026-08-05
**Status:** Approved

## Purpose

Six improvements before UAT, keeping everything else (zone tabs, exports,
scheduling, async fetch, 50-test suite) intact:

1. Config popup never displays saved connection values — status only.
2. Settings can upload the supplier-category Excel and `views.json`; the Excel
   is downloadable for updates; both show "Configured" status only.
3. The narrative becomes a first-class **Key Insights** panel on the dashboard.
4. PPT deck restyled to the UI's dark+gold theme with fixed-fit layout and
   clearly labeled zone slides.
5. Charts get a distinct non-repeating palette and rendered data labels.
6. Crosstab table is contained inside its tile (no width escape).

## Component 1 — Config privacy (server + modal)

Server (`dashboard.py` + `config_store.py`):
- `mask_config` gains a stricter variant used by `GET /api/config`:
  returns `{configured: bool, host_hint: str, has_secret: bool, schedule: {...}}`
  where `configured = is_configured(cfg)` and `host_hint` is a redacted
  hostname — first 2 chars + `…` + the last dot-suffix (e.g. `ap….onetrust.com`),
  empty string when unset. Raw `hostname` and `client_id` are **no longer
  returned** to the browser.
- `POST /api/config` semantics extended: hostname/client_id keep the existing
  "always overwrite with sent value" rule, but the UI only sends them when the
  user is actively editing (Edit mode), so saved values are never round-tripped.

UI (Settings modal):
- Connection section has two states:
  - **Configured view** (when `configured`): a status row
    "● Configured — <host_hint>" (green dot) + **Edit** button. No input fields
    visible.
  - **Edit view** (initial state when unconfigured, or after Edit click):
    empty `#cfgHost`, `#cfgId`, `#cfgSecret` fields (never prefilled — not even
    hostname), Test Connection, and Save. Cancel returns to the status row
    without saving.
- Save sends only the fields the user filled; blank secret still keeps the
  saved one; blank host/id while in Edit mode DO overwrite (explicit clearing
  stays possible).
- Schedule section unchanged (its values are not sensitive and remain shown).

## Component 2 — Data-file uploads (Excel + views.json)

Server routes:
- `GET /api/files` → `{catfile: {present: bool, suppliers: int|null},
  views: {present: bool, view_count: int|null}}`. `suppliers` = row count of the
  lookup (best-effort via openpyxl read; null if unreadable). `view_count` from
  `views.json`.
- `POST /api/upload/catfile` — raw request body = the `.xlsx` bytes
  (`Content-Type: application/octet-stream`). Validates: non-empty, starts with
  the ZIP magic `PK` (xlsx is a zip), size ≤ 10 MB. Writes atomically to
  `Supplier_Category_List.xlsx` (temp file + replace). Returns the new
  `/api/files` payload. 400 on validation failure.
- `POST /api/upload/views` — raw body = JSON bytes. Validates it parses as JSON
  with a `views` list; size ≤ 2 MB. Writes atomically to `views.json`. Returns
  `/api/files` payload. 400 on invalid JSON/shape.
- `GET /api/catfile` — streams `Supplier_Category_List.xlsx` with
  `Content-Disposition: attachment` (404 JSON error if absent). Enables the
  update round-trip (download → edit → re-upload).
- `views.json` is intentionally **not** downloadable (contains org UUID
  filters; users don't need it back).

UI (Settings modal, new "Data files" section):
- Two rows, each: label, status ("● Configured — 214 suppliers" / "● Configured
  — 1 view" / "Not uploaded"), and buttons:
  - Supplier categories: **Upload** (file input accepting `.xlsx`) +
    **Download** (only when present).
  - Views config: **Upload** (accepting `.json`).
- Upload posts the file bytes, refreshes the status row, toasts
  "Uploaded" / error message. File contents are never rendered.

`.gitignore` already covers `views.json`; add `Supplier_Category_List.xlsx`.

## Component 3 — Key Insights panel (dashboard)

Replace the raw-text narrative tile with a full-width **Key Insights** panel
(placed above the table row), built from the zone's existing stats (no new
server data):

- **Headline:** "A total of **N** risks identified across business units" with
  N emphasized in gold.
- **Concentration callout** (when ≥2 cats): "<Cat1> (n₁) and <Cat2> (n₂)
  together account for **P%** of total exposure".
- **Domain insight cards** (one per domain, desc): domain name (gold), a
  `%`-of-total badge, risk count, and the top cats as small chips
  ("COMMERCIAL · 14"). Zero-risk zone → single muted "No risks recorded" line.
- Copy/download buttons keep working: copy = plain-text narrative (existing
  `narrative` lines), download = same `.txt`. The visual panel and the text
  export both derive from the same stats, so they agree.

## Component 4 — PPT restyle (dark + gold, fixed fit, labeled slides)

`ppt_export.py` rewrite of the visual layer (same data, same
`build_payload` source):

- **Theme constants:** slide bg `#1b1a19`; title text gold `#E8C810`; body text
  cream `#F3F2F1`; muted `#C8C6C4`; tile-ish fills `#252423`; series palette =
  the UI chart palette (below). Every slide gets a full-bleed dark background
  rectangle (or slide background fill) — no white slides.
- **Title slide:** product title gold, view + generated date muted, thin gold
  rule.
- **Zone slides:** header band across the top: "**<ZONE> — Risk Insights**"
  (gold) + "N risks · till <date>" (muted). Footer: zone code + slide number.
- **Fixed layout grid (13.33×7.5in):** header band (0.9in) · KPI strip
  (0.5in) · charts row (pie left ~5.9in wide, stacked bar right ~6.4in, 3.1in
  tall) · bottom row (table left ~7.6in, key-insights text right ~4.7in,
  2.4in tall). All shapes positioned within these regions; nothing overlaps or
  overflows the slide.
- **Charts:** series colors from the palette (chart format colors set via
  python-pptx `series.format.fill`); legends on; **data labels on** (pie:
  value + percent; bar: values, hidden for zeros); label fonts sized to fit
  (8–10pt); chart text cream on dark.
- **Table fit:** columns sized to the region; font auto-steps down (10→8pt) by
  column count; if domains > 8 rows, keep top 7 by count + a final "Other (K)"
  aggregate row so the table never exceeds its region. Header row gold-on-dark
  fill; body rows alternate `#252423`/`#2F2E2D`; cream text.
- **Narrative box:** retitled "Key Insights", gold heading, cream 9pt body,
  word-wrapped within its region; text truncated with "…" past the region's
  capacity (full text remains in the dashboard/txt export).
- Charts/tables remain **native editable objects** — only styling changes.

## Component 5 — Chart detail (dashboard)

- **One shared 12-color palette** (colorblind-aware, distinct on dark):
  `#E8C810 #4FC3F7 #F0736A #34D399 #CE93D8 #FFB74D #90CAF9 #F48FB1 #AED581
  #FFD54F #4DD0E1 #BCAAA4`. Cats map to palette slots in union-column order;
  domains map in `by_domain` order; the `(blank)`/Uncategorized slot uses
  `#797775` gray. No two active series share a color (palette length ≥ any
  realistic series count; cycles only past 12).
- **Data labels:** add `chartjs-plugin-datalabels` via CDN. Pie: percent
  (≥5% slices) in dark text on light slices / light on dark, 11px bold;
  Bar: total count above each stacked column; per-segment labels for segments
  tall enough (≥12px), hidden otherwise. Respect `prefers-reduced-motion`
  (labels are static anyway).
- **Tooltips:** pie shows "domain: n (P%)"; bar shows "cat: n (of domain
  total T)".
- Legend: single legend per chart, generated labels use `catLabel`
  (no raw "(blank)").

## Component 6 — Table containment

- `#table-panel` returns to `overflow: hidden` (tile clips at radius).
- `.table-scroll` becomes the single scroll container: `overflow: auto`,
  `max-height: 420px`; sticky `thead` sticks within `.table-scroll` (works —
  the sticky ancestor chain no longer crosses an overflow-hidden element
  between thead and its scroller).
- Column count no longer stretches the tile: table gets `min-width:
  max-content` inside the scroller so wide crosstabs scroll horizontally
  within the tile instead of escaping it.

## Error handling

| Condition | Result |
|--|--|
| Upload not a valid xlsx/json | 400 `{error}`; modal shows red inline message. |
| Catfile download when absent | 404 `{error}`; toast. |
| Upload while fetch job running | allowed (files only read at next fetch). |
| GET /api/config on old clients | new shape; UI ships in same commit. |
| PPT with >8 domains | table aggregates to "Other (K)"; charts unaffected. |

## Security

- `GET /api/config` no longer exposes raw hostname/client_id (only
  `host_hint` + booleans). Secret handling unchanged (never sent/logged).
- Upload routes accept only the two fixed target filenames — the client never
  supplies a path. Size caps + content validation. Atomic writes.
- Server stays `127.0.0.1`; uploads are local-only.
- `.gitignore` adds `Supplier_Category_List.xlsx`.

## Testing

- `config_store`/`dashboard`: `strict_mask` (no hostname/client_id leak,
  host_hint format), `/api/files` shape, upload validation (bad magic → 400,
  valid xlsx bytes → file written, valid/invalid views.json), catfile download
  404/200 headers.
- `ppt_export`: existing 4 tests keep passing (slide count, chart+table,
  zero-zone, dimensions) + new: >8-domain payload → table rows ≤ 9 incl.
  "Other"; a zone slide title contains the zone code; theme smoke (slide bg
  shape/fill present).
- UI: headless smoke (ids present, datalabels CDN, no raw hostname in
  /api/config response) + `node --check`; visual checks manual.

## Out of scope

- Multi-file/version management for uploads (single fixed filename each).
- PPT template (.potx) support.
- Editing supplier categories in the browser.
