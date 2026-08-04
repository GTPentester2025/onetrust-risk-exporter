# Risk Insights Dashboard — Design

**Date:** 2026-08-04
**Status:** Approved

## Purpose

A simple browser dashboard for the OneTrust zone risk data. One command starts a
local server and opens `index.html`. A **Fetch** button pulls fresh risk rows
from the OneTrust API (reusing the existing exporter + refiner), caches them to
`refined.csv`, and renders the same insights the Excel/PPT report shows — KPI
cards, a domain pie, a domain×category stacked bar, an auto narrative, and a
crosstab table — with a **zone filter** to switch between GHQ, AFR, SAZ, MAZ,
NAZ, APAC, EUR, BEES, BEES-FT, and Overall instantly.

## Why a local server (not pure static HTML)

The OneTrust grid endpoint requires an OAuth2 client-credentials token and is
not CORS-enabled for browsers. So the browser cannot call it directly. A tiny
local Python server bridges the gap: the browser posts credentials to
`localhost`, the server runs the existing CLI pipeline, and returns computed
stats as JSON. This also lets the dashboard reuse `stats.py` verbatim (already
unit-tested), so the numbers match the Excel report exactly.

## Pipeline position

```
onetrust_view_export.py  ->  raw risk rows (API)          } reused as-is
refine_columns.py        ->  refined.csv (19 cols + Cat)  } via subprocess
zone_reports/stats.py    ->  per-zone ZoneStats           } imported directly
dashboard.py             ->  serves JSON + index.html     <-- NEW
index.html               ->  renders insights + filters   <-- NEW
```

No changes to existing scripts. Two new files only.

## Components

### 1. `dashboard.py` — local server (stdlib only)

`python dashboard.py [--port 8000]` starts `http.server` bound to `127.0.0.1`,
then opens the default browser at `http://127.0.0.1:8000/`. No new pip
dependencies; uses `http.server`, `subprocess`, `json`, `webbrowser`, and
imports `zone_reports.stats`.

Routes:

| Method | Path | Behavior |
|--|--|--|
| GET | `/` | Serve `index.html`. |
| GET | `/api/views` | Return `{"views": [name, ...]}` read from `views.json`. `{"views": []}` if the file is absent. |
| GET | `/api/data` | If `refined.csv` exists, return the all-zones payload (see below) computed from it, with `source: "cache"`. Else `{"empty": true}`. |
| POST | `/api/fetch` | Body `{hostname, client_id, client_secret, view}`. Run the pipeline, write `refined.csv`, return the all-zones payload with `source: "api"`. On failure return `{"error": "<message>"}` with HTTP 4xx/5xx. |

`/api/fetch` steps:
1. Validate body has `view` + creds; else 400 `{"error": ...}`.
2. `subprocess.run` `onetrust_view_export.py --view <view> --all-columns
   --out raw_export.csv --hostname .. --client-id .. --client-secret ..`.
   Credentials are passed as process args to the child only; never persisted.
3. `subprocess.run` `refine_columns.py --in raw_export.csv --out refined.csv
   --cat-file Supplier_Category_List.xlsx` (cat-file optional; Cat blank if
   absent).
4. Load `refined.csv` with `stats.load_rows`, build the payload, return it.
5. Any non-zero exit or exception → capture stderr tail → `{"error": ...}`.

**All-zones payload** (single JSON, so the zone filter never refetches):

```json
{
  "source": "api" | "cache",
  "generated_at": "2026-08-04T12:00:00",
  "view": "TPRM Global View",
  "cat_order": ["COMMERCIAL","LOGISTICS","NCI","PACKAGING","TECHNOLOGY","RAU","Fees"],
  "zones": {
    "GHQ": { "organization": "GHQ", "total": 66, "by_cat": {...},
             "grand_by_cat": [["COMMERCIAL",27], ...],
             "by_domain": [["Business Continuity Management",32], ...],
             "crosstab": { "domain": { "Cat": n } },
             "domain_pct": { "domain": 48 },
             "top_cats": { "domain": [["NCI",18], ...] },
             "narrative": ["line", ...] },
    "...": { ... }, "Overall": { ... }
  }
}
```

`generated_at` is stamped by the server (Python `datetime`), not the browser.
Each zone entry is `ZoneStats` serialized to a dict plus a `narrative` list from
`narrative.build_narrative`. The server iterates `stats.ZONES` (already the
correct sheet→organization map incl. `Overall: None`).

Errors: missing `views.json` → `/api/views` returns empty (UI shows a hint).
`onetrust_view_export.py` exits (403, bad creds, missing view) → error surfaced
verbatim. No Excel/COM is involved anywhere in the web path.

### 2. `index.html` — single-file UI

Plain HTML + CSS + vanilla JS. Chart.js from CDN for the pie and stacked bar. No
build step.

Regions:
- **Header bar:** product title; inputs for hostname / client id / client
  secret (secret uses `type=password`); a **view** `<select>` populated from
  `/api/views`; a **Fetch** button; a **status pill**.
- **Status pill states:** idle · `Fetching…` (button disabled, spinner) ·
  `Loaded N risks · <time> · <source>` (green) · `Error: <msg>` (red,
  `role="alert"`).
- **Zone filter:** a segmented control listing the ten zones. Selecting one
  re-renders all insight widgets from the cached payload — no network call.
- **Insights (per selected zone):**
  - KPI cards: total risks; top-2 domains combined share %; top supplier
    category (name + count).
  - **Domain pie:** share by `Category` (uses `by_domain`).
  - **Stacked bar:** domain (x) × supplier `Cat` (stacked series), from
    `crosstab` in `cat_order`.
  - **Narrative:** the `narrative` lines rendered as a readable block.
  - **Crosstab table:** rows = domains, columns = cats in `cat_order` + totals;
    tabular figures; client-side column sort.
- **Empty state:** before any data, a centered prompt ("Enter credentials and
  click Fetch") and, if `/api/data` returns cached data on load, render it
  immediately.

On load: `GET /api/views` (fill dropdown) and `GET /api/data` (render cache if
present). On Fetch: `POST /api/fetch`, store payload in a JS variable, default
the zone filter to `Overall`, render.

### Design system (ui-ux-pro-max)

Professional compliance/analytics dashboard: slate/indigo palette, Inter font,
card-based layout on an 8px spacing rhythm, Lucide SVG icons (no emoji), light +
dark via `prefers-color-scheme`. Accessibility: ≥4.5:1 text contrast, focus
rings, chart legends plus the crosstab table as a non-color data path,
`aria-live` status pill. Exact tokens finalized during build via
`search.py --design-system`.

## Security

- Server binds `127.0.0.1` only — not reachable off the machine.
- Client secret is posted to localhost, passed to the child process as args,
  and held only in browser memory. It is **never written to disk**; `refined.csv`
  contains risk rows only, no credentials.
- No secret is logged. Error messages surface stderr but the pipeline does not
  echo the secret.

## Error handling

| Condition | Result |
|--|--|
| `views.json` missing | `/api/views` empty; UI hint to create it. |
| Missing creds/view on fetch | 400 `{"error"}`; red status pill. |
| Exporter 403 / bad creds / unknown view | error surfaced in status pill. |
| Supplier cat-file absent | Cat blank; dashboard still renders. |
| No `refined.csv` on first load | empty state shown; no error. |
| Zone with zero rows | KPIs show 0; charts empty; narrative "0 risks". |

## Testing

- `stats.py` / `narrative.py` already unit-tested — reused unchanged.
- Add `tests/test_dashboard.py`: build the all-zones payload from a small
  fixture CSV and assert shape (keys present, `Overall.total` == sum, zone
  totals, `narrative` non-empty). No server socket, no Excel, no network — call
  the payload-builder function directly.
- Manual smoke: start server, load cached data, switch zones, one live fetch.

## Out of scope

- No changes to the Excel report pipeline (`build_zone_reports.py` etc.).
- No auth beyond localhost binding; single-user local tool.
- No scheduled/auto refresh; Fetch is manual.
- No write-back to OneTrust (read-only, same as the exporter).
