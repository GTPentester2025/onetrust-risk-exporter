# Risk Insights Dashboard v2 — Power BI Shell, Editable PPT, Config & Scheduling

**Date:** 2026-08-05
**Status:** Approved

## Purpose

Turn the existing local dashboard into a clean, user-facing analytics console:

1. **Power BI-style dark UI** — a professional report canvas of tiles, not a form.
2. **Editable PowerPoint export** — a full deck built with native PowerPoint
   charts/tables (not images), data identical to the on-screen numbers.
3. **Config popup + scheduling** — credentials and a fetch schedule live behind
   a gear modal so end-users never touch API details; fetches run
   asynchronously and on a cadence.

Existing behavior that must be preserved: zone insights (KPI, domain pie,
domain×Cat stacked bar, narrative, crosstab table), instant client-side zone
switching from a cached all-zones payload, reuse of `zone_reports/stats.py` +
`narrative.py`, `refined.csv` as the cache, hard-coded view "TPRM Global View".

## Architecture overview

```
config.json (gitignored)  ──►  scheduler thread ─┐
                                Refresh button  ──┤► async fetch job
                                                  │   onetrust_view_export.py
                                                  │   refine_columns.py
                                                  ▼
                                              refined.csv ──► stats/build_payload
                                                                 │        │
                                                        GET /api/data   GET /api/ppt
                                                                 │        │
                                                            index.html  <zone>.pptx
```

New pip dependency: **python-pptx** (for the PPT builder), auto-installed on
first use with the same pattern `build_zone_reports.py` uses for pywin32. The
HTTP server itself remains stdlib-only.

## Component 1 — Configuration store

New file `config_store.py` (pure, unit-testable; no HTTP, no threads).

Config shape (persisted to `config.json`, gitignored):

```json
{
  "hostname": "app-de.onetrust.com",
  "client_id": "…",
  "client_secret": "…",
  "schedule": { "mode": "off|hourly|daily|interval",
                "time": "06:00", "interval_hours": 6 }
}
```

Functions:
- `load_config(path) -> dict` — read JSON; missing file → default config
  (`{"hostname":"","client_id":"","client_secret":"","schedule":{"mode":"off","time":"06:00","interval_hours":6}}`).
  Corrupt file → default config (do not crash).
- `save_config(path, incoming, existing) -> dict` — merge `incoming` over
  `existing`; **if `incoming["client_secret"]` is empty/absent, keep
  `existing["client_secret"]`** (so the UI never has to resend the secret).
  Write pretty JSON. Return the merged config.
- `mask_config(cfg) -> dict` — return a copy safe for the browser:
  `client_secret` replaced by `has_secret: bool`; everything else intact.
  The raw secret is NEVER included.
- `is_configured(cfg) -> bool` — hostname + client_id + client_secret all
  non-empty.

## Component 2 — Schedule computation

In `config_store.py` (pure):

- `next_run_due(schedule, last_run, now) -> bool` — given the schedule dict,
  the last successful run (datetime or None), and `now` (datetime), return
  whether a run is due:
  - `off` → always False.
  - `hourly` → True if `last_run` is None or `now - last_run >= 1h`.
  - `interval` → same with `interval_hours`.
  - `daily` → True if today's `time` has passed and `last_run` is None or was
    before today's scheduled instant.
  `now` and `last_run` are passed in (never `datetime.now()` inside) so this is
  deterministic and testable.

## Component 3 — Async fetch job + scheduler (dashboard.py)

Shared job state (module-level, guarded by a `threading.Lock`):

```python
JOB = {"state": "idle", "last_run": None, "rows": None,
       "message": "", "source": None}
```

- `run_fetch_job()` — the worker: reads config; if not configured, sets
  `state="error", message="Not configured"`. Else sets `state="running"`, runs
  `export_cmd` + `refine_cmd` (reusing the existing functions, creds from
  config), on success sets `state="done", last_run=<now>, rows=<n>,
  source="api"`, on failure `state="error", message=<tail>`. Only one job at a
  time (guarded by a `running` flag; a second trigger while running is a no-op
  returning "already running").
- `start_fetch_async()` — spawn `run_fetch_job` in a daemon thread if not
  already running; return whether it started.
- **Scheduler thread** — a daemon loop: every 30s, load config, call
  `next_run_due(schedule, JOB["last_run"], now)`; if due and no job running,
  call `start_fetch_async()`. Started in `main()`; not started under tests.

Routes:
- `POST /api/fetch` → `start_fetch_async()`; 202 `{"started": bool,
  "running": bool}`. No credentials in the body — the server uses config.
- `GET /api/status` → `{state, last_run(iso|null), rows, message, source}`.
- `GET /api/config` → `mask_config(load_config())`.
- `POST /api/config` → `save_config(path, body, load_config())`; reschedule is
  automatic (scheduler re-reads config each tick); return `mask_config(merged)`.
- `POST /api/config/test` → attempt an OAuth token only (hostname+id+secret
  from body, or saved if body omits them), return `{ok: bool, message}`. Never
  echo the secret. (Reuses the exporter's token step via a small helper or a
  `--test`-style subprocess; a token-only check is acceptable.)
- `GET /api/data` → unchanged (cached payload or `{"empty": true}`).
- `GET /api/ppt` → Component 4.

`run_fetch` (the old synchronous, creds-in-body function) is removed from the
request path; its export/refine command builders are retained and reused by the
job. Tests that referenced `run_fetch` are updated to the job/command builders.

## Component 4 — Editable PowerPoint export

New file `ppt_export.py`:

- Ensures `python-pptx` (auto-install if missing, same pattern as pywin32).
- `build_deck(payload) -> Presentation` — builds an editable deck:
  - **Title slide:** "OneTrust Risk Insights", the view name, and
    `generated_at` (human-readable).
  - **One slide per zone** (all 10, in the fixed order), each containing:
    - Zone title + total risks.
    - **KPI text** (total, top-2 domains %, top supplier category).
    - **Native pie chart** (`XL_CHART_TYPE.PIE`) of risks by domain
      (`by_domain`), with its `CategoryChartData` so it is editable in
      PowerPoint.
    - **Native stacked bar** (`XL_CHART_TYPE.COLUMN_STACKED`) of domain × Cat
      from `crosstab`, one series per Cat in the union column order
      (`cat_order` + extras, "(blank)" shown as "Uncategorized").
    - **Native table** (`add_table`) of the crosstab (Domain + cats + Total).
    - **Narrative text box** with the `narrative` lines.
  - A zero-row zone renders the slide with "0 risks" and no chart data.
- `deck_to_bytes(payload) -> bytes` — build and return `.pptx` bytes.

`GET /api/ppt` loads `refined.csv` → `build_payload` → `deck_to_bytes`, responds
with `application/vnd.openxmlformats-officedocument.presentationml.presentation`
and `Content-Disposition: attachment; filename="risk-insights.pptx"`. If no
data, 400 `{"error": "No data — refresh first"}`.

Data accuracy: the deck is built from the same `build_payload` output the UI
renders, so on-screen and in-deck numbers are identical by construction.

## Component 5 — UI redesign (index.html), Power BI dark

Single dark theme (no light branch). Tokens:
`--canvas #1b1a19; --tile #252423; --tile-2 #2f2e2d; --fg #f3f2f1;
--fg-muted #c8c6c4; --fg-dim #979593; --border #3b3a39; --gold #e8c810;
--gold-deep #c99a1e; --live #34d399; --danger #f0736a; --radius 8px;`
Fira Sans / Fira Code. Contrast ≥4.5:1 for body/small text; ≥3:1 for large.

Layout:
- **App bar:** product mark (gold) left; center/left a **zone tab strip**
  (Power BI page tabs) — the 10 zones, active tab gold-underlined; right cluster:
  **Refresh** button, a **status chip** (idle / "Refreshing…" spinner /
  "Updated <time>" green dot / "Error"), **Export ▾** menu, **⚙ Settings**.
- **Report canvas:** grid of tiles on the charcoal canvas:
  - KPI tiles row (total, top-2 domains %, top supplier category).
  - Row: domain **pie** tile | domain×Cat **stacked bar** tile.
  - Row: **crosstab table** tile | **narrative** tile.
  - Each tile: title + hover-revealed toolbar (Copy, Download) as before; keep
    the per-tile copy/download and the per-zone screenshot (moved into Export ▾).
- **Export ▾ menu:** PowerPoint (.pptx) → `GET /api/ppt`; Zone image (PNG,
  html2canvas, existing); Crosstab CSV (current zone, existing).
- **Settings modal (⚙):** Connection (hostname, client id, client secret with
  placeholder `•••• saved` when `has_secret`), **Test connection** button
  showing the `/api/config/test` result; **Schedule** (Off / Hourly / Daily at
  [time] / Every [N] hours); **Save** (POST /api/config). Modal has a scrim,
  Escape/close, focus trap, and does not leak the secret (never prefilled).
- **Onboarding banner:** if `GET /api/config` reports not configured, show a
  dismissible banner "Connect your OneTrust data — open Settings" with a button
  that opens the modal.
- **Refresh flow:** click → `POST /api/fetch` → poll `GET /api/status` every
  ~2s → on `done` reload `/api/data` and re-render; on `error` show the message
  in the status chip. The same poll surfaces scheduled runs (last_run updates
  even if the user didn't click).
- Zone switching stays instant/client-side from the cached payload; charts
  destroy-before-recreate; catLabel "(blank)"→"Uncategorized"; crosstab columns
  are the union of `cat_order` + extras.

## Error handling

| Condition | Result |
|--|--|
| No config / missing creds on Refresh | status chip "Not configured"; banner prompts Settings. |
| Test connection fails (bad creds/host) | modal shows red "Connection failed: <msg>". |
| Fetch job fails (403, network) | status chip "Error: <msg>"; previous cached data stays. |
| `/api/ppt` with no data | 400 `{"error":"No data — refresh first"}`; UI toast. |
| python-pptx install fails | `/api/ppt` 500 with a clear message; UI toast. |
| Corrupt `config.json` | treated as default config; app still runs. |

## Security

- Client secret stored only in gitignored `config.json` on local disk; never
  returned by `GET /api/config` (masked to `has_secret`), never logged, never
  placed in the DOM.
- Server binds `127.0.0.1` only. Request bodies are not logged.
- `.gitignore` adds `config.json` and `*.pptx`.

## Testing

- `config_store.py`: `load_config` default/corrupt; `save_config` preserves
  secret when blank and overwrites when provided; `mask_config` removes secret
  and sets `has_secret`; `is_configured` truth table.
- `next_run_due`: off/hourly/interval/daily cases with fixed `now`/`last_run`.
- `ppt_export.build_deck`: build from a small payload, reopen with
  `Presentation(BytesIO)`, assert slide count == zones+1, and that a slide
  contains a chart shape and a table shape. (Skipped with a clear message if
  python-pptx is unavailable in the test env.)
- Route handlers `handle_get`/`handle_post` for `/api/config` (masked, no
  secret), `/api/status`, `/api/fetch` (returns started/running), using
  monkeypatched job/config functions — no threads, no sockets, no network.
- Existing 20 tests: payload builder, exporter/refiner command builders, routes
  — keep green (update only the removed-`run_fetch` references).

## Out of scope

- Multi-user auth (single-user localhost tool).
- Editing data in the browser; the dashboard is read/report only.
- Live OneTrust write-back.
- Per-user theme switching (single dark theme).
