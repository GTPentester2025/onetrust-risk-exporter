# OneTrust Risk Register — View Exporter

Export a saved risk view's rows to CSV using an OAuth2 client-credentials token.

## Important: what the token can and cannot do

OneTrust's **internal** risk-views and report-export endpoints
(`/api/risk-v2/v2/risk-views`, `/reports/export`) are gated to logged-in **user
sessions**. An OAuth2 **client-credentials** token gets `403` on them. So this
tool does not replicate the UI "Export" button.

Instead it uses the **OAuth-accessible grid endpoint**, which returns risk rows
as JSON:

```
POST /api/risk/v2/risks/views/pages     (filters -> rows, paginated)
```

The script posts a view's filters, pages through all rows, and writes a
flattened CSV. Because the token cannot read view definitions, each view's
filters must be captured once from the UI (see below).

## Setup

```bash
pip install requests
cp views.example.json views.json     # then edit views.json with real filters
```

`views.json` is gitignored — it holds production org/person UUIDs and must stay
local.

## Capturing a view's filters

1. Open the view in OneTrust. `F12` -> Network.
2. Trigger the grid/export; find the request whose payload has
   `filters: [ { field, operator, value } ]`.
3. Copy the `filters`, `visibleColumns`, and `sortInfo` into a `views.json`
   entry. Flatten filter values to plain UUID strings.

## Usage

```bash
python onetrust_view_export.py                 # pick a view, export to CSV
python onetrust_view_export.py --list
python onetrust_view_export.py --view "My View" --out my.csv
python onetrust_view_export.py --dump          # print first raw response, no CSV
python onetrust_view_export.py --endpoint /api/risk/v2/risks/pages   # alt endpoint
```

Credentials: client id + secret when prompted, or env vars
`ONETRUST_HOSTNAME`, `ONETRUST_CLIENT_ID`, `ONETRUST_CLIENT_SECRET`
(or `ONETRUST_TOKEN`). Hostname example: `app-de.onetrust.com`.

## Read/write

This tool only reads — it POSTs to a search/list endpoint and writes a local
CSV. It does not create, edit, or delete anything in OneTrust, and does not
trigger a server-side export job.

## Field names

Filter field names follow the UI export payload (e.g. `approvers`,
`organization`, `inherentLevel`). If the grid endpoint rejects a filter, run
`--dump` to inspect the raw response and adjust field names.

## Auth token

Create an OAuth2 client id + secret in OneTrust Global Settings. See the
[Quick Start Guide: APIs](https://developer.onetrust.com/onetrust/reference/quick-start-guide).
The client's role must have Risk read access, or `/risks/views/pages` returns 403.
