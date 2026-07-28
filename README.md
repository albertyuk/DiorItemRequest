# Dior item request — highlighted-SKU extractor

Small internal web tool. Upload a **WRTW sell-thru map** workbook; the tool
extracts every **yellow-highlighted** ("reorder") SKU from all sheets, looks
up all color/size variants in a stored **stock query export**, and produces a
filled **ProductsList** workbook for download.

## How it works

- Highlights in the map are theme-indexed fills (accent4 / theme 7, light
  tint), not RGB — detection handles both, configurable at the top of
  `pipeline.py` (`HIGHLIGHT_THEMES`, `YELLOW_FAMILY`).
- The structural light-blue banding on size rows (theme 4) is never treated
  as a highlight. Any other solid fill on a SKU cell (gray notes, etc.) is
  reported under "other fills detected" for a human to check.
- Base SKU = SKU minus the trailing color block (`X` + 4 alphanumerics).
  `652P92X3F74X8090` → `652P92X3F74` (bases can themselves contain an `X`).
- Query match: exact equality with the first ` - `-separated segment of the
  query `Sku` column. One output row per matching query row, sorted by base,
  color, size.
- Output = copy of `assets/ProductsListTemplate.xlsx` with mapped columns,
  `TotalCost`/`TotalSales` as formulas, barcodes as text, plus an
  `Unmatched` sheet listing highlighted bases with zero query rows.
- The UI is bilingual (English / 简体中文) — top-left toggle, remembered in a
  cookie. Strings live in `translations.py`. Run reports are persisted as
  JSON next to their output (`/report/<run_id>`), so they survive refresh
  and can be re-read later in either language; they are pruned together
  with their workbook (last 10 kept).

## Local run

```bash
python3 -m venv .venv && . .venv/bin/activate
pip install -r requirements.txt
python -m flask --app app run --port 8080
# open http://localhost:8080 — no password required locally unless APP_PASSWORD is set
```

Data (stored query export, last 10 outputs) lives in `DATA_DIR`
(default `/data` when present, else `./data`).

## Tests

Acceptance tests run against the three real sample workbooks, expected in
`./samples/` (not committed — they contain internal pricing data):

```
samples/sell_thru_map.xlsx        # WRTW FW26 sell-thru map
samples/query.xlsx                # stock query export
samples/ProductsListTemplate.xlsx # output template
```

```bash
pip install pytest
pytest test_pipeline.py -v
```

Sample-dependent tests skip automatically when `samples/` is absent;
layout-logic tests (including the LOOK TOTAL block layout, which this
sample copy of the map does not contain) run against generated fixtures.

## Deploy to Fly.io

```bash
fly launch --no-deploy          # accepts existing fly.toml; pick/adjust app name
fly volumes create data --size 1
fly secrets set APP_PASSWORD=<choose-a-strong-password>
fly deploy --ha=false
```

Notes:

- **`APP_PASSWORD` is required in production** — the app refuses to start on
  Fly without it (it fronts internal pricing data). Every route is behind
  HTTP Basic auth: any username, this password.
- The machine needs **1 GB memory** (set in `fly.toml`) — parsing the ~70 MB
  map read-only peaks well above the 256 MB default.
- **Deploy with `--ha=false` (single machine).** Fly's default first deploy
  creates two machines, each with its *own* independent `data` volume, so
  the stored query and generated outputs would randomly split between them
  (uploads landing on one machine, downloads 404ing on the other). If the
  app already has two machines, fix it with `fly scale count 1`.
- The `data` volume persists the stored query export and the last 10
  generated outputs across machine stops.
- Uploads are capped at 200 MB; a run takes tens of seconds (synchronous,
  gunicorn timeout 300 s).
