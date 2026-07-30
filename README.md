# Dior item request — highlighted-SKU extractor

Small internal web tool. Upload a **WRTW sell-thru map** workbook; the tool
extracts every **yellow-highlighted** ("reorder") SKU from all sheets, looks
up all color/size variants in a stored **stock query export**, and produces a
filled **ProductsList** workbook for download.

## How it works

- **Explicit selection beats color detection**: a sheet whose header area
  (first 15 rows) contains a column headed `PickUp` (any casing/spacing)
  is selected *only* by that column — rows marked `Y` (or yes / x / 1 /
  pickup / 是) are extracted, unrecognized values and highlighted-but-
  unmarked SKUs are surfaced on the report, and highlights there no
  longer select. Sheets without the column keep the highlight behavior
  below, so the seasonal map needs no changes.
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
- A bilingual user tutorial lives at `/help` (linked from the front page):
  what the tool does, a first-run walkthrough, how to read the report, and
  the common gotchas (pale theme-yellow highlights, filter-hidden rows).

## AI SKU-column detection (optional)

Workbooks other than the standard sell-thru map (delivery trackers,
transfer logs) carry SKUs in formats the pattern scanner cannot recognize —
accessory MMCs like `M0759OWKAM912` or underscore codes like
`641V19A1491_X8300` — under headers such as `TS SKU`, `SKU`, or
`Row Labels`. With an Anthropic API key configured, each run sends a small
preview of every sheet (first ~12 rows, values truncated — never the whole
workbook) to Claude, which returns the header row and SKU column(s) per
sheet. Highlighted cells in those columns are then extracted too, and the
run report shows what was detected and why (`sku_locator.py`).

Approved mappings are **remembered**: when a reviewer builds a run keeping
AI-found SKUs, the sheet's columns are stored under a fingerprint of its
header row (`data/column_memory.json`). The next upload of the same layout
gets those columns from memory — pre-approved, marked "remembered" on the
review page, no AI call for that sheet. Excluding every AI-found SKU of a
sheet at review drops its stored mapping. A layout change (shifted or
renamed headers) misses the fingerprint and falls back to fresh detection.
Remembered sheets keep working even when the API call for the other sheets
fails — the failure is reported on the review page and run report, but it
never costs the run its approved mappings.

Every run pauses on a **review checkpoint** before anything is written:
the page lists each extracted SKU (colorways, source sheets, stock-row
count, an AI badge for AI-located columns with sample values and the
model's reasoning), all individually selectable. Only ticked SKUs go into
the workbook; exclusions are recorded on the run report. Matching happens
at scan time, so the reviewed data cannot change while the page sits open.

```bash
fly secrets set ANTHROPIC_API_KEY=sk-ant-...   # or export it locally
```

- Without the key the feature is off and the app behaves exactly as before;
  the report says so.
- Model defaults to `claude-opus-5` (override with `ANTHROPIC_MODEL`).
  Server-side refusal fallback is enabled, so a safety-classifier decline
  re-runs on Anthropic's recommended substitute model automatically.
- One API call per run on a ~1 KB preview — cost is a fraction of a cent;
  detection failures (rate limit, network) never fail the run.

## Accounts & login

The app is a closed team tool — no self-signup. Three ideas carry the design:

- **`APP_PASSWORD` is a *setup code*, not a login password.** Its only power
  is creating (or recovering) the **admin** account at `/setup`. Nobody logs
  in with it day-to-day; treat it as the root secret.
- **Accounts are created by admins on `/team`** — with an email address an
  invite link is sent and the coworker picks their own password (the admin
  never learns it); without one the admin sets an initial password and
  shares it privately. Every run records who uploaded and who built.
- **Email addresses must prove ownership** (a confirmation link) before they
  can receive password-reset links — a typo'd address is a stranger's inbox.

Sessions are signed cookies versioned against the password hash: changing or
resetting a password signs that account out everywhere, instantly. Login,
setup, and forgot-password are all rate-limited. Everything degrades
gracefully: without an email provider, invites/reset disappear and admins set
passwords by hand.

| Env var | Default | Meaning |
|---|---|---|
| `APP_PASSWORD` | *(required in prod)* | Setup code: creates/recovers the admin at `/setup`. Unset → the app fails closed (503) unless `ALLOW_OPEN_ACCESS` is set. |
| `ALLOW_OPEN_ACCESS` | `0` | Explicit no-auth opt-out for local dev only. |
| `APP_SECRET` | *(generated)* | Session signing key; auto-generated and persisted to `DATA_DIR/session_secret` when unset. |
| `SESSION_COOKIE_SECURE` | `1` | Set `0` only for plain-HTTP local dev. |
| `RESEND_API_KEY` | *(unset)* | Enables invite / confirmation / reset emails (Resend). |
| `EMAIL_FROM` | resend onboarding | From address (domain must be verified with Resend). |
| `PUBLIC_BASE_URL` | *(unset)* | Absolute origin for emailed links, e.g. `https://dior-item-request.fly.dev`. Required for correct links in production. |
| `INVITE_TTL_HOURS` / `RESET_TTL_HOURS` | `72` / `2` | Link lifetimes. |

**Migrating from `APP_USERS`:** that variable is retired and ignored (a
warning is logged). Open `/setup`, enter the `APP_PASSWORD` value as the
setup code to create your admin account, then add coworkers on `/team`.

## Local run

```bash
python3 -m venv .venv && . .venv/bin/activate
pip install -r requirements.txt
ALLOW_OPEN_ACCESS=1 python -m flask --app app run --port 8080
# open http://localhost:8080 — or set APP_PASSWORD instead and log in via /setup
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
fly secrets set APP_PASSWORD=<choose-a-strong-setup-code>
fly secrets set PUBLIC_BASE_URL=https://<your-app>.fly.dev
fly secrets set RESEND_API_KEY=re_...        # optional: invite/reset emails
fly deploy --ha=false
```

Then open the site, follow the redirect to `/setup`, enter the setup code,
and create your admin account; add coworkers on `/team`.

Notes:

- **Authentication is required in production** — without a setup code (and
  no explicit `ALLOW_OPEN_ACCESS=1`) the app refuses to start, and the
  request gate independently fails closed with 503 (it fronts internal
  pricing data). See “Accounts & login” above for how accounts, invites,
  and password reset work.
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
