"""Locate SKU columns in arbitrary workbook layouts with the Anthropic API.

The regex scanner in pipeline.py finds standard Dior RTW SKUs anywhere on a
sheet, but other workbooks (delivery trackers, transfer logs) carry SKUs in
formats it cannot recognize — accessory MMCs like ``M0759OWKAM912`` or
underscore codes like ``641V19A1491_X8300`` — under headers such as
``TS SKU``, ``SKU``, ``MMC``, or ``Row Labels``. This module sends a small
preview of each sheet (the first rows, values truncated) to Claude, which
returns the header row and SKU column letters per sheet as structured JSON.
The pipeline then also treats highlighted cells in those columns as SKUs.

The feature is optional: without an ANTHROPIC_API_KEY the app runs exactly
as before. Only cell previews are sent — never the whole workbook.
"""

from __future__ import annotations

import json
import logging
import os
import re

from openpyxl import load_workbook

log = logging.getLogger(__name__)

# claude-opus-5 is the default per current Anthropic guidance; override with
# the ANTHROPIC_MODEL env var if your org standardizes on something else.
DEFAULT_MODEL = "claude-opus-5"

PREVIEW_ROWS = 12
PREVIEW_COLS = 40
CELL_CHARS = 40
REQUEST_TIMEOUT = 60.0

COLUMN_RE = re.compile(r"^[A-Z]{1,3}$")

SYSTEM_PROMPT = """\
You analyze Excel sheet previews from Dior merchandising workbooks
(sell-thru maps, delivery trackers, stock transfer logs). The user message
is JSON: one entry per sheet with the first rows of that sheet, keyed by
column letter; empty cells are omitted and long values are truncated.

For every sheet, identify:
- header_row: the 1-based row number of the header row (null if the sheet
  has no header row).
- sku_columns: the column(s) whose data cells contain product SKU / MMC
  codes — style-color item identifiers such as 644S16B7E72X5805,
  M0759OWKAM912, KCK554TFS_S03W, or 641V19A1491_X8300. Judge by the data
  cells, not only the header: SKU columns appear under headers like SKU,
  MMC, TS SKU, STOCK SKU, or Row Labels, and sometimes with no header at
  all. Prefer columns holding literal values; exclude columns whose cells
  are formulas deriving from another SKU column unless no literal SKU
  column exists on that sheet. Exclude fragment columns (style-only,
  color-only, size-only) and quantity/date/name columns. If a sheet has no
  SKU column, return an empty sku_columns list for it.
"""

OUTPUT_SCHEMA = {
    "type": "object",
    "properties": {
        "sheets": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "sheet": {"type": "string"},
                    "header_row": {"anyOf": [{"type": "integer"}, {"type": "null"}]},
                    "sku_columns": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "column": {
                                    "type": "string",
                                    "description": "Column letter, e.g. 'E'",
                                },
                                "header": {"anyOf": [{"type": "string"}, {"type": "null"}]},
                                "reason": {"type": "string"},
                            },
                            "required": ["column", "header", "reason"],
                            "additionalProperties": False,
                        },
                    },
                },
                "required": ["sheet", "header_row", "sku_columns"],
                "additionalProperties": False,
            },
        }
    },
    "required": ["sheets"],
    "additionalProperties": False,
}


class SkuLocatorError(Exception):
    """Raised when SKU-column detection fails; the pipeline continues
    without it and surfaces the message on the run report."""


def is_configured() -> bool:
    return bool(os.environ.get("ANTHROPIC_API_KEY")
                or os.environ.get("ANTHROPIC_AUTH_TOKEN"))


def build_previews(map_path) -> list[dict]:
    """Compact preview of each sheet: first rows, non-empty cells only."""
    wb = load_workbook(map_path, read_only=True)
    try:
        previews = []
        for ws in wb.worksheets:
            rows = {}
            for row in ws.iter_rows(min_row=1, max_row=PREVIEW_ROWS,
                                    max_col=PREVIEW_COLS):
                cells = {}
                row_num = None
                for cell in row:
                    value = cell.value
                    if value is None:
                        continue
                    # Read the row number from a cell that holds a value:
                    # in read-only mode, empty cells are a shared EmptyCell
                    # singleton with no .row/.column_letter at all, so
                    # row[0] must never be used for coordinates.
                    row_num = cell.row
                    cells[cell.column_letter] = str(value)[:CELL_CHARS]
                if cells:
                    rows[str(row_num)] = cells
            if rows:
                previews.append({"sheet": ws.title, "rows": rows})
        return previews
    finally:
        wb.close()


def _normalize(raw_sheets, known_sheets) -> list[dict]:
    """Validate the model's answer defensively before the scanner uses it."""
    known = set(known_sheets)
    detection = []
    for entry in raw_sheets:
        name = entry.get("sheet")
        if name not in known:
            continue
        header_row = entry.get("header_row")
        if not isinstance(header_row, int) or header_row < 1:
            header_row = None
        columns = []
        for col in entry.get("sku_columns") or []:
            letter = str(col.get("column", "")).strip().upper()
            if not COLUMN_RE.match(letter):
                continue
            columns.append({
                "column": letter,
                "header": col.get("header"),
                "reason": col.get("reason", ""),
            })
        # Sheets without SKU columns are dropped entirely: the model is
        # instructed to return them with an empty sku_columns list, and
        # keeping such entries would make "no columns found" look truthy
        # to callers (triggering a pointless verification page).
        if columns:
            detection.append({
                "sheet": name,
                "header_row": header_row,
                "sku_columns": columns,
                "source": "ai",
            })
    return detection


def column_samples(previews, detection, limit: int = 6) -> dict:
    """Real cell values from each detected column, for the human
    verification page: {sheet: {column: [values]}}. Values come from the
    preview rows below the header row."""
    by_sheet = {p["sheet"]: p["rows"] for p in previews}
    samples: dict = {}
    for entry in detection:
        rows = by_sheet.get(entry["sheet"]) or {}
        header_row = entry.get("header_row") or 0
        for col in entry["sku_columns"]:
            letter = col["column"]
            values = []
            for row_num in sorted(rows, key=int):
                if int(row_num) <= header_row:
                    continue
                value = rows[row_num].get(letter)
                if value and not value.startswith("=") and value not in values:
                    values.append(value)
                if len(values) >= limit:
                    break
            samples.setdefault(entry["sheet"], {})[letter] = values
    return samples


def locate(map_path, model: str | None = None, client=None) -> list[dict]:
    """Identify SKU columns for every sheet of the workbook at map_path.

    Returns [{sheet, header_row, sku_columns: [{column, header, reason}]}].
    Raises SkuLocatorError on any failure — callers treat it as "detection
    unavailable", never as a fatal pipeline error.
    """
    return locate_from_previews(build_previews(map_path), model=model,
                                client=client)


def locate_from_previews(previews, model: str | None = None,
                         client=None) -> list[dict]:
    """API half of locate(); takes previews from build_previews so callers
    can reuse them (e.g. for the verification page's sample values)."""
    import anthropic

    if not previews:
        return []

    if client is None:
        client = anthropic.Anthropic(timeout=REQUEST_TIMEOUT)
    model = model or os.environ.get("ANTHROPIC_MODEL", DEFAULT_MODEL)

    try:
        response = client.beta.messages.create(
            model=model,
            max_tokens=4096,
            system=SYSTEM_PROMPT,
            messages=[{
                "role": "user",
                "content": json.dumps(previews, ensure_ascii=False),
            }],
            output_config={
                # Header identification is a simple extraction task; low
                # effort keeps the call fast and cheap.
                "effort": "low",
                "format": {"type": "json_schema", "schema": OUTPUT_SCHEMA},
            },
            # Safety classifiers on claude-opus-5 can decline a request;
            # the server-side default fallback re-runs it on the
            # recommended substitute model instead of failing the call.
            betas=["server-side-fallback-2026-07-01"],
            extra_body={"fallbacks": "default"},
        )
    except anthropic.AuthenticationError as exc:
        raise SkuLocatorError(f"Anthropic API key was rejected ({exc.message})")
    except anthropic.RateLimitError:
        raise SkuLocatorError("Anthropic API rate limit reached — try again "
                              "in a minute")
    except anthropic.APIStatusError as exc:
        raise SkuLocatorError(f"Anthropic API error {exc.status_code}: "
                              f"{exc.message}")
    except anthropic.APIConnectionError:
        raise SkuLocatorError("could not reach the Anthropic API (network)")

    if response.stop_reason == "refusal":
        raise SkuLocatorError("the model declined to analyze this workbook")

    text = next((b.text for b in response.content if b.type == "text"), "")
    try:
        raw = json.loads(text)["sheets"]
    except (ValueError, KeyError) as exc:
        raise SkuLocatorError(f"unexpected model response: {exc}")

    detection = _normalize(raw, [p["sheet"] for p in previews])
    log.info("SKU column detection (%s): %s", model,
             [(d["sheet"], [c["column"] for c in d["sku_columns"]])
              for d in detection])
    return detection
