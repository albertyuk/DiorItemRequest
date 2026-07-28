"""Extract highlighted SKUs from a Dior sell-thru map, look them up in a
stock query export, and fill a ProductsList template workbook.

All heavy workbooks are opened with ``read_only=True`` — the sell-thru map
expands to multiple GB of memory in normal mode.
"""

from __future__ import annotations

import logging
import re
import shutil
from dataclasses import dataclass, field

from openpyxl import load_workbook

log = logging.getLogger(__name__)

# --- Highlight recognition (kept configurable: a future recolor of the map
# should be a one-line change here) -----------------------------------------
# Yellow highlights in the sell-thru map are theme-indexed, not RGB:
# solid fill, fgColor.theme == 7 (accent4), tint ~= 0.80. fgColor.rgb is
# unset on those cells, so RGB-only detection finds nothing.
HIGHLIGHT_THEMES = {7}
# Fallback for workbooks that use literal RGB yellows instead of theme colors.
YELLOW_FAMILY = {"FFFF00", "FFF2CC", "FFE699", "FFC000", "FFD966"}
# Theme 4 is the structural light-blue banding on FW26 size rows — never a
# highlight and not worth flagging for review.
STRUCTURAL_THEMES = {4}

# SKU-shaped strings: 8+ alphanumerics, an X, then a 3+ alphanumeric suffix.
# Covers FW26 parent rows (X + 4-char color) and LOOK TOTAL variants.
SKU_RE = re.compile(r"^[0-9A-Z]{8,}X[0-9A-Z]{3,}$")
# The color block is always the *trailing* X + 4 alphanumerics. Never use
# split('X'): bases themselves may contain an X (652P92X3F74X8090).
COLOR_SUFFIX_RE = re.compile(r"X[0-9A-Z]{4}$")

QUERY_SHEET = "query"
QUERY_HEADERS = [
    "Barcode", "Sku", "Division", "Department", "Season",
    "Product reference", "Color code", "Size", "Quantity",
    "Unit Weighted Average Cost", "Unit Retail Price incl. tax",
    "Item Type", "Path",
]

TEMPLATE_SHEET = "Sheet1"
TEMPLATE_COLS = 13  # A..M

# Cap the "other fills" review list so one oddly formatted sheet cannot
# balloon the report.
MAX_OTHER_FILLS = 200
# Cap the per-sheet highlighted-cell trace shown on the report page (a fully
# highlighted sheet must not produce an unbounded page).
MAX_HIGHLIGHT_DETAILS = 1000


def is_highlighted(cell) -> bool:
    """True when the cell carries a yellow 'reorder' highlight."""
    f = cell.fill
    if f is None or f.patternType != "solid":
        return False
    fg = f.fgColor
    theme = getattr(fg, "theme", None)
    if isinstance(theme, int):
        return theme in HIGHLIGHT_THEMES
    rgb = fg.rgb if isinstance(fg.rgb, str) else ""
    return rgb[-6:].upper() in YELLOW_FAMILY


def _fill_note(cell) -> str | None:
    """Describe a solid fill that is neither a highlight nor structural
    banding, for the human review list. Returns None for anything else."""
    f = cell.fill
    if f is None or f.patternType != "solid":
        return None
    fg = f.fgColor
    theme = getattr(fg, "theme", None)
    if isinstance(theme, int):
        if theme in HIGHLIGHT_THEMES or theme in STRUCTURAL_THEMES:
            return None
        tint = fg.tint or 0
        return f"theme {theme}, tint {tint:.2f}"
    rgb = fg.rgb if isinstance(fg.rgb, str) else ""
    if rgb:
        if rgb[-6:].upper() in YELLOW_FAMILY:
            return None
        return f"rgb {rgb[-6:].upper()}"
    # Legacy color types (indexed palette, auto) — openpyxl exposes unset
    # attributes as descriptor objects, so nothing above matched. These are
    # never treated as highlights, but must still reach the human review
    # list rather than vanish silently.
    indexed = getattr(fg, "indexed", None)
    if isinstance(indexed, int):
        return f"indexed color {indexed}"
    return "unrecognized solid fill"


def extract_base(sku: str) -> str:
    """Strip the trailing color block (X + 4 alphanumerics) from a SKU."""
    return COLOR_SUFFIX_RE.sub("", sku.strip())


@dataclass
class SheetScan:
    name: str
    sku_cells: int = 0
    highlighted_cells: int = 0
    highlighted_skus: set[str] = field(default_factory=set)
    # step-by-step trace for the report page: (cell coordinate, sku, base)
    highlighted_cell_details: list[tuple[str, str, str]] = field(default_factory=list)
    other_fills: list[tuple[str, str, str]] = field(default_factory=list)  # (coord, sku, note)

    @property
    def bases(self) -> set[str]:
        return {extract_base(s) for s in self.highlighted_skus}


@dataclass
class MapScan:
    sheets: list[SheetScan]

    @property
    def bases(self) -> dict[str, list[str]]:
        """Union of highlighted bases -> sorted list of sheets they appear on."""
        out: dict[str, set[str]] = {}
        for sheet in self.sheets:
            for base in sheet.bases:
                out.setdefault(base, set()).add(sheet.name)
        return {b: sorted(names) for b, names in sorted(out.items())}

    @property
    def other_fills(self) -> list[tuple[str, str, str, str]]:
        return [
            (sheet.name, coord, sku, note)
            for sheet in self.sheets
            for coord, sku, note in sheet.other_fills
        ]


def scan_sell_thru_map(path) -> MapScan:
    """Scan every cell of every sheet for SKU-shaped strings and classify
    their fills. Geometry-free on purpose: it handles both the FW26 row
    layout and the LOOK TOTAL block layout without hardcoded columns."""
    wb = load_workbook(path, read_only=True)
    try:
        sheets = []
        for ws in wb.worksheets:
            scan = SheetScan(name=ws.title)
            for row in ws.iter_rows():
                for cell in row:
                    value = cell.value
                    if not isinstance(value, str):
                        continue
                    sku = value.strip()
                    if not SKU_RE.match(sku):
                        continue
                    scan.sku_cells += 1
                    if is_highlighted(cell):
                        scan.highlighted_cells += 1
                        scan.highlighted_skus.add(sku)
                        if len(scan.highlighted_cell_details) < MAX_HIGHLIGHT_DETAILS:
                            scan.highlighted_cell_details.append(
                                (cell.coordinate, sku, extract_base(sku)))
                    elif len(scan.other_fills) < MAX_OTHER_FILLS:
                        note = _fill_note(cell)
                        if note:
                            scan.other_fills.append((cell.coordinate, sku, note))
            sheets.append(scan)
        return MapScan(sheets=sheets)
    finally:
        wb.close()


def read_query_headers(path) -> list:
    wb = load_workbook(path, read_only=True)
    try:
        ws = wb[QUERY_SHEET] if QUERY_SHEET in wb.sheetnames else wb.worksheets[0]
        first = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())
        headers = list(first)
        while headers and headers[-1] is None:
            headers.pop()
        return headers
    finally:
        wb.close()


def validate_query_file(path) -> tuple[bool, str, int]:
    """Check the 13 expected headers; returns (ok, message, data_row_count)."""
    try:
        headers = read_query_headers(path)
    except Exception as exc:  # malformed zip/xlsx
        return False, f"could not be read as an .xlsx workbook ({exc})", 0
    normalized = [str(h).strip() if h is not None else "" for h in headers]
    if normalized != QUERY_HEADERS:
        return (
            False,
            "header row does not match the expected 13 query columns "
            f"(got: {normalized[:14]})",
            0,
        )
    return True, "ok", count_query_rows(path)


def count_query_rows(path) -> int:
    wb = load_workbook(path, read_only=True)
    try:
        ws = wb[QUERY_SHEET] if QUERY_SHEET in wb.sheetnames else wb.worksheets[0]
        return sum(
            1
            for row in ws.iter_rows(min_row=2, values_only=True)
            if any(v is not None for v in row)
        )
    finally:
        wb.close()


def match_query(path, bases) -> dict[str, list[tuple]]:
    """Single pass over the query export. A query Sku is 'BASE - COLOR - SIZE'
    with literal ' - ' separators; match on exact first-segment equality
    (never prefix-match raw strings — first segments can contain X too).

    Returns {base: [query row tuples]} containing only matched bases.
    """
    bases = set(bases)
    matched: dict[str, list[tuple]] = {}
    wb = load_workbook(path, read_only=True)
    try:
        ws = wb[QUERY_SHEET] if QUERY_SHEET in wb.sheetnames else wb.worksheets[0]
        for row in ws.iter_rows(min_row=2, values_only=True):
            sku = row[1] if len(row) > 1 else None
            if not isinstance(sku, str):
                continue
            base = sku.split(" - ")[0]
            if base in bases:
                if len(row) < 13:  # ragged row: pad so consumers can unpack
                    row = tuple(row) + (None,) * (13 - len(row))
                matched.setdefault(base, []).append(row)
    finally:
        wb.close()
    return matched


def _sku_sort_key(row) -> tuple[str, str, str]:
    parts = [p.strip() for p in str(row[1]).split(" - ")]
    parts += [""] * (3 - len(parts))
    return parts[0], parts[1], parts[2]


def _barcode_text(value) -> str:
    """Barcodes must stay text — no lost leading zeros, no scientific
    notation. The query stores them as strings already; be defensive about
    numeric cells anyway."""
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if value is None:
        return ""
    return str(value)


def build_output(template_path, out_path, matched, unmatched_bases,
                 base_sheets=None) -> int:
    """Copy the template and append one row per matched query row, sorted by
    base, then color, then size. Returns the number of data rows written.

    The query and template columns do NOT align — rows are mapped by meaning,
    and query's Product reference / Color code / Size / Item Type / Path are
    intentionally dropped.
    """
    shutil.copyfile(template_path, out_path)
    wb = load_workbook(out_path)
    ws = wb[TEMPLATE_SHEET] if TEMPLATE_SHEET in wb.sheetnames else wb.worksheets[0]

    # Row 2 of the template only demonstrates the H/J formula pattern; clear
    # it so it is either overwritten as the first data row or not left behind.
    for col in range(1, TEMPLATE_COLS + 1):
        ws.cell(row=2, column=col).value = None

    rows = [row for base in sorted(matched) for row in matched[base]]
    rows.sort(key=_sku_sort_key)

    n = 1
    for query_row in rows:
        n += 1
        (barcode, sku, division, department, season,
         _ref, _color, _size, quantity, unit_cost, unit_retail,
         _item_type, _path) = query_row[:13]
        barcode_cell = ws.cell(row=n, column=1, value=_barcode_text(barcode))
        barcode_cell.number_format = "@"
        ws.cell(row=n, column=2, value=sku)
        ws.cell(row=n, column=3, value=division)
        ws.cell(row=n, column=4, value=department)
        ws.cell(row=n, column=5, value=season)
        ws.cell(row=n, column=6, value=quantity)
        ws.cell(row=n, column=7, value=unit_cost)
        ws.cell(row=n, column=8,
                value=f'=IF(AND(F{n}<>"", G{n}<>""), F{n}*G{n}, "")')
        ws.cell(row=n, column=9, value=unit_retail)
        ws.cell(row=n, column=10,
                value=f'=IF(AND(F{n}<>"", I{n}<>""), F{n}*I{n}, "")')
        # K, L, M (Gift Recipient / Organization / Position) stay blank.

    unmatched_ws = wb.create_sheet("Unmatched")
    unmatched_ws.append(["Base SKU", "Highlighted on sheet(s)"])
    base_sheets = base_sheets or {}
    for base in sorted(unmatched_bases):
        unmatched_ws.append([base, ", ".join(base_sheets.get(base, []))])

    wb.save(out_path)
    wb.close()
    return n - 1


@dataclass
class RunReport:
    sheets: list[SheetScan]
    bases: dict[str, list[str]]          # base -> sheets it was highlighted on
    base_skus: dict[str, list[str]]      # base -> highlighted colorway SKUs
    matched_counts: dict[str, int]       # base -> number of query rows
    matched_rows: dict[str, list[dict]]  # base -> query data pulled (trace)
    unmatched: list[str]
    other_fills: list[tuple[str, str, str, str]]
    rows_written: int


def _trace_rows(rows) -> list[dict]:
    """Shape matched query rows for the report page's step-by-step trace."""
    out = []
    for r in sorted(rows, key=_sku_sort_key):
        out.append({
            "barcode": _barcode_text(r[0]), "sku": r[1], "division": r[2],
            "department": r[3], "season": r[4], "quantity": r[8],
            "unit_cost": r[9], "unit_retail": r[10],
        })
    return out


def run_pipeline(map_path, query_path, template_path, out_path) -> RunReport:
    scan = scan_sell_thru_map(map_path)
    bases = scan.bases
    matched = match_query(query_path, bases)
    unmatched = sorted(set(bases) - set(matched))
    rows_written = build_output(
        template_path, out_path, matched, unmatched, base_sheets=bases
    )
    base_skus: dict[str, set] = {}
    for sheet in scan.sheets:
        for sku in sheet.highlighted_skus:
            base_skus.setdefault(extract_base(sku), set()).add(sku)
    report = RunReport(
        sheets=scan.sheets,
        bases=bases,
        base_skus={b: sorted(s) for b, s in sorted(base_skus.items())},
        matched_counts={b: len(matched[b]) for b in sorted(matched)},
        matched_rows={b: _trace_rows(matched[b]) for b in sorted(matched)},
        unmatched=unmatched,
        other_fills=scan.other_fills,
        rows_written=rows_written,
    )
    log.info(
        "run: sheets=%s bases=%d matched=%d unmatched=%d rows=%d other_fills=%d",
        [(s.name, s.highlighted_cells) for s in report.sheets],
        len(report.bases), len(report.matched_counts),
        len(report.unmatched), report.rows_written, len(report.other_fills),
    )
    return report
