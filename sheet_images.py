"""Extract embedded product images from an .xlsx by reading the zip's
drawing XML directly.

openpyxl's read_only mode (the only way to open the ~70 MB sell-thru map
without exploding memory) never exposes images, so this module goes one
level down: workbook.xml maps sheet names to sheet XML parts, each sheet's
rels point at a drawing part, the drawing anchors carry the (row, col) of
every picture, and the drawing's rels resolve each picture to a file under
xl/media/. Only the media files that are actually wanted are read.

Extraction is strictly best-effort: any failure returns what was found so
far (or nothing) and logs — a malformed drawing must never fail a run.
"""

from __future__ import annotations

import base64
import logging
import posixpath
import re
import zipfile
from xml.etree import ElementTree as ET

log = logging.getLogger(__name__)

_NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    "xdr": ("http://schemas.openxmlformats.org/drawingml/2006/"
            "spreadsheetDrawing"),
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}
_R_ID = f"{{{_NS['r']}}}id"
_R_EMBED = f"{{{_NS['r']}}}embed"

# Raster formats the output workbook can re-embed. Vector formats (emf,
# wmf) are silently skipped.
_RASTER_EXT = {"png", "jpeg", "jpg", "gif", "bmp"}

# --- Bounds. The uploaded workbook is untrusted input: every zip entry is
# checked against its DECLARED uncompressed size before being read, so a
# decompression bomb (a few KB inflating to gigabytes) is refused instead
# of allocated. Budgets are cumulative, so "many small bombs" is bounded
# too.
MAX_IMAGE_BYTES = 4 * 1024 * 1024        # one picture
MAX_TOTAL_IMAGE_BYTES = 48 * 1024 * 1024  # all pictures of one run
MAX_XML_BYTES = 64 * 1024 * 1024          # one sheet/drawing/rels part
MAX_ANCHORS_PER_SHEET = 5_000             # crafted drawing with 10^6 anchors


def _read_bounded(zf: zipfile.ZipFile, name: str, limit: int) -> bytes:
    """Read one zip member only if its uncompressed size fits `limit`.

    zipfile.read() would happily inflate a 1 GB entry into memory; the
    header's declared size lets us refuse first. The declared size is
    attacker-controlled, so ZipFile.open is still capped by reading
    limit+1 bytes and rejecting an under-declared entry."""
    try:
        info = zf.getinfo(name)
    except KeyError:
        raise
    if info.file_size > limit:
        log.warning("skipping oversized zip member %r (%d bytes declared)",
                    name[:80], info.file_size)
        raise ValueError("zip member exceeds size limit")
    with zf.open(name) as fh:
        data = fh.read(limit + 1)
    if len(data) > limit:
        log.warning("zip member %r under-declared its size", name[:80])
        raise ValueError("zip member exceeds size limit")
    return data


def _rels(zf: zipfile.ZipFile, part: str) -> dict[str, str]:
    """{rId: absolute part path} for one part's relationships."""
    rels_path = posixpath.join(posixpath.dirname(part), "_rels",
                               posixpath.basename(part) + ".rels")
    try:
        root = ET.fromstring(_read_bounded(zf, rels_path, MAX_XML_BYTES))
    except (KeyError, ET.ParseError, ValueError):
        return {}
    out = {}
    for rel in root.findall("pr:Relationship", _NS):
        target = rel.get("Target", "")
        if rel.get("TargetMode") == "External":
            continue
        if target.startswith("/"):
            resolved = target.lstrip("/")
        else:
            resolved = posixpath.normpath(
                posixpath.join(posixpath.dirname(part), target))
        # A crafted Target ("../../..") normalizes to a path outside the
        # package. Nothing is ever opened from the filesystem here (only
        # zip members), but an escaping name is malformed by definition —
        # drop it rather than look it up.
        if resolved.startswith("../") or resolved.startswith("/"):
            continue
        out[rel.get("Id", "")] = resolved
    return out


def _sheet_parts(zf: zipfile.ZipFile) -> dict[str, str]:
    """{sheet name: worksheet part path}."""
    try:
        wb = ET.fromstring(_read_bounded(zf, "xl/workbook.xml",
                                         MAX_XML_BYTES))
    except (KeyError, ET.ParseError, ValueError):
        return {}
    rels = _rels(zf, "xl/workbook.xml")
    out = {}
    for sheet in wb.findall(".//main:sheets/main:sheet", _NS):
        part = rels.get(sheet.get(_R_ID, ""))
        name = sheet.get("name")
        if part and name:
            out[name] = part
    return out


def _anchor_rows(anchor) -> tuple[int, int, int]:
    """(from_row, to_row, from_col), all 1-based; to_row == from_row for
    one-cell anchors."""
    frm = anchor.find("xdr:from", _NS)
    if frm is None:
        raise ValueError("anchor without xdr:from")
    row1 = int(frm.findtext("xdr:row", default="0", namespaces=_NS)) + 1
    col1 = int(frm.findtext("xdr:col", default="0", namespaces=_NS)) + 1
    to = anchor.find("xdr:to", _NS)
    row2 = row1
    if to is not None:
        row2 = int(to.findtext("xdr:row", default=str(row1 - 1),
                               namespaces=_NS)) + 1
    return row1, max(row1, row2), col1


def sheet_anchors(zf: zipfile.ZipFile, sheet_part: str) \
        -> list[tuple[int, int, int, str]]:
    """[(from_row, to_row, from_col, media part path), ...] for one sheet."""
    try:
        ws = ET.fromstring(_read_bounded(zf, sheet_part, MAX_XML_BYTES))
    except (KeyError, ET.ParseError, ValueError):
        return []
    sheet_rels = _rels(zf, sheet_part)
    anchors = []
    for drawing in ws.findall("main:drawing", _NS):
        drawing_part = sheet_rels.get(drawing.get(_R_ID, ""))
        if not drawing_part:
            continue
        try:
            root = ET.fromstring(_read_bounded(zf, drawing_part,
                                               MAX_XML_BYTES))
        except (KeyError, ET.ParseError, ValueError):
            continue
        drawing_rels = _rels(zf, drawing_part)
        for tag in ("xdr:twoCellAnchor", "xdr:oneCellAnchor",
                    "xdr:absoluteAnchor"):
            for anchor in root.findall(tag, _NS):
                blip = anchor.find(".//a:blip", _NS)
                if blip is None:
                    continue
                media = drawing_rels.get(blip.get(_R_EMBED, ""))
                if not media:
                    continue
                try:
                    row1, row2, col1 = _anchor_rows(anchor)
                except (TypeError, ValueError):
                    continue
                anchors.append((row1, row2, col1, media))
                if len(anchors) >= MAX_ANCHORS_PER_SHEET:
                    log.warning("anchor cap reached on %r", sheet_part[:80])
                    return anchors
    return anchors


def extract_for(map_path, wanted: dict[str, dict[int, str]]) \
        -> dict[str, dict]:
    """Pull one image per base SKU out of the workbook.

    wanted: {sheet name: {1-based row: base}} — the rows where selected
    SKUs live. An image whose anchor starts on a wanted row (or spans it)
    is credited to that row's base; the first hit per base wins.

    Returns {base: {"data": base64 str, "ext": "png"|"jpeg"|...}}.
    """
    if not wanted:
        return {}
    images: dict[str, dict] = {}
    budget = MAX_TOTAL_IMAGE_BYTES
    try:
        with zipfile.ZipFile(map_path) as zf:
            parts = _sheet_parts(zf)
            for sheet_name, rows in wanted.items():
                part = parts.get(sheet_name)
                if not part:
                    continue
                for row1, row2, _col, media in sheet_anchors(zf, part):
                    if budget <= 0:
                        log.warning("total image budget exhausted")
                        return images
                    base = rows.get(row1)
                    if base is None and row2 > row1:
                        base = next((rows[r] for r in range(row1, row2 + 1)
                                     if r in rows), None)
                    if base is None or base in images:
                        continue
                    ext = posixpath.splitext(media)[1].lstrip(".").lower()
                    if ext not in _RASTER_EXT:
                        continue
                    try:
                        data = _read_bounded(zf, media,
                                             min(MAX_IMAGE_BYTES, budget))
                    except (KeyError, ValueError):
                        continue
                    if not data:
                        continue
                    budget -= len(data)
                    images[base] = {
                        "data": base64.b64encode(data).decode("ascii"),
                        "ext": "jpeg" if ext == "jpg" else ext,
                    }
    except (OSError, zipfile.BadZipFile, ValueError) as exc:
        log.warning("image extraction failed: %s", exc)
        return images
    if images:
        log.info("extracted %d product image(s)", len(images))
    return images


_ROW_RE = re.compile(r"(\d+)$")


def rows_by_base(details) -> dict[int, str]:
    """{row: base} from a list of (cell coordinate, sku, base) details."""
    out: dict[int, str] = {}
    for coord, _sku, base in details:
        match = _ROW_RE.search(coord)
        if match:
            out.setdefault(int(match.group(1)), base)
    return out
