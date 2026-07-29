"""Remember human-approved SKU-column mappings, keyed by a fingerprint of
the sheet's header row.

When a run is built, the AI-detected columns the reviewer effectively kept
are stored per header fingerprint. On the next upload, any sheet whose
header row matches a stored fingerprint gets its columns from memory —
pre-approved, no API call needed for that sheet. If the reviewer excludes
every AI-found SKU of a sheet, its stored mapping is dropped (they changed
their mind).

The fingerprint includes column letters, so a layout change (an inserted
column shifts the letters) safely misses and falls back to fresh detection.
"""

from __future__ import annotations

import hashlib
import json
import logging
import threading
import time

log = logging.getLogger(__name__)

_lock = threading.Lock()


def _norm(value) -> str:
    return " ".join(str(value).split()).lower()


def fingerprint_row(cells: dict) -> str:
    """Stable hash of one preview row: {column letter: value}."""
    parts = [f"{letter}:{_norm(value)}"
             for letter, value in sorted(cells.items(),
                                         key=lambda kv: (len(kv[0]), kv[0]))
             if str(value).strip()]
    if len(parts) < 2:  # a near-empty row is no identity
        return ""
    return hashlib.sha256("|".join(parts).encode()).hexdigest()


def _load(path) -> dict:
    try:
        return json.loads(path.read_text())
    except (OSError, ValueError):
        return {}


def attach_fingerprints(detection: list, previews: list) -> None:
    """Annotate fresh AI detection entries with their sheet's header-row
    fingerprint (and the header cells), so the build step can remember
    them if the reviewer approves."""
    rows_by_sheet = {p["sheet"]: p["rows"] for p in previews}
    for entry in detection:
        rows = rows_by_sheet.get(entry["sheet"]) or {}
        cells = rows.get(str(entry.get("header_row"))) or {}
        fp = fingerprint_row(cells)
        if fp:
            entry["fingerprint"] = fp
            entry["header_cells"] = cells


def lookup(previews: list, path) -> tuple[list, list]:
    """Split previews into (remembered detection entries, unknown previews).

    A sheet matches when one of its preview rows fingerprints to a stored
    mapping AND the stored column letters still carry the same header text
    (belt and braces against hash collisions or stale entries)."""
    store = _load(path)
    remembered, unknown = [], []
    for sheet in previews:
        hit = None
        if store:
            for row_num, cells in sheet["rows"].items():
                entry = store.get(fingerprint_row(cells))
                if not entry:
                    continue
                headers_ok = all(
                    col.get("header") is None
                    or _norm(cells.get(col["column"], "")) == _norm(col["header"])
                    for col in entry["sku_columns"])
                if headers_ok:
                    hit = {
                        "sheet": sheet["sheet"],
                        "header_row": int(row_num),
                        "sku_columns": entry["sku_columns"],
                        "source": "memory",
                        "fingerprint": fingerprint_row(cells),
                        "header_cells": cells,
                    }
                    break
        if hit:
            remembered.append(hit)
        else:
            unknown.append(sheet)
    return remembered, unknown


def remember(path, entry: dict, user: str = "") -> None:
    """Store (or refresh) the approved columns for one detection entry."""
    fp = entry.get("fingerprint")
    if not fp:
        return
    with _lock:
        store = _load(path)
        store[fp] = {
            "sheet": entry.get("sheet", ""),
            "headers": entry.get("header_cells", {}),
            "sku_columns": [
                {"column": c["column"], "header": c.get("header"),
                 "reason": c.get("reason", "")}
                for c in entry.get("sku_columns", [])
            ],
            "saved_at": time.strftime("%Y-%m-%d %H:%M"),
            "saved_by": user or "",
        }
        path.write_text(json.dumps(store, ensure_ascii=False))
    log.info("column memory: remembered %s (%s)", entry.get("sheet"),
             [c["column"] for c in entry.get("sku_columns", [])])


def forget(path, fingerprint: str) -> None:
    if not fingerprint:
        return
    with _lock:
        store = _load(path)
        if store.pop(fingerprint, None) is not None:
            path.write_text(json.dumps(store, ensure_ascii=False))
            log.info("column memory: forgot one mapping")
