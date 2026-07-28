"""Internal web tool: upload a Dior sell-thru map (+ optionally a stock query
export), extract yellow-highlighted SKUs, and download a filled ProductsList
workbook. See pipeline.py for the extraction/matching logic."""

from __future__ import annotations

import hmac
import json
import logging
import os
import re
import uuid
from datetime import datetime, timezone
from pathlib import Path

from flask import (
    Flask, Response, abort, g, redirect, render_template, request, send_file,
    url_for,
)
from werkzeug.exceptions import RequestEntityTooLarge

import pipeline
from translations import STRINGS, SUPPORTED_LANGS

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s %(message)s")
log = logging.getLogger("app")

RUN_ID_RE = re.compile(r"^\d{8}_\d{6}_[0-9a-f]{6}$")
KEEP_OUTPUTS = 10
MAX_UPLOAD_BYTES = 200 * 1024 * 1024


def report_to_dict(report: "pipeline.RunReport", map_filename: str) -> dict:
    """JSON-serializable snapshot of a run, so the report page can be
    re-rendered later (and in either language)."""
    return {
        "map_filename": map_filename,
        "sheets": [
            {
                "name": s.name,
                "sku_cells": s.sku_cells,
                "highlighted_cells": s.highlighted_cells,
                "highlighted_skus": sorted(s.highlighted_skus),
                "bases": sorted(s.bases),
                "highlighted_cell_details": s.highlighted_cell_details,
                "other_fills": s.other_fills,
            }
            for s in report.sheets
        ],
        "bases": report.bases,
        "base_skus": report.base_skus,
        "matched_counts": report.matched_counts,
        "matched_rows": report.matched_rows,
        "unmatched": report.unmatched,
        "other_fills": report.other_fills,
        "rows_written": report.rows_written,
    }


def default_data_dir() -> Path:
    if os.environ.get("DATA_DIR"):
        return Path(os.environ["DATA_DIR"])
    fly_volume = Path("/data")
    if fly_volume.is_dir():
        return fly_volume
    return Path("./data")


def create_app(data_dir: Path | str | None = None,
               password: str | None = None) -> Flask:
    data_dir = Path(data_dir) if data_dir else default_data_dir()
    if password is None:
        password = os.environ.get("APP_PASSWORD", "")
    if os.environ.get("FLY_APP_NAME") and not password:
        # This tool handles internal pricing data; it must never sit on a
        # public URL unauthenticated.
        raise RuntimeError(
            "APP_PASSWORD is not set. Refusing to start unauthenticated in "
            "production — run: fly secrets set APP_PASSWORD=..."
        )

    outputs_dir = data_dir / "outputs"
    tmp_dir = data_dir / "tmp"
    stored_query = data_dir / "query.xlsx"
    query_meta_path = data_dir / "query_meta.json"
    template_path = Path(__file__).parent / "assets" / "ProductsListTemplate.xlsx"
    for d in (data_dir, outputs_dir, tmp_dir):
        d.mkdir(parents=True, exist_ok=True)

    app = Flask(__name__)
    app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_BYTES

    # --- helpers ------------------------------------------------------------

    def read_query_meta() -> dict | None:
        if not (stored_query.exists() and query_meta_path.exists()):
            return None
        try:
            return json.loads(query_meta_path.read_text())
        except (OSError, ValueError):
            return None

    def save_upload(file_storage, label: str) -> Path:
        path = tmp_dir / f"{label}_{uuid.uuid4().hex}.xlsx"
        file_storage.save(path)
        return path

    def prune_outputs() -> None:
        files = sorted(outputs_dir.glob("ProductsList_*.xlsx"),
                       key=lambda p: p.stat().st_mtime, reverse=True)
        for stale in files[KEEP_OUTPUTS:]:
            run_id = stale.stem[len("ProductsList_"):]
            (outputs_dir / f"report_{run_id}.json").unlink(missing_ok=True)
            stale.unlink(missing_ok=True)

    def tr(key: str, **fmt) -> str:
        lang = getattr(g, "lang", "en")
        s = STRINGS.get(lang, STRINGS["en"]).get(key) or STRINGS["en"][key]
        return s.format(**fmt) if fmt else s

    def error_page(key: str, status: int = 400, **fmt):
        return render_template("error.html", message=tr(key, **fmt)), status

    # --- language -----------------------------------------------------------

    @app.before_request
    def resolve_lang():
        lang = request.args.get("lang")
        if lang not in SUPPORTED_LANGS:
            lang = request.cookies.get("lang")
        g.lang = lang if lang in SUPPORTED_LANGS else "en"

    @app.after_request
    def remember_lang(resp):
        lang = request.args.get("lang")
        if lang in SUPPORTED_LANGS:
            resp.set_cookie("lang", lang, max_age=365 * 24 * 3600,
                            samesite="Lax")
        return resp

    @app.context_processor
    def inject_i18n():
        return {"t": tr, "lang": getattr(g, "lang", "en")}

    # --- auth ---------------------------------------------------------------

    @app.before_request
    def require_auth():
        if not password:
            return None  # local development: open access
        auth = request.authorization
        # Compare as bytes: compare_digest on str rejects non-ASCII input
        # with a TypeError, which would turn a login typo into a 500.
        if (auth and auth.type == "basic" and auth.password
                and hmac.compare_digest(auth.password.encode("utf-8"),
                                        password.encode("utf-8"))):
            return None
        return Response(
            "Authentication required.", 401,
            {"WWW-Authenticate": 'Basic realm="Dior item request"'},
        )

    # --- routes -------------------------------------------------------------

    @app.get("/")
    def index():
        return render_template("index.html", query_meta=read_query_meta())

    @app.post("/process")
    def process():
        map_file = request.files.get("map_file")
        if map_file is None or not map_file.filename:
            return error_page("err_map_required")

        query_file = request.files.get("query_file")
        if query_file is not None and query_file.filename:
            query_tmp = save_upload(query_file, "query")
            try:
                ok, message, row_count = pipeline.validate_query_file(query_tmp)
                if not ok:
                    return error_page("err_query_rejected", detail=message)
                query_tmp.replace(stored_query)
                query_meta_path.write_text(json.dumps({
                    "filename": query_file.filename,
                    "uploaded_at": datetime.now(timezone.utc)
                                   .strftime("%Y-%m-%d %H:%M UTC"),
                    "rows": row_count,
                }))
                log.info("stored new query file %r (%d rows)",
                         query_file.filename, row_count)
            finally:
                query_tmp.unlink(missing_ok=True)

        if not stored_query.exists():
            return error_page("err_no_query")

        map_tmp = save_upload(map_file, "map")
        run_at = datetime.now()
        # The uuid suffix keeps concurrent runs (2 gunicorn threads) from
        # sharing an output path and serving each other's workbooks.
        run_id = f"{run_at.strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:6]}"
        out_path = outputs_dir / f"ProductsList_{run_id}.xlsx"
        try:
            report = pipeline.run_pipeline(
                map_tmp, stored_query, template_path, out_path)
        except Exception as exc:
            # Malformed maps fail in many shapes (BadZipFile, XML
            # ParseError, KeyError, ...) — all must land on the friendly
            # error page, never the bare 500.
            out_path.unlink(missing_ok=True)
            log.exception("processing failed")
            return error_page("err_map_failed", detail=str(exc))
        finally:
            map_tmp.unlink(missing_ok=True)

        # Persist the run report next to the output so the report page is
        # refresh-safe and can be re-rendered later in either language.
        (outputs_dir / f"report_{run_id}.json").write_text(
            json.dumps(report_to_dict(report, map_file.filename),
                       ensure_ascii=False))
        prune_outputs()
        return redirect(url_for("report_page", run_id=run_id))

    @app.get("/report/<run_id>")
    def report_page(run_id: str):
        if not RUN_ID_RE.match(run_id):
            abort(404)
        report_path = outputs_dir / f"report_{run_id}.json"
        if not report_path.exists():
            abort(404)
        data = json.loads(report_path.read_text())
        return render_template(
            "report.html",
            report=data,
            map_filename=data.get("map_filename", ""),
            run_id=run_id,
            download_name=f"ProductsList_{run_id[:13]}.xlsx",
        )

    @app.get("/download/<run_id>")
    def download(run_id: str):
        if not RUN_ID_RE.match(run_id):
            abort(404)
        path = outputs_dir / f"ProductsList_{run_id}.xlsx"
        if not path.exists():
            abort(404)
        return send_file(
            path,
            as_attachment=True,
            download_name=f"ProductsList_{run_id[:13]}.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument"
                     ".spreadsheetml.sheet",
        )

    @app.errorhandler(RequestEntityTooLarge)
    def too_large(_exc):
        return error_page("err_too_large", 413)

    @app.errorhandler(404)
    def not_found(_exc):
        return error_page("err_not_found", 404)

    @app.errorhandler(500)
    def internal_error(_exc):
        return error_page("err_server", 500)

    return app


app = create_app()
