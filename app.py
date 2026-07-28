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
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from flask import (
    Flask, Response, abort, redirect, render_template, request, send_file,
    url_for,
)
from werkzeug.exceptions import RequestEntityTooLarge

import pipeline

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s %(message)s")
log = logging.getLogger("app")

RUN_ID_RE = re.compile(r"^\d{8}_\d{6}$")
KEEP_OUTPUTS = 10
MAX_UPLOAD_BYTES = 200 * 1024 * 1024


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
            stale.unlink(missing_ok=True)

    def error_page(message: str, status: int = 400):
        return render_template("error.html", message=message), status

    # --- auth ---------------------------------------------------------------

    @app.before_request
    def require_auth():
        if not password:
            return None  # local development: open access
        auth = request.authorization
        if (auth and auth.type == "basic" and auth.password
                and hmac.compare_digest(auth.password, password)):
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
            return error_page("Please choose a sell-thru map file (.xlsx) — "
                              "it is required for every run.")

        query_file = request.files.get("query_file")
        if query_file is not None and query_file.filename:
            query_tmp = save_upload(query_file, "query")
            try:
                ok, message, row_count = pipeline.validate_query_file(query_tmp)
                if not ok:
                    return error_page(
                        f"The query file was rejected and the previously "
                        f"stored one (if any) was kept: {message}")
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
            return error_page(
                "No stock query export is stored yet. Upload one in the "
                "'Stock query export' field and try again.")

        map_tmp = save_upload(map_file, "map")
        run_at = datetime.now()
        run_id = run_at.strftime("%Y%m%d_%H%M%S")
        out_path = outputs_dir / f"ProductsList_{run_id}.xlsx"
        try:
            report = pipeline.run_pipeline(
                map_tmp, stored_query, template_path, out_path)
        except (zipfile.BadZipFile, KeyError, ValueError, OSError) as exc:
            out_path.unlink(missing_ok=True)
            log.warning("processing failed: %s", exc)
            return error_page(
                "The sell-thru map could not be processed as an .xlsx "
                f"workbook: {exc}")
        finally:
            map_tmp.unlink(missing_ok=True)

        prune_outputs()
        return render_template(
            "report.html",
            report=report,
            map_filename=map_file.filename,
            run_id=run_id,
            download_name=f"ProductsList_{run_at.strftime('%Y%m%d_%H%M')}.xlsx",
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
        return error_page(
            "Upload too large: the combined upload must stay under 200 MB.",
            413)

    @app.errorhandler(404)
    def not_found(_exc):
        return error_page("Not found — the file may have been pruned "
                          "(only the last 10 outputs are kept).", 404)

    return app


app = create_app()
