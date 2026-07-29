"""Internal web tool: upload a Dior sell-thru map (+ optionally a stock query
export), extract yellow-highlighted SKUs, and download a filled ProductsList
workbook. See pipeline.py for the extraction/matching logic."""

from __future__ import annotations

import hmac
import json
import logging
import os
import re
import shutil
import threading
import time
import uuid
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from flask import (
    Flask, Response, abort, g, redirect, render_template, request, send_file,
    url_for,
)
from werkzeug.exceptions import RequestEntityTooLarge

import pipeline
import sku_locator
from translations import STRINGS, SUPPORTED_LANGS

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s %(message)s")
log = logging.getLogger("app")

RUN_ID_RE = re.compile(r"^\d{8}_\d{6}_[0-9a-f]{6}$")
KEEP_OUTPUTS = 10
KEEP_PENDING = 5  # uploads awaiting human column verification
MAX_UPLOAD_BYTES = 200 * 1024 * 1024


def _new_id() -> str:
    return f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:6]}"


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
                "ai_sku_cells": s.ai_sku_cells,
                "ai_highlighted_cell_details": s.ai_highlighted_cell_details,
            }
            for s in report.sheets
        ],
        "ai_enabled": report.ai_enabled,
        "ai_detection": report.ai_detection,
        "ai_note": report.ai_note,
        "ai_confirmed": report.ai_confirmed,
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
    pending_dir = data_dir / "pending"
    stored_query = data_dir / "query.xlsx"
    query_meta_path = data_dir / "query_meta.json"
    template_path = Path(__file__).parent / "assets" / "ProductsListTemplate.xlsx"
    for d in (data_dir, outputs_dir, tmp_dir, pending_dir):
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

    def prune_pending() -> None:
        files = sorted(pending_dir.glob("map_*.xlsx"),
                       key=lambda p: p.stat().st_mtime, reverse=True)
        for stale in files[KEEP_PENDING:]:
            pending_id = stale.stem[len("map_"):]
            (pending_dir / f"pending_{pending_id}.json").unlink(missing_ok=True)
            (pending_dir / f"query_{pending_id}.xlsx").unlink(missing_ok=True)
            stale.unlink(missing_ok=True)

    def error_page(key: str, status: int = 400, **fmt):
        message = tr(key, **fmt)
        if request.headers.get("X-Requested-With") == "XMLHttpRequest":
            return {"error": message}, status
        return render_template("error.html", message=message), status

    def respond_next(url: str):
        """302 for plain form posts; JSON for the fetch/XHR submission that
        drives the client-side upload progress bar."""
        if request.headers.get("X-Requested-With") == "XMLHttpRequest":
            return {"next": url}
        return redirect(url)

    # --- background processing jobs -----------------------------------------
    # The heavy pipeline runs in a background thread so the browser can poll
    # a real progress bar instead of hanging on one long request. One
    # gunicorn worker process serves the app, so this in-process registry is
    # the whole truth.
    jobs: dict[str, dict] = {}
    jobs_lock = threading.Lock()

    def purge_jobs() -> None:
        cutoff = time.time() - 2 * 3600
        with jobs_lock:
            for job_id in [j for j, job in jobs.items()
                           if job.get("created", 0) < cutoff
                           and job.get("status") != "running"]:
                del jobs[job_id]

    def _run_job(job_id: str, map_path: Path, map_filename: str,
                 ai_result: dict, query_snapshot: Path | None) -> None:
        job = jobs[job_id]
        # stage -> (start%, end%) of the overall bar, weighted by how long
        # each stage takes on real workbooks
        weights = {"scan": (2, 60), "match": (60, 92), "write": (92, 99)}

        def progress(stage, done, total):
            lo, hi = weights.get(stage, (0, 2))
            frac = min(done / total, 1.0) if total else 0.5
            job["stage"] = stage
            job["percent"] = round(lo + frac * (hi - lo))

        def cleanup():
            map_path.unlink(missing_ok=True)
            if query_snapshot is not None:
                query_snapshot.unlink(missing_ok=True)

        run_id = _new_id()
        out_path = outputs_dir / f"ProductsList_{run_id}.xlsx"
        try:
            report = pipeline.run_pipeline(
                map_path, query_snapshot or stored_query, template_path,
                out_path, ai_result=ai_result, progress=progress)
            # Persist the run report next to the output so the report page
            # is refresh-safe and re-renderable in either language.
            (outputs_dir / f"report_{run_id}.json").write_text(
                json.dumps(report_to_dict(report, map_filename),
                           ensure_ascii=False))
            prune_outputs()
        except Exception as exc:
            # Malformed maps fail in many shapes (BadZipFile, XML
            # ParseError, KeyError, ...) — all must surface as a friendly
            # message on the progress page, never a dead bar.
            out_path.unlink(missing_ok=True)
            log.exception("processing failed")
            cleanup()
            job.update(status="error", detail=str(exc))
            return
        cleanup()  # before flipping status: pollers may react instantly
        job.update(status="done", percent=100, run_id=run_id)

    def start_job(map_path: Path, map_filename: str, ai_result: dict,
                  query_snapshot: Path | None = None) -> str:
        purge_jobs()
        job_id = _new_id()
        with jobs_lock:
            jobs[job_id] = {"status": "running", "percent": 0,
                            "stage": "start", "created": time.time()}
        threading.Thread(
            target=_run_job, daemon=True,
            args=(job_id, map_path, map_filename, ai_result, query_snapshot),
        ).start()
        return job_id

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

    @app.get("/help")
    def help_page():
        return render_template("help.html")

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
        if not zipfile.is_zipfile(map_tmp):
            map_tmp.unlink(missing_ok=True)
            return error_page("err_map_failed",
                              detail="not a valid .xlsx workbook")
        ai_result = {"enabled": False, "detection": None, "note": None,
                     "confirmed": False}
        if sku_locator.is_configured():
            ai_result["enabled"] = True
            try:
                previews = sku_locator.build_previews(map_tmp)
                detection = sku_locator.locate_from_previews(previews)
            except Exception as exc:
                log.warning("SKU column detection unavailable: %s", exc)
                ai_result["note"] = str(exc)
            else:
                # Defense in depth: _normalize already drops sheets with no
                # SKU columns, but never pause verification unless there is
                # at least one actual column to verify.
                detection = [e for e in detection if e.get("sku_columns")]
                if detection:
                    # Human verification: park the upload and show the
                    # detected columns (with sample values) before the
                    # heavy processing run uses them.
                    pending_id = _new_id()
                    map_tmp.replace(pending_dir / f"map_{pending_id}.xlsx")
                    # Snapshot the query too: the confirmed run must use
                    # the query that was current at upload time, even if
                    # someone replaces the stored one during the pause.
                    shutil.copyfile(stored_query,
                                    pending_dir / f"query_{pending_id}.xlsx")
                    samples = sku_locator.column_samples(previews, detection)
                    (pending_dir / f"pending_{pending_id}.json").write_text(
                        json.dumps({
                            "map_filename": map_file.filename,
                            "detection": detection,
                            "samples": samples,
                        }, ensure_ascii=False))
                    prune_pending()
                    return respond_next(
                        url_for("confirm_page", pending_id=pending_id))
                ai_result["detection"] = detection  # []: nothing detected

        job_id = start_job(map_tmp, map_file.filename, ai_result)
        return respond_next(url_for("progress_page", job_id=job_id))

    @app.get("/confirm/<pending_id>")
    def confirm_page(pending_id: str):
        """The human-verification page — a real URL, so it is refresh-safe
        and reachable after the upload's POST/redirect."""
        if not RUN_ID_RE.match(pending_id):
            abort(404)
        meta_path = pending_dir / f"pending_{pending_id}.json"
        if not (meta_path.exists()
                and (pending_dir / f"map_{pending_id}.xlsx").exists()):
            abort(404)
        meta = json.loads(meta_path.read_text())
        return render_template(
            "confirm.html",
            pending_id=pending_id,
            map_filename=meta.get("map_filename", ""),
            detection=meta.get("detection", []),
            samples=meta.get("samples", {}),
        )

    @app.post("/process/<pending_id>")
    def process_confirm(pending_id: str):
        """Phase 2: the human reviewed the AI-detected columns; run the
        pipeline with only the columns they kept ticked."""
        if not RUN_ID_RE.match(pending_id):
            abort(404)
        map_path = pending_dir / f"map_{pending_id}.xlsx"
        meta_path = pending_dir / f"pending_{pending_id}.json"
        if not (map_path.exists() and meta_path.exists()):
            abort(404)
        meta = json.loads(meta_path.read_text())
        meta_path.unlink(missing_ok=True)

        selected = set(request.form.getlist("col"))
        confirmed = []
        for si, entry in enumerate(meta.get("detection", [])):
            kept = [col for ci, col in enumerate(entry["sku_columns"])
                    if f"{si}:{ci}" in selected]
            if kept:
                confirmed.append({**entry, "sku_columns": kept})
        log.info("column verification %s: kept %s", pending_id,
                 [(e["sheet"], [c["column"] for c in e["sku_columns"]])
                  for e in confirmed])
        query_snapshot = pending_dir / f"query_{pending_id}.xlsx"
        snapshot = query_snapshot if query_snapshot.exists() else None
        if snapshot is None and not stored_query.exists():
            map_path.unlink(missing_ok=True)
            return error_page("err_no_query")
        job_id = start_job(
            map_path, meta.get("map_filename", ""),
            {"enabled": True, "detection": confirmed, "note": None,
             "confirmed": True},
            query_snapshot=snapshot)
        return respond_next(url_for("progress_page", job_id=job_id))

    @app.get("/progress/<job_id>")
    def progress_page(job_id: str):
        if not RUN_ID_RE.match(job_id) or job_id not in jobs:
            abort(404)
        return render_template("progress.html", job_id=job_id)

    @app.get("/progress/<job_id>/status")
    def progress_status(job_id: str):
        job = jobs.get(job_id)
        if job is None:
            return {"status": "lost"}, 404
        return {key: job.get(key)
                for key in ("status", "percent", "stage", "run_id", "detail")}

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
