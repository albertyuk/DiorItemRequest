"""Internal web tool: upload a Dior sell-thru map (+ optionally a stock query
export), extract yellow-highlighted SKUs, and download a filled ProductsList
workbook. See pipeline.py for the extraction/matching logic."""

from __future__ import annotations

import hmac
import json
import logging
import os
import re
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

import column_memory
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


def parse_users(raw: str) -> dict[str, str]:
    """APP_USERS='albert:pw1,vivian:pw2' -> {'albert': 'pw1', ...}.
    Passwords may contain ':'; malformed entries are skipped loudly."""
    users: dict[str, str] = {}
    for pair in raw.split(","):
        pair = pair.strip()
        if not pair:
            continue
        name, sep, password = pair.partition(":")
        if not sep or not name.strip() or not password:
            log.warning("ignoring malformed APP_USERS entry %r…", pair[:12])
            continue
        users[name.strip()] = password
    return users


def report_to_dict(report: "pipeline.RunReport", map_filename: str,
                   uploaded_by: str = "", built_by: str = "") -> dict:
    """JSON-serializable snapshot of a run, so the report page can be
    re-rendered later (and in either language)."""
    return {
        "map_filename": map_filename,
        "sheets": [s if isinstance(s, dict) else pipeline.sheet_to_dict(s)
                   for s in report.sheets],
        "ai_enabled": report.ai_enabled,
        "ai_detection": report.ai_detection,
        "ai_note": report.ai_note,
        "ai_confirmed": report.ai_confirmed,
        "bases": report.bases,
        "base_skus": report.base_skus,
        "base_sources": report.base_sources,
        "matched_counts": report.matched_counts,
        "matched_rows": report.matched_rows,
        "unmatched": report.unmatched,
        "excluded": report.excluded or [],
        "other_fills": report.other_fills,
        "rows_written": report.rows_written,
        "query_info": report.query_info,
        "uploaded_by": uploaded_by,
        "built_by": built_by,
    }


def default_data_dir() -> Path:
    if os.environ.get("DATA_DIR"):
        return Path(os.environ["DATA_DIR"])
    fly_volume = Path("/data")
    if fly_volume.is_dir():
        return fly_volume
    return Path("./data")


def create_app(data_dir: Path | str | None = None,
               password: str | None = None,
               users: dict[str, str] | None = None) -> Flask:
    data_dir = Path(data_dir) if data_dir else default_data_dir()
    if password is None:
        password = os.environ.get("APP_PASSWORD", "")
    if users is None:
        users = parse_users(os.environ.get("APP_USERS", ""))
    if os.environ.get("FLY_APP_NAME") and not (password or users):
        # This tool handles internal pricing data; it must never sit on a
        # public URL unauthenticated.
        raise RuntimeError(
            "Neither APP_USERS nor APP_PASSWORD is set. Refusing to start "
            "unauthenticated in production — run: fly secrets set "
            "APP_USERS='name:password,name2:password2' (or APP_PASSWORD=...)"
        )

    outputs_dir = data_dir / "outputs"
    tmp_dir = data_dir / "tmp"
    pending_dir = data_dir / "pending"
    stored_query = data_dir / "query.xlsx"
    query_meta_path = data_dir / "query_meta.json"
    memory_path = data_dir / "column_memory.json"
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

    def prune_drafts() -> None:
        files = sorted(pending_dir.glob("draft_*.json"),
                       key=lambda p: p.stat().st_mtime, reverse=True)
        for stale in files[KEEP_PENDING:]:
            stale.unlink(missing_ok=True)

    def recent_runs(limit: int = 10) -> list[dict]:
        """Newest stored runs for the front page: link report + download."""
        runs = []
        for path in sorted(outputs_dir.glob("report_*.json"),
                           key=lambda p: p.stat().st_mtime, reverse=True)[:limit]:
            run_id = path.stem[len("report_"):]
            try:
                data = json.loads(path.read_text())
            except (OSError, ValueError):
                continue
            runs.append({
                "run_id": run_id,
                "map_filename": data.get("map_filename", ""),
                "rows_written": data.get("rows_written", 0),
                "built_by": data.get("built_by", ""),
                "when": datetime.fromtimestamp(path.stat().st_mtime)
                        .strftime("%Y-%m-%d %H:%M"),
                "has_output": (outputs_dir / f"ProductsList_{run_id}.xlsx").exists(),
            })
        return runs

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

    def _build_and_store(draft: dict, selected, reviewed: bool,
                         built_by: str = "") -> str:
        """Build the workbook + report from a draft. Fast — runs in-request."""
        run_id = _new_id()
        out_path = outputs_dir / f"ProductsList_{run_id}.xlsx"
        ai = draft.get("ai") or {}
        report = pipeline.build_from_selection(
            template_path, out_path, draft, selected=selected,
            ai_info={"enabled": ai.get("enabled"),
                     "detection": ai.get("detection"),
                     "note": ai.get("note"),
                     "confirmed": reviewed and bool(ai.get("detection"))})
        # Persist the run report next to the output so the report page is
        # refresh-safe and re-renderable in either language.
        (outputs_dir / f"report_{run_id}.json").write_text(
            json.dumps(report_to_dict(report, draft.get("map_filename", ""),
                                      uploaded_by=draft.get("uploaded_by", ""),
                                      built_by=built_by),
                       ensure_ascii=False))
        prune_outputs()
        return run_id

    def _run_scan_job(job_id: str, map_path: Path, map_filename: str,
                      ai: dict, uploaded_by: str = "") -> None:
        """Scan + match in the background; the human review filters the
        result before anything is built."""
        job = jobs[job_id]
        weights = {"scan": (2, 60), "match": (60, 96)}

        def progress(stage, done, total):
            lo, hi = weights.get(stage, (0, 2))
            frac = min(done / total, 1.0) if total else 0.5
            job["stage"] = stage
            job["percent"] = round(lo + frac * (hi - lo))

        try:
            draft = pipeline.scan_and_match(
                map_path, stored_query, ai_detection=ai.get("detection"),
                progress=progress)
            draft.update({
                "map_filename": map_filename,
                "ai": ai,
                "query_info": read_query_meta(),
                "uploaded_by": uploaded_by,
            })
            if draft["bases"]:
                draft_id = _new_id()
                (pending_dir / f"draft_{draft_id}.json").write_text(
                    json.dumps(draft, ensure_ascii=False))
                prune_drafts()
                next_url = f"/review/{draft_id}"
            else:
                # nothing to review — build the (empty) output directly
                run_id = _build_and_store(draft, selected=None,
                                          reviewed=False)
                next_url = f"/report/{run_id}"
        except Exception as exc:
            # Malformed maps fail in many shapes (BadZipFile, XML
            # ParseError, KeyError, ...) — all must surface as a friendly
            # message on the progress page, never a dead bar.
            log.exception("processing failed")
            map_path.unlink(missing_ok=True)
            job.update(status="error", detail=str(exc))
            return
        map_path.unlink(missing_ok=True)  # the map is not needed after scan
        job.update(status="done", percent=100, next=next_url)

    def start_scan_job(map_path: Path, map_filename: str, ai: dict,
                       uploaded_by: str = "") -> str:
        purge_jobs()
        job_id = _new_id()
        with jobs_lock:
            jobs[job_id] = {"status": "running", "percent": 0,
                            "stage": "start", "created": time.time()}
        threading.Thread(
            target=_run_scan_job, daemon=True,
            args=(job_id, map_path, map_filename, ai, uploaded_by),
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
        if not password and not users:
            g.user = ""  # local development: open access, anonymous
            return None
        auth = request.authorization
        # Compare as bytes: compare_digest on str rejects non-ASCII input
        # with a TypeError, which would turn a login typo into a 500.
        if auth and auth.type == "basic" and auth.password:
            typed = (auth.username or "").strip()
            for name, user_password in users.items():
                if (name.lower() == typed.lower()
                        and hmac.compare_digest(
                            auth.password.encode("utf-8"),
                            user_password.encode("utf-8"))):
                    g.user = name  # canonical casing from the config
                    return None
            if password and hmac.compare_digest(
                    auth.password.encode("utf-8"),
                    password.encode("utf-8")):
                g.user = typed  # shared-password fallback: name as typed
                return None
        return Response(
            "Authentication required.", 401,
            {"WWW-Authenticate": 'Basic realm="Dior item request"'},
        )

    # --- routes -------------------------------------------------------------

    @app.get("/")
    def index():
        return render_template("index.html", query_meta=read_query_meta(),
                               recent=recent_runs())

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
        ai = {"enabled": False, "detection": None, "note": None,
              "samples": {}}
        if sku_locator.is_configured() or memory_path.exists():
            try:
                previews = sku_locator.build_previews(map_tmp)
                # sheets whose header row was approved on an earlier run
                # come from memory; only the rest go to the AI
                remembered, unknown = column_memory.lookup(previews,
                                                           memory_path)
                detection = list(remembered)
                if unknown and sku_locator.is_configured():
                    fresh = sku_locator.locate_from_previews(unknown)
                    fresh = [e for e in fresh if e.get("sku_columns")]
                    column_memory.attach_fingerprints(fresh, previews)
                    detection += fresh
                ai["enabled"] = (sku_locator.is_configured()
                                 or bool(remembered))
                ai["detection"] = detection
                ai["samples"] = sku_locator.column_samples(previews, detection)
                if remembered:
                    log.info("column memory: %d sheet(s) pre-approved",
                             len(remembered))
            except Exception as exc:
                log.warning("SKU column detection unavailable: %s", exc)
                ai["enabled"] = True
                ai["note"] = str(exc)

        job_id = start_scan_job(map_tmp, map_file.filename, ai,
                                uploaded_by=g.get("user", ""))
        return respond_next(url_for("progress_page", job_id=job_id))

    @app.get("/review/<draft_id>")
    def review_page(draft_id: str):
        """The single human checkpoint: AI-detected columns and every
        extracted SKU, each individually selectable before the build."""
        if not RUN_ID_RE.match(draft_id):
            abort(404)
        draft_path = pending_dir / f"draft_{draft_id}.json"
        if not draft_path.exists():
            abort(404)
        draft = json.loads(draft_path.read_text())
        entries = []
        for base, sheets in draft.get("bases", {}).items():
            entries.append({
                "base": base,
                "skus": draft.get("base_skus", {}).get(base, []),
                "sheets": sheets,
                "sources": draft.get("base_sources", {}).get(base, []),
                "rows": len(draft.get("matched", {}).get(base, [])),
            })
        entries.sort(key=lambda e: (e["rows"] == 0, e["base"]))
        ai = draft.get("ai") or {}
        return render_template(
            "review.html",
            draft_id=draft_id,
            map_filename=draft.get("map_filename", ""),
            entries=entries,
            total_rows=sum(e["rows"] for e in entries),
            matched_count=sum(1 for e in entries if e["rows"]),
            ai=ai,
            sheets=draft.get("sheets", []),
            query_info=draft.get("query_info"),
        )

    @app.post("/review/<draft_id>/build")
    def review_build(draft_id: str):
        """Build the workbook from the draft, keeping only selected SKUs."""
        if not RUN_ID_RE.match(draft_id):
            abort(404)
        draft_path = pending_dir / f"draft_{draft_id}.json"
        if not draft_path.exists():
            abort(404)
        draft = json.loads(draft_path.read_text())
        draft_path.unlink(missing_ok=True)  # single-use: replays 404
        selected = request.form.getlist("sku")
        log.info("review %s by %s: kept %d of %d bases", draft_id,
                 g.get("user", "") or "anonymous",
                 len(set(selected) & set(draft.get("bases", {}))),
                 len(draft.get("bases", {})))

        # Column memory: an approved sheet's mapping is stored for future
        # uploads; a sheet whose AI-found SKUs were ALL unticked is treated
        # as rejected and its stored mapping (if any) is dropped.
        selected_set = set(selected)
        ai_bases_by_sheet = {
            s["name"]: {detail[2] for detail in s.get(
                "ai_highlighted_cell_details", [])}
            for s in draft.get("sheets", [])
        }
        for entry in (draft.get("ai") or {}).get("detection") or []:
            sheet_ai_bases = ai_bases_by_sheet.get(entry.get("sheet"), set())
            if sheet_ai_bases and not (sheet_ai_bases & selected_set):
                column_memory.forget(memory_path, entry.get("fingerprint", ""))
            else:
                column_memory.remember(memory_path, entry,
                                       user=g.get("user", ""))

        run_id = _build_and_store(draft, selected=selected, reviewed=True,
                                  built_by=g.get("user", ""))
        return redirect(url_for("report_page", run_id=run_id))

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
                for key in ("status", "percent", "stage", "next", "detail")}

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
