"""Security regressions.

Each test pins a specific hardening measure so a future refactor cannot
quietly remove it. The uploaded workbook is untrusted input, so most of
these feed deliberately hostile files through the real code paths.
"""

from __future__ import annotations

import io
import tempfile
import zipfile
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Color, PatternFill

import auth
import pipeline
import sheet_images
from app import create_app


def _yellow():
    return PatternFill(fill_type="solid", start_color=Color(theme=7, tint=0.8))


def _png(size=(60, 40)):
    from PIL import Image as PILImage
    buf = io.BytesIO()
    PILImage.new("RGB", size, (10, 20, 30)).save(buf, format="PNG")
    return buf.getvalue()


def _map_with_image(path, image_bytes=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "FW26"
    ws["A1"] = "MMC"
    ws["A2"] = "644S16B7E72X5805"
    ws["A2"].fill = _yellow()
    wb.save(path)
    # add the picture by hand so hostile payloads bypass openpyxl/PIL
    data = image_bytes if image_bytes is not None else _png()
    _inject_image(path, data)


def _inject_image(xlsx_path, media_bytes, media_name="image1.png"):
    """Put a picture into an existing xlsx via raw zip surgery, so the
    payload never passes through a library that would sanitize it."""
    media_path = f"xl/media/{media_name}"
    drawing = (
        '<?xml version="1.0"?><xdr:wsDr '
        'xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/'
        'spreadsheetDrawing" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/'
        'relationships">'
        '<xdr:oneCellAnchor><xdr:from><xdr:col>2</xdr:col><xdr:colOff>0'
        '</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff>'
        '</xdr:from><xdr:ext cx="100" cy="100"/><xdr:pic><xdr:nvPicPr>'
        '<xdr:cNvPr id="1" name="p"/><xdr:cNvPicPr/></xdr:nvPicPr>'
        '<xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill>'
        '<xdr:spPr/></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>'
        '</xdr:wsDr>')
    drawing_rels = (
        '<?xml version="1.0"?><Relationships xmlns="http://schemas.'
        'openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
        f'officeDocument/2006/relationships/image" Target="../media/'
        f'{media_name}"/></Relationships>')
    sheet_rels = (
        '<?xml version="1.0"?><Relationships xmlns="http://schemas.'
        'openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rIdD1" Type="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships/drawing" '
        'Target="../drawings/drawing1.xml"/></Relationships>')

    src = zipfile.ZipFile(xlsx_path)
    items = {i.filename: src.read(i.filename) for i in src.infolist()}
    src.close()
    sheet_name = next(n for n in items if n.startswith("xl/worksheets/sheet"))
    sheet_xml = items[sheet_name].decode()
    if "<drawing" not in sheet_xml:
        sheet_xml = sheet_xml.replace(
            "</worksheet>", '<drawing r:id="rIdD1"/></worksheet>')
        items[sheet_name] = sheet_xml.encode()
    items["xl/drawings/drawing1.xml"] = drawing.encode()
    items["xl/drawings/_rels/drawing1.xml.rels"] = drawing_rels.encode()
    items[f"xl/worksheets/_rels/{Path(sheet_name).name}.rels"] = \
        sheet_rels.encode()
    items[media_path] = media_bytes

    with zipfile.ZipFile(xlsx_path, "w", zipfile.ZIP_DEFLATED) as out:
        for name, data in items.items():
            out.writestr(name, data)


def _wanted():
    return {"FW26": {2: "644S16B7E72"}}


# --- untrusted workbook: decompression bombs & crafted parts ---------------

def test_zip_bomb_media_is_refused_not_inflated(tmp_path):
    """A media entry that inflates to ~500 MB must be skipped on its
    declared size — never read into memory."""
    path = tmp_path / "bomb.xlsx"
    _map_with_image(path, image_bytes=b"\0" * (500 * 1024 * 1024))
    compressed = path.stat().st_size
    assert compressed < 5 * 1024 * 1024, "fixture should be a small file"

    images = sheet_images.extract_for(path, _wanted())
    assert images == {}, "the bomb must not be extracted"


def test_total_image_budget_is_bounded(tmp_path, monkeypatch):
    """Many individually-legal images cannot sum past the run budget."""
    monkeypatch.setattr(sheet_images, "MAX_TOTAL_IMAGE_BYTES", 1000)
    monkeypatch.setattr(sheet_images, "MAX_IMAGE_BYTES", 800)
    path = tmp_path / "many.xlsx"
    _map_with_image(path, image_bytes=b"\0" * 900)
    assert sheet_images.extract_for(path, _wanted()) == {}


def test_rels_target_escaping_the_package_is_dropped(tmp_path):
    """A crafted rels Target must not resolve outside the package."""
    path = tmp_path / "trav.xlsx"
    _map_with_image(path)
    src = zipfile.ZipFile(path)
    items = {i.filename: src.read(i.filename) for i in src.infolist()}
    src.close()
    items["xl/drawings/_rels/drawing1.xml.rels"] = (
        '<?xml version="1.0"?><Relationships xmlns="http://schemas.'
        'openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships/image" '
        'Target="../../../../../../etc/passwd"/></Relationships>'
    ).encode()
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as out:
        for name, data in items.items():
            out.writestr(name, data)

    assert sheet_images.extract_for(path, _wanted()) == {}


def test_oversized_pixel_dimensions_dropped_at_build(tmp_path):
    """A small file declaring enormous dimensions (a decompression bomb)
    is dropped rather than handed to the image layer."""
    import base64

    from PIL import Image as PILImage
    buf = io.BytesIO()
    # 8k x 8k = 64M pixels: over OUR 40M ceiling but under Pillow's own
    # bomb threshold, so this pins our check rather than Pillow's.
    PILImage.new("1", (8000, 8000), 0).save(buf, format="PNG")
    huge = buf.getvalue()
    assert len(huge) < 1024 * 1024

    draft = {
        "sheets": [],
        "bases": {"644S16B7E72": ["FW26"]},
        "matched": {"644S16B7E72": [[
            "361", "644S16B7E72 - X5805 - T36", "d", "dep", "s",
            None, None, None, 1, 10, 20, "t", "p"]]},
        "images": {"644S16B7E72": {
            "data": base64.b64encode(huge).decode(),
            "ext": "png"}},
    }
    template = Path(__file__).parent / "assets" / "ProductsListTemplate.xlsx"
    out = tmp_path / "out.xlsx"
    report = pipeline.build_from_selection(template, out, draft)
    assert report.images_count == 0          # refused
    assert report.rows_written == 1          # the run still succeeds
    assert len(load_workbook(out).active._images) == 0


def test_malformed_drawing_never_fails_a_run(tmp_path):
    path = tmp_path / "bad.xlsx"
    _map_with_image(path)
    src = zipfile.ZipFile(path)
    items = {i.filename: src.read(i.filename) for i in src.infolist()}
    src.close()
    items["xl/drawings/drawing1.xml"] = b"<xdr:wsDr><<<not xml"
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as out:
        for name, data in items.items():
            out.writestr(name, data)
    assert sheet_images.extract_for(path, _wanted()) == {}


# --- web layer -------------------------------------------------------------

def _app(tmp_path, **kw):
    kw.setdefault("setup_code", "root-code")
    kw.setdefault("secret", "s" * 64)
    kw.setdefault("cookie_secure", False)
    return create_app(data_dir=tmp_path / "data", **kw)


def test_security_headers_on_every_response(tmp_path):
    client = _app(tmp_path).test_client()
    for path in ("/login", "/healthz", "/setup"):
        r = client.get(path)
        csp = r.headers.get("Content-Security-Policy", "")
        assert "default-src 'self'" in csp
        assert "frame-ancestors 'none'" in csp
        assert "object-src 'none'" in csp
        assert r.headers["X-Content-Type-Options"] == "nosniff"
        assert r.headers["X-Frame-Options"] == "DENY"
        assert r.headers["Referrer-Policy"] == "strict-origin-when-cross-origin"


def test_hsts_only_under_tls(tmp_path):
    plain = _app(tmp_path, cookie_secure=False).test_client()
    assert "Strict-Transport-Security" not in plain.get("/login").headers

    secure = _app(tmp_path / "s", cookie_secure=True).test_client()
    r = secure.get("/login", base_url="https://example.test")
    assert "max-age=31536000" in r.headers["Strict-Transport-Security"]
    # behind a TLS-terminating proxy the forwarded header is the signal
    r = secure.get("/login", headers={"X-Forwarded-Proto": "https"})
    assert "Strict-Transport-Security" in r.headers


def test_lang_cookie_is_httponly(tmp_path):
    client = _app(tmp_path).test_client()
    cookie = client.get("/login?lang=zh").headers.get("Set-Cookie", "")
    assert "HttpOnly" in cookie and "SameSite=Lax" in cookie


def test_oversized_form_body_rejected(tmp_path):
    """A multi-megabyte credential field must not be parsed into memory."""
    client = _app(tmp_path).test_client()
    r = client.post("/login", data={"username": "u",
                                    "password": "x" * 5_000_000})
    assert r.status_code == 413


def test_stuck_and_stale_jobs_are_purged(tmp_path):
    """A worker killed mid-scan leaves a 'running' entry behind; those
    must age out too, or the registry grows without bound."""
    import time as _time

    app = _app(tmp_path)
    jobs = app.extensions["jobs"]
    purge = app.extensions["purge_jobs"]
    now = _time.time()
    jobs.update({
        "fresh_running": {"status": "running", "created": now},
        "stuck_running": {"status": "running", "created": now - 13 * 3600},
        "recent_done": {"status": "done", "created": now - 60},
        "stale_done": {"status": "done", "created": now - 3 * 3600},
    })
    purge()
    assert set(jobs) == {"fresh_running", "recent_done"}

    # and an absolute ceiling regardless of age
    for i in range(400):
        jobs[f"j{i}"] = {"status": "done", "created": now}
    purge()
    assert len(jobs) <= 200


# --- auth input bounds -----------------------------------------------------

def test_credential_inputs_are_bounded():
    assert len(auth.normalize_username("a" * 5000)) <= auth.MAX_USERNAME_LEN
    assert len(auth.clamp_password("p" * 10_000)) == auth.MAX_PASSWORD_LEN
    # an absurdly long password is refused rather than hashed
    assert not auth.valid_password("p" * (auth.MAX_PASSWORD_LEN + 1))
    assert auth.valid_password("p" * 12)


def test_long_username_cannot_bloat_throttle_keys(tmp_path):
    app = _app(tmp_path)
    client = app.test_client()
    client.post("/login", data={"username": "u" * 10_000, "password": "x"})
    ctx = app.extensions["auth"]
    assert all(len(key) <= auth.MAX_USERNAME_LEN
               for scope, key in ctx.throttle._buckets if scope == "user")


def test_client_ip_header_ignored_off_fly(tmp_path):
    """Off Fly's edge the client-IP header is client input: believing it
    would hand an attacker a fresh throttle bucket per request."""
    app = _app(tmp_path)
    ctx = app.extensions["auth"]
    assert ctx.trust_ip_header is False
    client = app.test_client()
    for i in range(25):
        client.post("/login", headers={"Fly-Client-IP": f"9.9.9.{i}"},
                    data={"username": f"user{i}", "password": "wrong-pass"})
    # all failures landed in ONE bucket (the socket peer), not 25
    ip_buckets = [k for scope, k in ctx.throttle._buckets if scope == "ip"]
    assert len(ip_buckets) == 1
    r = client.post("/login", headers={"Fly-Client-IP": "9.9.9.250"},
                    data={"username": "boss", "password": "whatever1"})
    assert r.status_code == 429           # still throttled despite the header


def test_client_ip_header_trusted_behind_fly(tmp_path, monkeypatch):
    monkeypatch.setenv("FLY_APP_NAME", "dior-item-request")
    app = _app(tmp_path / "fly")
    assert app.extensions["auth"].trust_ip_header is True


def test_state_files_are_owner_only(tmp_path, monkeypatch):
    monkeypatch.delenv("APP_SECRET", raising=False)
    # no explicit secret: the app generates and persists one
    app = create_app(data_dir=tmp_path / "data", setup_code="root-code",
                     cookie_secure=False)
    client = app.test_client()
    client.post("/setup", data={"code": "root-code", "username": "boss",
                                "password": "longenough1"})
    db_path = tmp_path / "data" / "auth.db"
    assert db_path.exists()
    assert db_path.stat().st_mode & 0o077 == 0, "auth.db must be owner-only"
    secret = tmp_path / "data" / "session_secret"
    assert secret.exists()
    assert secret.stat().st_mode & 0o077 == 0
    assert (tmp_path / "data").stat().st_mode & 0o077 == 0
