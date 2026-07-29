"""Acceptance tests against the real sample workbooks in ./samples/ plus
generated fixtures for layouts the sample copy does not contain.

The real samples hold internal pricing data and are not committed; tests that
need them skip when ./samples/ is absent.
"""

from __future__ import annotations

import base64
import io
import re
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Color, PatternFill

import pipeline
import sku_locator
from app import create_app

SAMPLES = Path(__file__).parent / "samples"
MAP = SAMPLES / "sell_thru_map.xlsx"
QUERY = SAMPLES / "query.xlsx"
TEMPLATE = SAMPLES / "ProductsListTemplate.xlsx"

needs_samples = pytest.mark.skipif(
    not (MAP.exists() and QUERY.exists() and TEMPLATE.exists()),
    reason="real sample workbooks not present in ./samples/",
)


@pytest.fixture(scope="session")
def map_scan():
    return pipeline.scan_sell_thru_map(MAP)


@pytest.fixture(scope="session")
def fw26_matched(map_scan):
    fw26 = next(s for s in map_scan.sheets if s.name == "FW26")
    return pipeline.match_query(QUERY, fw26.bases)


# --- 1. highlight extraction counts ----------------------------------------

@needs_samples
def test_fw26_highlight_counts(map_scan):
    fw26 = next(s for s in map_scan.sheets if s.name == "FW26")
    assert fw26.highlighted_cells == 43
    assert len(fw26.bases) == 41
    # the structural theme-4 banding on size rows must never be flagged
    assert fw26.other_fills == []


@needs_samples
def test_look_total_highlight_counts(map_scan):
    lt = next((s for s in map_scan.sheets if s.name == "LOOK TOTAL"), None)
    if lt is None:
        pytest.skip(
            "this sample copy of the map has no 'LOOK TOTAL' sheet "
            "(only FW26 and an empty '-->' sheet); the block layout is "
            "covered by test_look_total_block_layout_synthetic"
        )
    assert lt.highlighted_cells == 60
    assert len(lt.bases) == 39


def _yellow():
    return PatternFill(fill_type="solid", start_color=Color(theme=7, tint=0.8))


def test_look_total_block_layout_synthetic(tmp_path):
    """Replicates the documented LOOK TOTAL geometry: 8-column look blocks
    starting at A, J, S, AB, ... with SKUs under 'MMC' (block-relative col 3),
    data from row 8; 60 highlighted cells over 46 unique SKUs / 39 bases,
    plus gray-filled cells that must be flagged, not extracted."""
    bases = [f"644T{i:03d}A99" for i in range(39)]
    skus = [b + "X5800" for b in bases] + [b + "X0900" for b in bases[:7]]
    assert len(set(skus)) == 46
    highlight_cells = skus + skus[:14]  # 60 highlighted cells
    plain_cells = [b + "X5800" for b in bases[:20]]

    wb = Workbook()
    ws = wb.active
    ws.title = "LOOK TOTAL"
    block_cols = [3, 12, 21, 30, 39, 48, 57, 66]  # C, L, U, AD, AM, AV, BE, BN
    for col in block_cols:
        ws.cell(row=7, column=col, value="MMC")

    slots = [(row, col) for row in range(8, 40) for col in block_cols]
    it = iter(slots)
    for sku in highlight_cells:
        row, col = next(it)
        c = ws.cell(row=row, column=col, value=sku)
        c.fill = _yellow()
    for sku in plain_cells:
        row, col = next(it)
        ws.cell(row=row, column=col, value=sku)
    gray_specs = [(0, -0.15), (0, -0.15), (0, -0.15), (6, 0.8), (6, 0.8)]
    for i, (theme, tint) in enumerate(gray_specs):
        row, col = next(it)
        c = ws.cell(row=row, column=col, value=bases[20 + i] + "X5800")
        c.fill = PatternFill(fill_type="solid",
                             start_color=Color(theme=theme, tint=tint))
    # structural banding (theme 4) must be silently ignored
    row, col = next(it)
    c = ws.cell(row=row, column=col, value=bases[30] + "X5800")
    c.fill = PatternFill(fill_type="solid",
                         start_color=Color(theme=4, tint=0.8))

    path = tmp_path / "look_total.xlsx"
    wb.save(path)

    scan = pipeline.scan_sell_thru_map(path)
    lt = next(s for s in scan.sheets if s.name == "LOOK TOTAL")
    assert lt.highlighted_cells == 60
    assert len(lt.highlighted_skus) == 46
    assert len(lt.bases) == 39
    grays = [note for _, _, note in lt.other_fills]
    assert len(grays) == 5
    assert all(note.startswith("theme") for note in grays)
    # the step-trace records where each highlight was found and its base
    assert len(lt.highlighted_cell_details) == 60
    assert lt.highlighted_cell_details[0] == ("C8", "644T000A99X5800", "644T000A99")


def test_rgb_yellow_fallback(tmp_path):
    """A future map that uses literal RGB yellows must still be detected."""
    wb = Workbook()
    ws = wb.active
    c = ws.cell(row=1, column=1, value="644T000A99X5800")
    c.fill = PatternFill(fill_type="solid", start_color="FFFFF2CC")
    path = tmp_path / "rgb.xlsx"
    wb.save(path)
    scan = pipeline.scan_sell_thru_map(path)
    assert scan.sheets[0].highlighted_cells == 1


# --- 2. base-SKU extraction -------------------------------------------------

def test_base_extraction_double_x():
    assert pipeline.extract_base("652P92X3F74X8090") == "652P92X3F74"
    assert re.sub(r"X[0-9A-Z]{4}$", "", "652P92X3F74X8090") == "652P92X3F74"
    assert pipeline.extract_base("644S83A7A26X5883") == "644S83A7A26"
    assert pipeline.extract_base("  644S83A7A26X5883 ") == "644S83A7A26"
    # underscore-separated color codes (delivery-tracker style)
    assert pipeline.extract_base("641V19A1491_X8300") == "641V19A1491"
    # no recognizable color block -> base is the SKU itself
    assert pipeline.extract_base("KCK554TFS_S03W") == "KCK554TFS_S03W"


# --- 2b. AI-located SKU columns ---------------------------------------------

def _track_workbook(path):
    """Delivery-tracker-style sheet: SKUs under a 'SKU' header in a format
    the strict regex cannot recognize."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Track"
    ws.append(["SKU", "QTY"])
    c = ws.cell(row=2, column=1, value="M0715OUQO_M900_TU")
    c.fill = _yellow()
    ws.cell(row=3, column=1, value="KCV536VCR_S900_T37")
    c = ws.cell(row=4, column=1, value="1222")  # numeric — not a SKU
    c.fill = _yellow()
    ws.cell(row=5, column=1, value="641V19A1491_X8300").fill = _yellow()
    wb.save(path)


def test_ai_columns_extend_scan(tmp_path):
    path = tmp_path / "track.xlsx"
    _track_workbook(path)

    # without AI columns the strict regex finds nothing here
    plain = pipeline.scan_sell_thru_map(path).sheets[0]
    assert plain.sku_cells == 0 and plain.highlighted_cells == 0

    ai_columns = {"Track": {"header_row": 1, "columns": {"A": "SKU"}}}
    scan = pipeline.scan_sell_thru_map(path, ai_columns=ai_columns).sheets[0]
    assert scan.ai_sku_cells == 3          # '1222' filtered as implausible
    assert scan.highlighted_cells == 2     # M0715... and 641V19...
    assert ("A2", "M0715OUQO_M900_TU", "M0715OUQO_M900_TU") \
        in scan.ai_highlighted_cell_details
    # underscore color code is stripped for the base
    assert ("A5", "641V19A1491_X8300", "641V19A1491") \
        in scan.ai_highlighted_cell_details
    # the header cell itself is never extracted
    assert all(coord != "A1" for coord, _, _ in scan.ai_highlighted_cell_details)


def test_pipeline_with_fake_locator(tmp_path):
    map_path = tmp_path / "map.xlsx"
    _track_workbook(map_path)
    qwb = Workbook()
    qws = qwb.active
    qws.title = "query"
    qws.append(pipeline.QUERY_HEADERS)
    qws.append(["3617000000001", "641V19A1491 - X8300 - T36", "d", "dep",
                "s", None, None, None, 1, 10, 20, "t", "p"])
    query_path = tmp_path / "q.xlsx"
    qwb.save(query_path)

    detection = [{"sheet": "Track", "header_row": 1,
                  "sku_columns": [{"column": "A", "header": "SKU",
                                   "reason": "data cells hold item codes"}]}]
    report = pipeline.run_pipeline(
        map_path, query_path,
        Path(__file__).parent / "assets" / "ProductsListTemplate.xlsx",
        tmp_path / "out.xlsx",
        locator=lambda p: detection,
    )
    assert report.ai_enabled and report.ai_detection == detection
    assert report.ai_note is None
    # 641V19A1491 matched the query; the underscore code did not
    assert report.matched_counts == {"641V19A1491": 1}
    assert "M0715OUQO_M900_TU" in report.unmatched
    assert report.rows_written == 1


def test_pipeline_survives_locator_failure(tmp_path):
    map_path = tmp_path / "map.xlsx"
    _track_workbook(map_path)
    qwb = Workbook()
    qws = qwb.active
    qws.title = "query"
    qws.append(pipeline.QUERY_HEADERS)
    query_path = tmp_path / "q.xlsx"
    qwb.save(query_path)

    def broken(_path):
        raise RuntimeError("api unreachable")

    report = pipeline.run_pipeline(
        map_path, query_path,
        Path(__file__).parent / "assets" / "ProductsListTemplate.xlsx",
        tmp_path / "out.xlsx",
        locator=broken,
    )
    assert report.ai_enabled
    assert report.ai_note == "api unreachable"
    assert report.rows_written == 0  # run completed regardless


def test_locator_previews_and_normalization(tmp_path):
    path = tmp_path / "track.xlsx"
    _track_workbook(path)
    previews = sku_locator.build_previews(path)
    assert previews[0]["sheet"] == "Track"
    assert previews[0]["rows"]["1"]["A"] == "SKU"
    assert previews[0]["rows"]["2"]["A"] == "M0715OUQO_M900_TU"

    raw = [
        {"sheet": "Track", "header_row": 1,
         "sku_columns": [{"column": "a", "header": "SKU", "reason": "ok"},
                         {"column": "5", "header": None, "reason": "bad"}]},
        {"sheet": "Ghost", "header_row": 1, "sku_columns": []},
    ]
    normalized = sku_locator._normalize(raw, ["Track"])
    assert normalized == [{"sheet": "Track", "header_row": 1,
                           "sku_columns": [{"column": "A", "header": "SKU",
                                            "reason": "ok"}]}]


def _verification_app(tmp_path, monkeypatch, detection):
    """App with a faked AI locator: configured, returns `detection`."""
    monkeypatch.setattr(sku_locator, "is_configured", lambda: True)
    monkeypatch.setattr(sku_locator, "locate_from_previews",
                        lambda previews: detection)
    app = create_app(data_dir=tmp_path / "data", password="")
    map_path = tmp_path / "track.xlsx"
    _track_workbook(map_path)
    qwb = Workbook()
    qws = qwb.active
    qws.title = "query"
    qws.append(pipeline.QUERY_HEADERS)
    qws.append(["3617000000001", "641V19A1491 - X8300 - T36", "d", "dep",
                "s", None, None, None, 1, 10, 20, "t", "p"])
    query_path = tmp_path / "q.xlsx"
    qwb.save(query_path)
    return app.test_client(), map_path, query_path


def test_human_verification_flow(tmp_path, monkeypatch):
    detection = [{"sheet": "Track", "header_row": 1,
                  "sku_columns": [{"column": "A", "header": "SKU",
                                   "reason": "data cells hold item codes"}]}]
    client, map_path, query_path = _verification_app(
        tmp_path, monkeypatch, detection)

    # phase 1: upload pauses on the verification page instead of processing
    with map_path.open("rb") as m, query_path.open("rb") as q:
        resp = client.post(
            "/process",
            data={"map_file": (m, "track.xlsx"), "query_file": (q, "q.xlsx")},
            content_type="multipart/form-data",
        )
    assert resp.status_code == 200
    html = resp.data.decode()
    assert "Verify the AI-detected SKU columns" in html
    assert "M0715OUQO_M900_TU" in html          # sample values shown
    assert 'name="col" value="0:0" checked' in html.replace("\n", " ").replace("  ", " ") \
        or re.search(r'name="col"\s+value="0:0"\s+checked', html)
    action = re.search(r'action="/process/(\d{8}_\d{6}_[0-9a-f]{6})"', html)
    assert action, "confirmation form must post to /process/<pending_id>"
    pending_id = action.group(1)

    # phase 2: confirm the column -> full run -> report marks it confirmed
    resp = client.post(f"/process/{pending_id}", data={"col": "0:0"},
                       follow_redirects=True)
    assert resp.status_code == 200
    html = resp.data.decode()
    assert "You reviewed and confirmed these columns" in html
    assert "641V19A1491" in html                # matched via the AI column
    assert "M0715OUQO_M900_TU" in html          # unmatched, still traced

    # the pending upload is consumed: files gone, replay 404s
    assert not list((tmp_path / "data" / "pending").glob("*"))
    assert client.post(f"/process/{pending_id}", data={}).status_code == 404


def test_verification_deselect_all_falls_back_to_standard_scan(
        tmp_path, monkeypatch):
    detection = [{"sheet": "Track", "header_row": 1,
                  "sku_columns": [{"column": "A", "header": "SKU",
                                   "reason": "codes"}]}]
    client, map_path, query_path = _verification_app(
        tmp_path, monkeypatch, detection)
    with map_path.open("rb") as m, query_path.open("rb") as q:
        resp = client.post(
            "/process",
            data={"map_file": (m, "track.xlsx"), "query_file": (q, "q.xlsx")},
            content_type="multipart/form-data",
        )
    pending_id = re.search(r'action="/process/([0-9_a-f]+)"',
                           resp.data.decode()).group(1)

    # submit with every checkbox unticked
    resp = client.post(f"/process/{pending_id}", data={},
                       follow_redirects=True)
    assert resp.status_code == 200
    html = resp.data.decode()
    assert "You unticked every detected column" in html
    # none of the non-standard SKUs were extracted
    assert "M0715OUQO_M900_TU" not in html


def test_verification_skipped_when_nothing_detected(tmp_path, monkeypatch):
    client, map_path, query_path = _verification_app(
        tmp_path, monkeypatch, detection=[])
    with map_path.open("rb") as m, query_path.open("rb") as q:
        resp = client.post(
            "/process",
            data={"map_file": (m, "track.xlsx"), "query_file": (q, "q.xlsx")},
            content_type="multipart/form-data",
            follow_redirects=True,
        )
    # no confirmation page: straight to the report
    assert resp.status_code == 200
    html = resp.data.decode()
    assert "Verify the AI-detected" not in html
    assert "found no SKU columns" in html


@pytest.mark.skipif(
    not sku_locator.is_configured()
    or not (SAMPLES / "delivery_track.xlsx").exists(),
    reason="needs ANTHROPIC_API_KEY and samples/delivery_track.xlsx",
)
def test_locator_live_on_delivery_track():
    detection = sku_locator.locate(SAMPLES / "delivery_track.xlsx")
    by_sheet = {d["sheet"]: [c["column"] for c in d["sku_columns"]]
                for d in detection}
    assert "E" in by_sheet.get("All Cat", [])   # 'TS SKU' column
    assert "A" in by_sheet.get("Sheet2", [])    # 'SKU' column


# --- 3. query matching ------------------------------------------------------

@needs_samples
def test_every_fw26_base_matches_query(map_scan, fw26_matched):
    fw26 = next(s for s in map_scan.sheets if s.name == "FW26")
    assert set(fw26_matched) == fw26.bases  # every base has >= 1 row
    assert all(len(rows) >= 1 for rows in fw26_matched.values())
    total = sum(len(rows) for rows in fw26_matched.values())
    # "~283" in the spec; the exact figure shifts by a row or two between
    # query export versions, so assert a tight band, not one copy's value.
    assert 275 <= total <= 295


def test_match_is_exact_first_segment(tmp_path):
    """'644S16B7E7' must not prefix-match '644S16B7E72 - ...'."""
    wb = Workbook()
    ws = wb.active
    ws.title = "query"
    ws.append(pipeline.QUERY_HEADERS)
    ws.append(["1", "644S16B7E72 - X4804 - T36", "d", "dep", "s",
               None, None, None, 1, 10, 20, "t", "p"])
    path = tmp_path / "q.xlsx"
    wb.save(path)
    assert pipeline.match_query(path, {"644S16B7E7"}) == {}
    assert list(pipeline.match_query(path, {"644S16B7E72"})) == ["644S16B7E72"]


# --- 4. output workbook -----------------------------------------------------

@needs_samples
def test_output_workbook_structure(tmp_path, fw26_matched):
    out = tmp_path / "out.xlsx"
    rows_written = pipeline.build_output(
        TEMPLATE, out, fw26_matched,
        unmatched_bases=["FAKEBASE001"],
        base_sheets={"FAKEBASE001": ["LOOK TOTAL"]},
    )
    # every matched query row must land in the workbook, no more, no fewer
    assert rows_written == sum(len(rows) for rows in fw26_matched.values())

    wb = load_workbook(out)
    ws = wb["Sheet1"]
    twb = load_workbook(TEMPLATE)
    assert [c.value for c in ws[1]] == [c.value for c in twb["Sheet1"][1]]
    twb.close()

    last = rows_written + 1
    for n in (2, 3, last):
        assert ws.cell(row=n, column=8).value == \
            f'=IF(AND(F{n}<>"", G{n}<>""), F{n}*G{n}, "")'
        assert ws.cell(row=n, column=10).value == \
            f'=IF(AND(F{n}<>"", I{n}<>""), F{n}*I{n}, "")'
        for col in (11, 12, 13):  # K, L, M stay blank
            assert ws.cell(row=n, column=col).value is None
        barcode = ws.cell(row=n, column=1)
        assert isinstance(barcode.value, str)
        assert barcode.number_format == "@"

    # sorted by base, then color, then size
    skus = [ws.cell(row=n, column=2).value for n in range(2, last + 1)]
    keys = [tuple(s.split(" - ")) for s in skus]
    assert keys == sorted(keys)

    um = wb["Unmatched"]
    assert um["A1"].value == "Base SKU"
    assert um["A2"].value == "FAKEBASE001"
    assert um["B2"].value == "LOOK TOTAL"
    wb.close()


# --- 5. Flask end-to-end ----------------------------------------------------

@needs_samples
def test_end_to_end_upload_and_download(tmp_path):
    app = create_app(data_dir=tmp_path, password="")
    client = app.test_client()

    page = client.get("/")
    assert page.status_code == 200
    assert b"No stock query export stored yet" in page.data

    with MAP.open("rb") as m, QUERY.open("rb") as q:
        resp = client.post(
            "/process",
            data={"map_file": (m, "map.xlsx"), "query_file": (q, "query.xlsx")},
            content_type="multipart/form-data",
            follow_redirects=True,  # POST redirects to GET /report/<run_id>
        )
    assert resp.status_code == 200
    html = resp.data.decode()
    match = re.search(r"/download/(\d{8}_\d{6}_[0-9a-f]{6})", html)
    assert match, "report page must contain a download link"
    rows = re.search(r"wrote <strong>(\d+)</strong> rows", html)
    assert rows, "report page must state the number of rows written"
    rows_written = int(rows.group(1))
    assert 275 <= rows_written <= 295  # "~283", varies with query version

    # the report must trace internal steps: highlighted cells found, base
    # extraction, and the query data pulled per base
    assert "Processing steps" in html
    # without an API key configured, the AI section says it is off
    assert "AI column detection" in html
    cell_trace = re.search(
        r"<td>([A-Z]{1,3}\d+)</td><td>([0-9A-Z]{8,})</td><td>([0-9A-Z]{8,})</td>",
        html.replace("\n", ""))
    assert cell_trace, "step 1 must list highlighted cells with SKU and base"
    detail = re.search(r"<summary>([0-9A-Z]{8,}) — (\d+) query", html)
    assert detail, "step 3 must show per-base query detail"
    assert detail.group(1) in html  # base appears in the matched table too

    dl = client.get(f"/download/{match.group(1)}")
    assert dl.status_code == 200
    assert "ProductsList_" in dl.headers["Content-Disposition"]
    wb = load_workbook(io.BytesIO(dl.data))
    assert "Sheet1" in wb.sheetnames and "Unmatched" in wb.sheetnames
    # the downloaded workbook must contain exactly the reported rows
    assert wb["Sheet1"].max_row == rows_written + 1
    wb.close()

    # the query is now stored: a map-only run must also succeed, and get
    # its own run_id even within the same second (concurrent-run safety)
    page = client.get("/")
    assert b"Stored stock query" in page.data
    with MAP.open("rb") as m:
        resp = client.post("/process", data={"map_file": (m, "map.xlsx")},
                           content_type="multipart/form-data",
                           follow_redirects=True)
    assert resp.status_code == 200
    second = re.search(r"/download/(\d{8}_\d{6}_[0-9a-f]{6})",
                       resp.data.decode())
    assert second and second.group(1) != match.group(1)

    # the report is persistent (refresh-safe) and language-switchable:
    # the same run re-renders in Chinese with the same download link
    zh = client.get(f"/report/{match.group(1)}?lang=zh")
    assert zh.status_code == 200
    zh_html = zh.data.decode()
    assert "处理步骤" in zh_html            # "Processing steps"
    assert "标黄基础款号" in zh_html         # cohesive summary wording
    assert f"/download/{match.group(1)}" in zh_html
    # and back to English via the toggle
    en = client.get(f"/report/{match.group(1)}?lang=en")
    assert "Processing steps" in en.data.decode()


def test_query_upload_with_wrong_headers_is_rejected(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "query"
    ws.append(["Totally", "Wrong", "Headers"])
    bad = tmp_path / "bad_query.xlsx"
    wb.save(bad)

    ok, message, _ = pipeline.validate_query_file(bad)
    assert not ok and "header" in message

    app = create_app(data_dir=tmp_path / "data", password="")
    client = app.test_client()
    with bad.open("rb") as b, bad.open("rb") as m:
        resp = client.post(
            "/process",
            data={"map_file": (m, "map.xlsx"), "query_file": (b, "q.xlsx")},
            content_type="multipart/form-data",
        )
    assert resp.status_code == 400
    assert b"rejected" in resp.data


def test_malformed_map_upload_is_a_friendly_error(tmp_path):
    app = create_app(data_dir=tmp_path, password="")
    client = app.test_client()
    # store a minimal valid query first
    wb = Workbook()
    ws = wb.active
    ws.title = "query"
    ws.append(pipeline.QUERY_HEADERS)
    good_q = tmp_path / "q.xlsx"
    wb.save(good_q)
    with good_q.open("rb") as q:
        resp = client.post(
            "/process",
            data={"map_file": (io.BytesIO(b"this is not a zip"), "map.xlsx"),
                  "query_file": (q, "q.xlsx")},
            content_type="multipart/form-data",
        )
    assert resp.status_code == 400
    assert b"could not be processed" in resp.data


def test_basic_auth_required_when_password_set(tmp_path):
    app = create_app(data_dir=tmp_path, password="s3cret")
    client = app.test_client()
    assert client.get("/").status_code == 401
    token = base64.b64encode(b"anyuser:s3cret").decode()
    assert client.get("/", headers={"Authorization": f"Basic {token}"}) \
        .status_code == 200
    wrong = base64.b64encode(b"anyuser:nope").decode()
    assert client.get("/", headers={"Authorization": f"Basic {wrong}"}) \
        .status_code == 401


def test_language_toggle_and_cookie(tmp_path):
    app = create_app(data_dir=tmp_path, password="")
    client = app.test_client()

    # default is English, with a top-left switch offering Chinese
    page = client.get("/")
    assert "No stock query export stored yet" in page.data.decode()
    assert 'class="lang"' in page.data.decode()
    assert "中文" in page.data.decode()

    # ?lang=zh renders Chinese and persists the choice in a cookie
    page = client.get("/?lang=zh")
    html = page.data.decode()
    assert "还没有保存库存查询表" in html
    assert "开始处理" in html          # Process button
    assert "English" in html           # switch now offers English
    assert "lang=zh" in page.headers.get("Set-Cookie", "")

    # subsequent plain requests stay in Chinese via the cookie
    page = client.get("/")
    assert "还没有保存库存查询表" in page.data.decode()

    # switching back works
    page = client.get("/?lang=en")
    assert "No stock query export stored yet" in page.data.decode()

    # error pages are translated too
    resp = client.post("/process?lang=zh", data={},
                       content_type="multipart/form-data")
    assert resp.status_code == 400
    assert "请先选择 Sell-Thru Map" in resp.data.decode()


def test_help_page_in_both_languages(tmp_path):
    app = create_app(data_dir=tmp_path, password="")
    client = app.test_client()

    # linked from the front page
    assert b"/help" in client.get("/").data

    en = client.get("/help").data.decode()
    assert en.count("<h2>") >= 5
    assert "What this tool does" in en
    assert "Your first run, step by step" in en
    assert "filter" in en  # covers the hidden-rows gotcha

    zh = client.get("/help?lang=zh").data.decode()
    assert "这个工具是做什么的" in zh
    assert "第一次使用" in zh
    assert "筛选" in zh

    # help is behind auth like everything else
    locked = create_app(data_dir=tmp_path / "locked", password="pw")
    assert locked.test_client().get("/help").status_code == 401


def test_basic_auth_handles_non_ascii_passwords(tmp_path):
    """compare_digest on str raises TypeError for non-ASCII — a login typo
    must yield 401 and a non-ASCII APP_PASSWORD must still work, never 500."""
    app = create_app(data_dir=tmp_path, password="Zürich2026")
    client = app.test_client()
    right = base64.b64encode("u:Zürich2026".encode()).decode()
    assert client.get("/", headers={"Authorization": f"Basic {right}"}) \
        .status_code == 200
    wrong = base64.b64encode("u:pässword".encode()).decode()
    assert client.get("/", headers={"Authorization": f"Basic {wrong}"}) \
        .status_code == 401


def test_corrupt_xml_map_gets_friendly_error(tmp_path):
    """A valid zip with truncated sheet XML raises ElementTree.ParseError —
    it must land on the friendly error page, not a bare 500."""
    import zipfile as zf

    src = Path(__file__).parent / "assets" / "ProductsListTemplate.xlsx"
    corrupt = io.BytesIO()
    with zf.ZipFile(src) as zin, zf.ZipFile(corrupt, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("xl/worksheets/"):
                data = data[: len(data) // 2]
            zout.writestr(item, data)
    corrupt.seek(0)

    app = create_app(data_dir=tmp_path, password="")
    client = app.test_client()
    wb = Workbook()
    ws = wb.active
    ws.title = "query"
    ws.append(pipeline.QUERY_HEADERS)
    good_q = tmp_path / "q.xlsx"
    wb.save(good_q)
    with good_q.open("rb") as q:
        resp = client.post(
            "/process",
            data={"map_file": (corrupt, "map.xlsx"), "query_file": (q, "q.xlsx")},
            content_type="multipart/form-data",
        )
    assert resp.status_code == 400
    assert b"could not be processed" in resp.data


def test_indexed_color_fills_are_flagged_not_dropped(tmp_path):
    """Legacy indexed-palette fills are outside the highlight definition but
    must surface in 'other fills detected' instead of vanishing silently."""
    wb = Workbook()
    ws = wb.active
    c = ws.cell(row=1, column=1, value="644T000A99X5800")
    c.fill = PatternFill(fill_type="solid", start_color=Color(indexed=6))
    path = tmp_path / "indexed.xlsx"
    wb.save(path)
    scan = pipeline.scan_sell_thru_map(path)
    sheet = scan.sheets[0]
    assert sheet.highlighted_cells == 0
    assert len(sheet.other_fills) == 1
    assert "indexed" in sheet.other_fills[0][2]
