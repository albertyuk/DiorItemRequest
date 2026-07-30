"""UI strings for the two supported languages.

The Chinese is written for the merch workflow this tool serves, using the
sell-thru map's own vocabulary (标黄 = yellow-highlighted, 补货 = reorder,
基础款号 = base SKU) rather than literal word-by-word translation. File
names and Excel artifacts (Sell-Thru Map, ProductsList, Unmatched, SKU,
TotalCost) stay in English in both languages — they are proper nouns the
team sees in the actual files.

Technical error details (exception text) are appended untranslated.
"""

SUPPORTED_LANGS = ("en", "zh")


def translate(lang: str, key: str, **fmt) -> str:
    """Shared lookup used by both app.py and auth_routes.py."""
    s = STRINGS.get(lang, STRINGS["en"]).get(key) or STRINGS["en"][key]
    return s.format(**fmt) if fmt else s

STRINGS = {
    "en": {
        "site_title": "Dior item request — highlighted-SKU extractor",
        "page_upload": "upload",
        "page_report": "run report",
        "page_error": "error",
        "page_help": "tutorial",
        "help_link": "New here? Read the tutorial →",
        "lang_switch": "中文",

        # index
        "intro": ("Extracts the reorder SKUs from a sell-thru map — rows "
                  "marked <code>Y</code> in a <code>PickUp</code> column "
                  "when a sheet has one, yellow-highlighted cells otherwise "
                  "— looks up all color/size variants in the stored stock "
                  "query export, and produces a filled ProductsList "
                  "workbook. Nothing is written until you approve the "
                  "result on the review page."),
        "stored_query": ("Stored stock query: <strong>{name}</strong> "
                         "&mdash; {rows} rows, uploaded {time}."),
        "no_query_warning": ("<strong>No stock query export stored yet.</strong> "
                             "Upload one below — it is kept for future runs "
                             "until replaced."),
        "map_label": ("Sell-thru map (.xlsx) — the workbook with the "
                      "yellow-highlighted styles; upload it every run"),
        "query_label": ("Stock query export (.xlsx) — optional: leave empty "
                        "to use the stored one, or choose a file to replace "
                        "it"),
        "process_btn": "Process",

        # report
        "new_run": "&larr; new run",
        "processed_summary": ("Processed <strong>{map}</strong> — wrote "
                              "<strong>{rows}</strong> rows for "
                              "<strong>{matched}</strong> of "
                              "<strong>{total}</strong> highlighted base SKUs."),
        "download_btn": "Download",
        "steps_heading": "Processing steps",
        "step1_title": "Step 1 — scan every sheet for highlighted SKUs",
        "step1_desc": ("Every cell is checked for a SKU-shaped value; a cell "
                       "counts as highlighted when it carries the yellow "
                       "\"reorder\" fill."),
        "sheet_details_summary": "{sheet} — the {n} highlighted cells found",
        "th_sheet": "Sheet", "th_sku_cells": "SKU cells",
        "th_highlighted_cells": "Highlighted cells",
        "th_unique_skus": "Unique SKUs", "th_unique_bases": "Unique bases",
        "th_cell": "Cell", "th_sku_as_highlighted": "SKU (as highlighted)",
        "th_base_extracted": "Base extracted",
        "ai_heading": "AI column detection",
        "ai_desc": ("Claude reads each sheet's header area and points at the "
                    "column(s) holding SKU codes, so highlighted SKUs in "
                    "non-standard formats (accessory MMCs, underscore codes) "
                    "are extracted too — not only the standard pattern."),
        "ai_th_header_row": "Header row", "ai_th_column": "SKU column",
        "ai_th_header": "Header", "ai_th_reason": "Why",
        "ai_found_cells": ("{sheet} — {n} highlighted cells found only via "
                           "AI-located columns"),
        "ai_none": ("The AI reviewed the sheet headers and found no SKU "
                    "columns beyond what the standard scan already covers."),
        "ai_confirmed_note": ("&#10003; You reviewed and confirmed these "
                              "columns before processing."),
        "ai_all_deselected": ("You unticked every detected column, so only "
                              "the standard scan was used."),
        "th_samples": "Sample values",
        "confirm_cancel": "Cancel and start over",
        "nav_home": "Home", "nav_help": "Tutorial",
        "step_upload": "Upload", "step_review": "Review",
        "step_report": "Report",
        "page_upload_heading": "New run",
        "page_review": "review",
        "review_heading": "Review before building",
        "review_intro": ("Everything extracted from <strong>{map}</strong> "
                         "is listed below. Ticked SKUs go into the "
                         "ProductsList; untick the ones that should stay "
                         "out, then press the build button — it shows "
                         "exactly how many rows will be written, and no file "
                         "is created until you press it."),
        "review_ai_intro": ("Claude located these SKU columns by reading the "
                            "sheet headers. Check the sample values: do they "
                            "look like real SKU codes? SKUs from these "
                            "columns carry an “AI” badge in "
                            "the list below — untick any that look wrong."),
        "review_sku_heading": "Extracted SKUs",
        "th_colorways": "Highlighted colorway(s)",
        "th_stock_rows": "Stock rows",
        "badge_no_stock": "no stock", "badge_ai": "AI",
        "sel_all": "Select all", "sel_none": "Select none",
        "sel_matched": "Only with stock",
        "review_counter": "{n} of {m} SKUs selected",
        "btn_build": "Build ProductsList ({rows} rows)",
        "recent_heading": "Recent runs",
        "recent_none": "No runs yet — the last 10 stay downloadable here.",
        "recent_map": "Map", "recent_when": "When",
        "recent_report": "report", "recent_download": "download",
        "th_rows_written": "Rows",
        "query_used": ("Stock query used: <strong>{name}</strong> — {rows} "
                       "rows, uploaded {time}."),
        "th_source": "Source", "src_ai": "AI", "src_memory": "remembered",
        "badge_pickup": "PickUp",
        "th_value": "Value",
        "th_image": "Image",
        "report_images_note": ("{n} product image(s) were copied from the "
                               "map into column <code>N</code>, at the "
                               "right edge of the sheet — one per base "
                               "SKU, on the first row of its block."),
        "review_pickup_note": ("Sheet {sheet}: selection comes from its "
                               "PickUp column ({col}) — {n} marked row(s); "
                               "highlights on this sheet do not select."),
        "review_pickup_warn": ("{sheet}: {bad} unrecognized PickUp "
                               "value(s) were ignored, {nosku} marked "
                               "row(s) held no SKU, and {unmarked} "
                               "highlighted SKU(s) are not marked and are "
                               "therefore NOT selected — details go on the "
                               "run report."),
        "pickup_heading": "PickUp column (explicit selection)",
        "pickup_desc": ("When a sheet has a column headed "
                        "<code>PickUp</code>, that column is the sheet's "
                        "selector: rows marked <code>Y</code> (also "
                        "accepted: yes / x / 1 / pickup / 是) are "
                        "extracted, and highlights there are reported for "
                        "reference only. Sheets without the column keep "
                        "the yellow-highlight behavior."),
        "pickup_sheet_line": ("{sheet}: PickUp column <strong>{col}</strong> "
                              "(header on row {row}) — <strong>{n}</strong> "
                              "marked row(s)."),
        "pickup_details_summary": "{sheet} — the {n} SKUs picked up",
        "pickup_unrecognized_note": ("{n} value(s) in the PickUp column "
                                     "were not recognized and did NOT "
                                     "select their row (only Y / yes / x / "
                                     "1 / pickup / 是 count):"),
        "pickup_no_sku_note": ("{n} marked row(s) contained no "
                               "recognizable SKU:"),
        "pickup_unmarked_note": ("{n} highlighted SKU(s) on this sheet are "
                                 "NOT marked in the PickUp column and were "
                                 "NOT selected — mark them with Y and "
                                 "re-run if they should be ordered:"),
        "pickup_none": ("No PickUp column on any sheet — SKUs were "
                        "selected by yellow highlight (the standard "
                        "behavior)."),
        "review_memory_note": ("Columns marked “remembered” were approved by "
                               "a reviewer on an earlier run, so they needed "
                               "no AI call this time. To make the tool "
                               "forget a sheet's remembered columns, untick "
                               "every SKU of that sheet before building."),
        "by_uploaded": "Uploaded by {name}",
        "by_built": "Reviewed & built by {name}",
        "th_user": "By",
        "excluded_heading": "Excluded at review ({n})",
        "excluded_note": ("These highlighted SKUs were unticked on the "
                          "review page and left out of the workbook."),
        "page_progress": "processing",
        "progress_heading": "Processing — this page updates itself",
        "stage_start": "Starting…",
        "stage_scan": "Scanning the map for highlighted SKUs",
        "stage_match": "Matching bases against the stock query",
        "stage_write": "Writing the ProductsList workbook",
        "progress_lost": ("The run was interrupted (probably a server "
                          "restart). Nothing was saved — go back and upload "
                          "again."),
        "upload_uploading": "Uploading…",
        "upload_analyzing": "Upload complete — reading the sheet headers…",
        "ai_failed": ("AI column detection was unavailable this run: "
                      "{detail}. The standard scan still ran in full, and "
                      "any “remembered” columns were still used."),
        "ai_off": ("AI column detection is off — set the ANTHROPIC_API_KEY "
                   "secret to enable it. The standard scan ran in full."),
        "step2_title": "Step 2 — extract base SKUs",
        "step2_desc": ("The trailing color block (<code>X</code> + 4 "
                       "characters) is stripped from each highlighted SKU; "
                       "duplicates collapse to <strong>{n}</strong> unique "
                       "bases."),
        "base_skus_summary": "Base SKUs and the highlighted colorways behind them",
        "th_base": "Base SKU", "th_from_skus": "From highlighted SKU(s)",
        "th_sheets": "Sheet(s)",
        "step3_title": "Step 3 — look up every color/size variant in the stock query",
        "step3_desc": ("A query row matches when the first "
                       "<code>&nbsp;-&nbsp;</code>-separated segment of its "
                       "<code>Sku</code> equals the base exactly. Expand a "
                       "base to see the data pulled for it."),
        "base_query_rows": "{base} — {n} query rows",
        "th_barcode": "Barcode", "th_qsku": "Sku", "th_division": "Division",
        "th_department": "Department", "th_season": "Season", "th_qty": "Qty",
        "th_unit_cost": "Unit cost", "th_unit_retail": "Unit retail",
        "more_rows": ("&hellip; and {n} more rows (all are in the output "
                      "workbook)."),
        "no_matches": "No bases matched the stored query.",
        "step4_title": "Step 4 — build the ProductsList workbook",
        "step4_desc": ("One row per matched query row (<strong>{rows}</strong> "
                       "total), sorted by base, color, size. "
                       "<code>TotalCost</code> and <code>TotalSales</code> are "
                       "written as formulas; barcodes as text; Gift Recipient "
                       "/ Organization / Position left blank. Unmatched bases "
                       "go on the <code>Unmatched</code> sheet."),
        "matched_heading": "Matched bases ({n})",
        "th_query_rows": "Query rows", "th_highlighted_on": "Highlighted on",
        "none_label": "None.",
        "unmatched_heading": "Unmatched bases ({n})",
        "unmatched_note": ("These highlighted bases had <strong>no rows</strong> "
                           "in the stored stock query. They are also listed on "
                           "the <code>Unmatched</code> sheet of the output "
                           "workbook."),
        "none_unmatched": ("None — every highlighted base matched at least one "
                           "query row."),
        "other_fills_heading": "Other fills detected ({n})",
        "other_fills_note": ("These SKU cells carry a solid fill that is "
                             "neither the yellow highlight nor the structural "
                             "banding (gray notes, etc.). They were "
                             "<strong>not</strong> extracted — please check "
                             "them by hand."),
        "th_fill": "Fill",

        # auth: login / setup / forgot
        "auth_login_title": "Log in",
        "auth_username": "Username",
        "auth_password": "Password",
        "auth_login_btn": "Log in",
        "auth_login_failed": "Wrong username or password.",
        "auth_throttled": ("Too many attempts — wait {n} seconds, then try "
                           "again."),
        "auth_login_hint": ("Accounts are created by an admin on the Team "
                            "page. No account yet? Ask an admin to invite "
                            "you."),
        "auth_login_setup_link": ("Create the admin account (needs the setup "
                                  "code)"),
        "auth_login_forgot_link": "Forgot your password?",
        "auth_setup_title": "Admin setup",
        "auth_setup_hint": ("This page creates — or recovers — the admin "
                            "account, using the server's setup code. An "
                            "existing account is reset to the password you "
                            "enter here; fields left blank keep their "
                            "current values."),
        "auth_setup_code": "Setup code",
        "auth_setup_display": "Display name (optional)",
        "auth_setup_email": "Email (optional — enables password reset)",
        "auth_setup_btn": "Create admin account",
        "auth_setup_wrong_code": "Wrong setup code.",
        "auth_setup_disabled": ("Setup is unavailable — no setup code is "
                                "configured on the server."),
        "auth_bad_username": ("Usernames are 2–32 characters — lowercase "
                              "letters, digits, dots, dashes or "
                              "underscores, starting with a letter or "
                              "digit."),
        "auth_bad_password": "Passwords need at least 8 characters.",
        "auth_bad_email": "That does not look like a valid email address.",
        "auth_email_taken": ("That email address already belongs to another "
                             "account."),
        "auth_username_taken": "That username is already taken.",
        "auth_forgot_title": "Reset your password",
        "auth_forgot_intro": ("Enter your account's email address. If it is "
                              "confirmed, a reset link will be sent to it."),
        "auth_forgot_email": "Email address",
        "auth_forgot_btn": "Send reset link",
        "auth_forgot_sent": ("If that address belongs to an account, a "
                             "reset link is on its way."),
        "auth_forgot_disabled": ("Email sending is not configured — ask an "
                                 "admin to set a new password for you on "
                                 "the Team page."),
        "auth_pw_invite_title": "Choose your password",
        "auth_pw_reset_title": "Choose a new password",
        "auth_pw_new": "New password (at least 8 characters)",
        "auth_pw_btn": "Save password",
        "auth_token_dead_title": "This link is no longer valid",
        "auth_token_dead": ("The link has expired or was already used. Ask "
                            "for a fresh one."),
        "auth_verified_title": "Email confirmed",
        "auth_verified_msg": ("Your email address is confirmed — password "
                              "reset now works for this account."),
        "auth_go_login": "Go to login",
        "err_session_expired": "Your session expired — log in again.",
        "err_auth_unconfigured": "Authentication is not configured.",
        "nav_team": "Team",
        "nav_logout": "Log out",

        # auth: team page
        "team_heading": "Team",
        "team_th_username": "Username",
        "team_th_display": "Display name",
        "team_th_email": "Email",
        "team_th_status": "Status",
        "team_th_role": "Role",
        "team_th_actions": "Actions",
        "team_status_pending": "Invite pending",
        "team_status_verified": "Verified",
        "team_status_unverified": "Unverified",
        "team_role_admin": "Admin",
        "team_role_member": "Member",
        "team_resend_btn": "Resend invite",
        "team_delete_btn": "Delete",
        "team_add_heading": "Add a coworker",
        "team_add_note": ("With an email address, an invite link is sent "
                          "and the coworker picks their own password (leave "
                          "the initial password empty). Without one, set an "
                          "initial password and share it privately."),
        "team_add_note_manual": ("Email sending is not configured, so set "
                                 "an initial password and share it "
                                 "privately."),
        "team_initial_pw_label": "Initial password (when not sending an invite)",
        "team_is_admin_label": "Admin account",
        "team_add_btn": "Add",
        "team_email_heading": "Email address",
        "team_email_label": "Email address (leave empty to remove)",
        "team_target_label": "Account (username — admins may enter anyone's)",
        "team_verify_btn": "Re-send confirmation",
        "team_pw_heading": "Change password",
        "team_pw_note": ("Changing a password signs that account out "
                         "everywhere — including here, if it is your own."),
        "team_pw_label": "New password (at least 8 characters)",
        "team_save_btn": "Save",
        "msg_user_added": "Account created.",
        "msg_user_deleted": "Account deleted.",
        "msg_invite_sent": "Invite sent.",
        "msg_invite_resent": "Invite re-sent.",
        "msg_invite_send_failed": ("The account was created, but the invite "
                                   "email failed to send — use “Resend "
                                   "invite”, or set an initial password via "
                                   "“Change password”."),
        "msg_send_failed": "The email failed to send — try again.",
        "msg_pw_changed": ("Password changed. Every session for that "
                           "account was signed out."),
        "msg_email_saved": ("Email saved. A confirmation link was sent — "
                            "password reset stays unavailable until the "
                            "address is confirmed."),
        "msg_email_saved_noconfirm": ("Email saved. Email sending is off, "
                                      "so the address cannot be confirmed "
                                      "or used for password reset yet."),
        "msg_email_saved_sendfail": ("Email saved, but the confirmation "
                                     "email failed to send — use “Re-send "
                                     "confirmation” to try again."),
        "msg_email_removed": "Email address removed.",
        "msg_verify_sent": "Confirmation link sent.",
        "err_not_admin": "Only admins can do that.",
        "err_no_such_user": "No such account.",
        "err_delete_self": "You cannot delete your own account.",
        "err_delete_last_admin": "You cannot delete the last admin account.",
        "err_pending_only": "That account is not awaiting an invite.",
        "err_no_email_on_file": "That account has no email address on file.",
        "err_email_off": "Email sending is not configured.",
        "err_password_required": ("An initial password is required when no "
                                  "invite email can be sent."),

        # errors
        "error_heading": "Something went wrong.",
        "back_to_upload": "&larr; back to upload",
        "err_map_required": ("Please choose a sell-thru map file (.xlsx) — it "
                             "is required for every run."),
        "err_query_rejected": ("The query file was rejected and the previously "
                               "stored one (if any) was kept: {detail}"),
        "err_no_query": ("No stock query export is stored yet. Upload one in "
                         "the 'Stock query export' field and try again."),
        "err_map_failed": ("The sell-thru map could not be processed as an "
                           ".xlsx workbook: {detail}"),
        "err_too_large": ("Upload too large: the combined upload must stay "
                          "under 200 MB."),
        "err_not_found": ("Not found — the file may have been pruned (only "
                          "the last 10 outputs are kept)."),
        "err_server": ("Unexpected server error — details are in the "
                       "application log."),
    },

    "zh": {
        "site_title": "Dior 补货申请 — 标黄 SKU 提取工具",
        "page_upload": "上传",
        "page_report": "运行报告",
        "page_error": "错误",
        "page_help": "使用教程",
        "help_link": "第一次使用？点这里看教程 →",
        "lang_switch": "English",

        # index
        "intro": ("自动从 Sell-Thru Map（销售进度表）中提取需补货的 SKU —— "
                  "工作表如有 <code>PickUp</code> 列，则以标 <code>Y</code> "
                  "的行为准；否则按标黄单元格提取。随后在已保存的库存查询表中"
                  "匹配出全部颜色和尺码，生成可直接使用的 ProductsList 工作簿。"
                  "生成前会先进入确认页，经你核对确认后才会写入文件。"),
        "stored_query": ("当前保存的库存查询表：<strong>{name}</strong>，"
                         "共 {rows} 行，上传于 {time}。"),
        "no_query_warning": ("<strong>还没有保存库存查询表。</strong>"
                             "请在下方上传；保存后无需每次重复上传，"
                             "之后上传新文件即可替换。"),
        "map_label": "Sell-Thru Map（.xlsx）— 含标黄款式的工作簿，每次处理都需要上传",
        "query_label": ("库存查询表（.xlsx）— 可选：留空则使用已保存的版本，"
                        "选择新文件即替换"),
        "process_btn": "开始处理",

        # report
        "new_run": "&larr; 返回首页",
        "processed_summary": ("<strong>{map}</strong> 处理完成：共识别出 "
                              "<strong>{total}</strong> 个标黄基础款号，其中 "
                              "<strong>{matched}</strong> 个在库存查询表中匹配到记录，"
                              "已写入 <strong>{rows}</strong> 行。"),
        "download_btn": "下载",
        "steps_heading": "处理步骤",
        "step1_title": "第 1 步：扫描各工作表，找出标黄 SKU",
        "step1_desc": ("工具会逐格检查所有工作表：内容是 SKU、"
                       "底色为黄色补货标记的单元格，都会被识别为标黄。"),
        "sheet_details_summary": "{sheet}：{n} 个标黄单元格（点击查看明细）",
        "th_sheet": "工作表", "th_sku_cells": "SKU 单元格数",
        "th_highlighted_cells": "标黄单元格数",
        "th_unique_skus": "去重后 SKU 数", "th_unique_bases": "去重后基础款号数",
        "th_cell": "单元格", "th_sku_as_highlighted": "标黄的 SKU",
        "th_base_extracted": "提取出的基础款号",
        "ai_heading": "AI 识别 SKU 列",
        "ai_desc": ("Claude 会阅读每个工作表的表头区域，判断哪些列存放 SKU 编码。"
                    "这样即使 SKU 格式不标准（如配饰 MMC、下划线编码），"
                    "标黄的也能被提取，不再局限于标准格式。"),
        "ai_th_header_row": "表头行", "ai_th_column": "SKU 列",
        "ai_th_header": "表头", "ai_th_reason": "判断依据",
        "ai_found_cells": "{sheet}：{n} 个标黄单元格仅通过 AI 识别的列找到（点击查看）",
        "ai_none": "AI 检查了各工作表的表头，没有发现标准扫描之外的 SKU 列。",
        "ai_confirmed_note": "&#10003; 这些列已经过你的人工确认，才用于本次处理。",
        "ai_all_deselected": "你取消了所有识别出的列，本次仅使用标准扫描。",
        "th_samples": "示例值",
        "confirm_cancel": "取消并返回",
        "nav_home": "首页", "nav_help": "使用教程",
        "step_upload": "上传", "step_review": "确认",
        "step_report": "报告",
        "page_upload_heading": "新建处理",
        "page_review": "确认",
        "review_heading": "生成前确认",
        "review_intro": ("以下是从 <strong>{map}</strong> 中提取到的全部内容。"
                         "保持勾选的款号会写入 ProductsList；不需要的请取消勾选，"
                         "然后点击生成按钮 — 按钮上会显示将写入的行数，"
                         "点击之前不会生成任何文件。"),
        "review_ai_intro": ("以下 SKU 列由 Claude 阅读表头后识别。"
                            "请核对示例值是否确实是 SKU 编码；"
                            "来自这些列的款号在下方清单中带“AI”标记，"
                            "发现识别有误就取消勾选。"),
        "review_sku_heading": "提取到的 SKU",
        "th_colorways": "标黄颜色款",
        "th_stock_rows": "库存行数",
        "badge_no_stock": "无库存", "badge_ai": "AI",
        "sel_all": "全选", "sel_none": "全不选",
        "sel_matched": "仅选有库存",
        "review_counter": "已选 {n} / {m} 个款号",
        "btn_build": "生成 ProductsList（{rows} 行）",
        "recent_heading": "最近的运行",
        "recent_none": "还没有运行记录 — 最近 10 次会保留在这里供下载。",
        "recent_map": "文件", "recent_when": "时间",
        "recent_report": "报告", "recent_download": "下载",
        "th_rows_written": "行数",
        "query_used": ("使用的库存查询表：<strong>{name}</strong>，"
                       "共 {rows} 行，上传于 {time}。"),
        "th_source": "来源", "src_ai": "AI", "src_memory": "已记住",
        "badge_pickup": "PickUp",
        "th_value": "内容",
        "th_image": "图片",
        "report_images_note": ("已从 Sell-Thru Map 复制 {n} 张产品图片到表格"
                               "最右侧的 <code>N</code> 列 — 每个基础款号一张，"
                               "放在该款号的第一行。"),
        "review_pickup_note": ("工作表 {sheet}：以其 PickUp 列（{col} 列）为准 — "
                               "共标记 {n} 行；该表的标黄不再用于选择。"),
        "review_pickup_warn": ("{sheet}：{bad} 个 PickUp 值无法识别（已忽略），"
                               "{nosku} 个已标记的行没有 SKU，"
                               "{unmarked} 个标黄 SKU 未标记（因此未被选择）— "
                               "详情见生成后的运行报告。"),
        "pickup_heading": "PickUp 列（显式选择）",
        "pickup_desc": ("当工作表存在表头为 <code>PickUp</code> 的列时，"
                        "该列即为这张表的选择依据：标 <code>Y</code>"
                        "（也接受 yes / x / 1 / pickup / 是）的行会被提取，"
                        "该表的标黄仅作参考显示。没有 PickUp 列的工作表"
                        "仍按标黄提取。"),
        "pickup_sheet_line": ("{sheet}：PickUp 列为 <strong>{col}</strong> 列"
                              "（表头在第 {row} 行），共标记 "
                              "<strong>{n}</strong> 行。"),
        "pickup_details_summary": "{sheet}：通过 PickUp 选中的 {n} 个 SKU（点击查看）",
        "pickup_unrecognized_note": ("PickUp 列中有 {n} 个无法识别的值，"
                                     "对应的行没有被选择（只认 Y / yes / x / "
                                     "1 / pickup / 是）："),
        "pickup_no_sku_note": "{n} 个已标记的行中没有找到可识别的 SKU：",
        "pickup_unmarked_note": ("本表有 {n} 个标黄 SKU 未在 PickUp 列标记，"
                                 "因此没有被选择 — 如需订购，请标上 Y 后重新"
                                 "运行："),
        "pickup_none": "所有工作表都没有 PickUp 列 — 本次按标黄单元格提取（默认方式）。",
        "review_memory_note": ("标注“已记住”的列在之前的运行中已经过人工确认，"
                               "本次无需再调用 AI。如需让工具忘记某个工作表"
                               "已记住的列，生成前将该表的款号全部取消勾选即可。"),
        "by_uploaded": "上传：{name}",
        "by_built": "确认并生成：{name}",
        "th_user": "操作人",
        "excluded_heading": "确认时排除的款号（{n} 个）",
        "excluded_note": "以下标黄款号在确认页被取消勾选，未写入工作簿。",
        "page_progress": "正在处理",
        "progress_heading": "正在处理 — 本页会自动更新进度",
        "stage_start": "正在启动…",
        "stage_scan": "正在扫描表格中的标黄 SKU",
        "stage_match": "正在库存查询表中匹配基础款号",
        "stage_write": "正在生成 ProductsList 工作簿",
        "progress_lost": ("处理被中断（可能是服务器重启）。本次没有保存任何结果，"
                          "请返回重新上传。"),
        "upload_uploading": "正在上传…",
        "upload_analyzing": "上传完成 — 正在读取表头…",
        "ai_failed": ("本次运行 AI 识别不可用：{detail}。"
                      "标准扫描已正常完成，“已记住”的列（如有）也照常使用。"),
        "ai_off": ("AI 识别功能未开启 — 配置 ANTHROPIC_API_KEY 密钥即可启用。"
                   "标准扫描已正常完成。"),
        "step2_title": "第 2 步：提取基础款号",
        "step2_desc": ("每个标黄 SKU 去掉末尾的颜色码（<code>X</code> + 4 位）"
                       "就是基础款号，去重后共 <strong>{n}</strong> 个。"),
        "base_skus_summary": "各基础款号对应的标黄颜色款（点击查看明细）",
        "th_base": "基础款号", "th_from_skus": "来自标黄 SKU",
        "th_sheets": "所在工作表",
        "step3_title": "第 3 步：在库存查询表中匹配全部颜色和尺码",
        "step3_desc": ("查询表 <code>Sku</code> 列以“<code>&nbsp;-&nbsp;</code>”分隔，"
                       "第一段与基础款号完全相同即视为匹配。"
                       "点开任意基础款号，可以查看匹配到的具体数据。"),
        "base_query_rows": "{base}：匹配到 {n} 行库存记录",
        "th_barcode": "条形码", "th_qsku": "SKU", "th_division": "事业部",
        "th_department": "部门", "th_season": "季节", "th_qty": "数量",
        "th_unit_cost": "成本价", "th_unit_retail": "零售价",
        "more_rows": "&hellip; 其余 {n} 行此处未显示，输出文件中包含全部数据。",
        "no_matches": "所有基础款号在库存查询表中都没有匹配到记录。",
        "step4_title": "第 4 步：生成 ProductsList 工作簿",
        "step4_desc": ("匹配到的记录逐行写入（共 <strong>{rows}</strong> 行），"
                       "按基础款号、颜色、尺码排序。<code>TotalCost</code> 和 "
                       "<code>TotalSales</code> 是自动计算的公式；条形码保存为文本"
                       "格式（保留前导零）；Gift Recipient / Organization / "
                       "Position 三列留空；未匹配的基础款号单独列在 "
                       "<code>Unmatched</code> 工作表。"),
        "matched_heading": "匹配成功的基础款号（{n} 个）",
        "th_query_rows": "匹配行数", "th_highlighted_on": "标黄所在工作表",
        "none_label": "无。",
        "unmatched_heading": "未匹配的基础款号（{n} 个）",
        "unmatched_note": ("以下标黄基础款号在库存查询表中<strong>没有找到任何"
                           "记录</strong>，请确认查询表是否为最新版本、是否覆盖"
                           "对应季节。这些款号也已列入输出文件的 "
                           "<code>Unmatched</code> 工作表。"),
        "none_unmatched": "无 — 所有标黄基础款号都匹配成功。",
        "other_fills_heading": "其他底色的 SKU 单元格（{n} 个）",
        "other_fills_note": ("以下 SKU 单元格的底色既不是黄色补货标记，"
                             "也不是表格自带的浅蓝条纹（常见的如灰色备注色）。"
                             "工具<strong>没有提取</strong>这些单元格，"
                             "请人工确认是否有漏标。"),
        "th_fill": "底色",

        # auth: login / setup / forgot
        "auth_login_title": "登录",
        "auth_username": "用户名",
        "auth_password": "密码",
        "auth_login_btn": "登录",
        "auth_login_failed": "用户名或密码错误。",
        "auth_throttled": "尝试次数过多，请等待 {n} 秒后重试。",
        "auth_login_hint": ("账号由管理员在“团队”页创建。还没有账号？"
                            "请联系管理员发送邀请。"),
        "auth_login_setup_link": "创建管理员账号（需要初始化密码）",
        "auth_login_forgot_link": "忘记密码？",
        "auth_setup_title": "管理员初始化",
        "auth_setup_hint": ("本页使用服务器的初始化密码来创建（或找回）管理员账号。"
                            "如果账号已存在，密码会被重置为你在此输入的新密码；"
                            "留空的字段保持原值不变。"),
        "auth_setup_code": "初始化密码",
        "auth_setup_display": "显示名称（可选）",
        "auth_setup_email": "邮箱（可选 — 用于找回密码）",
        "auth_setup_btn": "创建管理员账号",
        "auth_setup_wrong_code": "初始化密码不正确。",
        "auth_setup_disabled": "初始化不可用 — 服务器没有配置初始化密码。",
        "auth_bad_username": ("用户名需为 2–32 个字符：小写字母、数字、点、"
                              "横线或下划线，且以字母或数字开头。"),
        "auth_bad_password": "密码至少需要 8 个字符。",
        "auth_bad_email": "邮箱地址格式不正确。",
        "auth_email_taken": "该邮箱已被其他账号使用。",
        "auth_username_taken": "该用户名已被占用。",
        "auth_forgot_title": "重置密码",
        "auth_forgot_intro": ("输入账号绑定的邮箱地址。如果该邮箱已确认，"
                              "重置链接会发送到该邮箱。"),
        "auth_forgot_email": "邮箱地址",
        "auth_forgot_btn": "发送重置链接",
        "auth_forgot_sent": "如果该邮箱对应某个账号，重置链接已在路上。",
        "auth_forgot_disabled": ("邮件功能未配置 — 请联系管理员在“团队”页"
                                 "为你设置新密码。"),
        "auth_pw_invite_title": "设置你的密码",
        "auth_pw_reset_title": "设置新密码",
        "auth_pw_new": "新密码（至少 8 个字符）",
        "auth_pw_btn": "保存密码",
        "auth_token_dead_title": "链接已失效",
        "auth_token_dead": "该链接已过期或已被使用，请重新获取。",
        "auth_verified_title": "邮箱已确认",
        "auth_verified_msg": "邮箱确认成功 — 此账号现在可以使用“忘记密码”功能。",
        "auth_go_login": "去登录",
        "err_session_expired": "登录已过期，请重新登录。",
        "err_auth_unconfigured": "服务器尚未配置身份验证。",
        "nav_team": "团队",
        "nav_logout": "退出",

        # auth: team page
        "team_heading": "团队",
        "team_th_username": "用户名",
        "team_th_display": "显示名称",
        "team_th_email": "邮箱",
        "team_th_status": "状态",
        "team_th_role": "角色",
        "team_th_actions": "操作",
        "team_status_pending": "待接受邀请",
        "team_status_verified": "已确认",
        "team_status_unverified": "未确认",
        "team_role_admin": "管理员",
        "team_role_member": "成员",
        "team_resend_btn": "重发邀请",
        "team_delete_btn": "删除",
        "team_add_heading": "添加同事",
        "team_add_note": ("填写邮箱时会发送邀请链接，由同事自行设置密码"
                          "（初始密码留空即可）；不填邮箱则需要设置初始密码，"
                          "并私下告知对方。"),
        "team_add_note_manual": "邮件功能未配置：请设置初始密码，并私下告知对方。",
        "team_initial_pw_label": "初始密码（不发送邀请时必填）",
        "team_is_admin_label": "设为管理员",
        "team_add_btn": "添加",
        "team_email_heading": "邮箱地址",
        "team_email_label": "邮箱地址（留空即解除绑定）",
        "team_target_label": "账号（用户名 — 管理员可填写任何账号）",
        "team_verify_btn": "重发确认邮件",
        "team_pw_heading": "修改密码",
        "team_pw_note": ("修改密码会使该账号在所有设备上退出登录 — "
                         "修改自己的密码时也包括当前页面。"),
        "team_pw_label": "新密码（至少 8 个字符）",
        "team_save_btn": "保存",
        "msg_user_added": "账号已创建。",
        "msg_user_deleted": "账号已删除。",
        "msg_invite_sent": "邀请已发送。",
        "msg_invite_resent": "邀请已重新发送。",
        "msg_invite_send_failed": ("账号已创建，但邀请邮件发送失败 — "
                                   "请使用“重发邀请”，或通过“修改密码”"
                                   "为其设置初始密码。"),
        "msg_send_failed": "邮件发送失败，请重试。",
        "msg_pw_changed": "密码已修改，该账号的所有登录会话均已退出。",
        "msg_email_saved": ("邮箱已保存，确认邮件已发送 — 在邮箱确认之前，"
                            "“忘记密码”功能暂不可用。"),
        "msg_email_saved_noconfirm": ("邮箱已保存。邮件功能未配置，"
                                      "该邮箱暂时无法确认，也无法用于找回密码。"),
        "msg_email_saved_sendfail": ("邮箱已保存，但确认邮件发送失败 — "
                                     "请点击“重发确认邮件”重试。"),
        "msg_email_removed": "邮箱已解除绑定。",
        "msg_verify_sent": "确认邮件已发送。",
        "err_not_admin": "只有管理员可以执行此操作。",
        "err_no_such_user": "账号不存在。",
        "err_delete_self": "不能删除自己的账号。",
        "err_delete_last_admin": "不能删除最后一个管理员账号。",
        "err_pending_only": "该账号并非待接受邀请状态。",
        "err_no_email_on_file": "该账号没有绑定邮箱。",
        "err_email_off": "邮件功能未配置。",
        "err_password_required": "无法发送邀请邮件时，必须设置初始密码。",

        # errors
        "error_heading": "出错了。",
        "back_to_upload": "&larr; 返回首页",
        "err_map_required": "请先选择 Sell-Thru Map 文件（.xlsx），每次处理都需要上传。",
        "err_query_rejected": ("上传的库存查询表未通过校验，之前保存的版本（如有）"
                               "不受影响。原因：{detail}"),
        "err_no_query": "还没有保存库存查询表，请在“库存查询表”一栏选择文件后重试。",
        "err_map_failed": ("无法读取该 Sell-Thru Map 文件，请确认它是有效的 .xlsx "
                           "工作簿。错误详情：{detail}"),
        "err_too_large": "上传文件过大：合计不能超过 200 MB。",
        "err_not_found": "找不到该文件 — 可能已被自动清理（仅保留最近 10 次输出）。",
        "err_server": "服务器出现异常，详情请查看应用日志。",
    },
}
