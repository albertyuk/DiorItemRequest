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
        "intro": ("Extracts yellow-highlighted (reorder) SKUs from a sell-thru "
                  "map, looks up all color/size variants in the stored stock "
                  "query export, and produces a filled ProductsList workbook."),
        "stored_query": ("Stored stock query: <strong>{name}</strong> "
                         "&mdash; {rows} rows, uploaded {time}."),
        "no_query_warning": ("<strong>No stock query export stored yet.</strong> "
                             "Upload one below — it is kept for future runs "
                             "until replaced."),
        "map_label": "Sell-thru map (.xlsx) — required each run",
        "query_label": ("Stock query export (.xlsx) — optional, replaces the "
                        "stored one"),
        "process_btn": "Process",
        "spinner": ("&#9203; Processing&hellip; parsing a large map takes up "
                    "to a minute or two. Leave this page open."),

        # report
        "new_run": "&larr; new run",
        "processed_summary": ("Processed <strong>{map}</strong> — wrote "
                              "<strong>{rows}</strong> rows for "
                              "<strong>{matched}</strong> of "
                              "<strong>{total}</strong> highlighted base SKUs."),
        "download_label": "Download {name}",
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
        "intro": ("自动从 Sell-Thru Map（销售进度表）中提取所有标黄（需补货）的 SKU，"
                  "在已保存的库存查询表中匹配出全部颜色和尺码，"
                  "并生成可直接使用的 ProductsList 工作簿。"),
        "stored_query": ("当前保存的库存查询表：<strong>{name}</strong>，"
                         "共 {rows} 行，上传于 {time}。"),
        "no_query_warning": ("<strong>还没有保存库存查询表。</strong>"
                             "请在下方上传；保存后无需每次重复上传，"
                             "之后上传新文件即可替换。"),
        "map_label": "Sell-Thru Map（.xlsx）— 每次都需要上传",
        "query_label": "库存查询表（.xlsx）— 可选，上传后替换之前保存的版本",
        "process_btn": "开始处理",
        "spinner": ("&#9203; 处理中&hellip; 文件较大时需要一两分钟，"
                    "请勿关闭页面。"),

        # report
        "new_run": "&larr; 返回首页",
        "processed_summary": ("<strong>{map}</strong> 处理完成：共识别出 "
                              "<strong>{total}</strong> 个标黄基础款号，其中 "
                              "<strong>{matched}</strong> 个在库存查询表中匹配到记录，"
                              "已写入 <strong>{rows}</strong> 行。"),
        "download_label": "下载 {name}",
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
