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
        "site_title": "Dior 单品申请 — 标黄 SKU 提取工具",
        "page_upload": "上传",
        "page_report": "运行报告",
        "page_error": "错误",
        "lang_switch": "English",

        # index
        "intro": ("从 Sell-Thru Map（销售进度表）中提取标黄（需补货 REORDER）的 SKU，"
                  "在已存储的库存查询表中查找对应的所有颜色和尺码，"
                  "并生成填写完整的 ProductsList 工作簿。"),
        "stored_query": ("已存储的库存查询表：<strong>{name}</strong>"
                         "&mdash; 共 {rows} 行，上传时间 {time}。"),
        "no_query_warning": ("<strong>尚未存储库存查询表。</strong>"
                             "请在下方上传 — 上传后会一直保留供后续运行使用，"
                             "直到被新文件替换。"),
        "map_label": "Sell-Thru Map（.xlsx）— 每次运行都必须上传",
        "query_label": "库存查询导出表（.xlsx）— 可选；上传后将替换已存储的版本",
        "process_btn": "开始处理",
        "spinner": ("&#9203; 处理中&hellip; 解析较大的表格可能需要一到两分钟，"
                    "请不要关闭本页面。"),

        # report
        "new_run": "&larr; 新的运行",
        "processed_summary": ("已处理 <strong>{map}</strong>：共识别 "
                              "<strong>{total}</strong> 个标黄基础款号，其中 "
                              "<strong>{matched}</strong> 个在库存查询表中找到记录，"
                              "共写入 <strong>{rows}</strong> 行。"),
        "download_label": "下载 {name}",
        "steps_heading": "处理步骤",
        "step1_title": "第 1 步 — 扫描所有工作表中的标黄 SKU",
        "step1_desc": ("逐个检查每个单元格中形如 SKU 的值；"
                       "带有黄色“补货”标记填充色的单元格计为标黄。"),
        "sheet_details_summary": "{sheet} — 找到的 {n} 个标黄单元格",
        "th_sheet": "工作表", "th_sku_cells": "SKU 单元格数",
        "th_highlighted_cells": "标黄单元格数",
        "th_unique_skus": "去重后 SKU 数", "th_unique_bases": "去重后基础款号数",
        "th_cell": "单元格", "th_sku_as_highlighted": "SKU（表中原样）",
        "th_base_extracted": "提取出的基础款号",
        "step2_title": "第 2 步 — 提取基础款号",
        "step2_desc": ("从每个标黄 SKU 中去掉末尾的颜色码（<code>X</code> + 4 位字符）；"
                       "去重后共得到 <strong>{n}</strong> 个基础款号。"),
        "base_skus_summary": "各基础款号及其对应的标黄颜色款",
        "th_base": "基础款号", "th_from_skus": "来自标黄 SKU",
        "th_sheets": "所在工作表",
        "step3_title": "第 3 步 — 在库存查询表中查找所有颜色和尺码",
        "step3_desc": ("当查询表 <code>Sku</code> 列以“<code>&nbsp;-&nbsp;</code>”"
                       "分隔的第一段与基础款号完全一致时视为匹配。"
                       "点开任一基础款号即可查看为它取到的数据。"),
        "base_query_rows": "{base} — 共 {n} 行查询记录",
        "th_barcode": "条形码", "th_qsku": "SKU", "th_division": "事业部",
        "th_department": "部门", "th_season": "季节", "th_qty": "数量",
        "th_unit_cost": "单位成本", "th_unit_retail": "单位零售价",
        "more_rows": "&hellip; 另有 {n} 行未在此显示（输出工作簿中包含全部行）。",
        "no_matches": "没有任何基础款号在已存储的查询表中找到记录。",
        "step4_title": "第 4 步 — 生成 ProductsList 工作簿",
        "step4_desc": ("每个匹配的查询行写入一行（共 <strong>{rows}</strong> 行），"
                       "按基础款号、颜色、尺码排序。<code>TotalCost</code> 与 "
                       "<code>TotalSales</code> 以公式写入；条形码以文本格式保存"
                       "（保留前导零）；Gift Recipient / Organization / Position "
                       "三列留空。未匹配的基础款号列在 <code>Unmatched</code> "
                       "工作表中。"),
        "matched_heading": "已匹配的基础款号（{n}）",
        "th_query_rows": "查询行数", "th_highlighted_on": "标黄所在工作表",
        "none_label": "无。",
        "unmatched_heading": "未匹配的基础款号（{n}）",
        "unmatched_note": ("以下标黄基础款号在已存储的库存查询表中<strong>没有任何"
                           "记录</strong>。它们同时也列在输出工作簿的 "
                           "<code>Unmatched</code> 工作表中。"),
        "none_unmatched": "无 — 所有标黄基础款号都在查询表中找到了至少一行记录。",
        "other_fills_heading": "检测到的其他填充色（{n}）",
        "other_fills_note": ("以下 SKU 单元格带有既不是黄色标记、也不是结构性条纹"
                             "底色的填充色（例如灰色备注）。这些单元格<strong>未被"
                             "提取</strong> — 请人工核对。"),
        "th_fill": "填充色",

        # errors
        "error_heading": "出错了。",
        "back_to_upload": "&larr; 返回上传页",
        "err_map_required": "请选择 Sell-Thru Map 文件（.xlsx）— 每次运行都必须上传。",
        "err_query_rejected": ("库存查询文件被拒绝，之前存储的版本（如有）已保留。"
                               "原因：{detail}"),
        "err_no_query": ("尚未存储库存查询表。请在“库存查询导出表”一栏"
                         "上传后重试。"),
        "err_map_failed": ("无法将该 Sell-Thru Map 作为 .xlsx 工作簿处理：{detail}"),
        "err_too_large": "上传文件过大：合计不能超过 200 MB。",
        "err_not_found": "未找到 — 该文件可能已被清理（只保留最近 10 次输出）。",
        "err_server": "服务器发生意外错误 — 详情见应用日志。",
    },
}
