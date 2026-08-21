"""
Translations for the Cut Planning Model v2 UI.

Two independent lookup tables:
  - TRANSLATIONS: general UI text (labels, buttons, descriptions, messages),
    keyed by a short string ID.
  - COLUMN_LABELS: display label for each INTERNAL column name used across
    Tab 1 / Tab 2 / Tab 3's tables. The internal column name itself (used in
    data-col attributes, Python dict keys, JS lookups, etc.) always stays in
    English - only the label shown to the user changes with the language.

Usage from Python: translate(key, lang) / column_label(col, lang)
Usage from Jinja:   {{ t('key') }} / {{ col_label(col) }}   (registered as
Jinja globals in app.py, bound to the current session's language)
"""

DEFAULT_LANG = "en"
SUPPORTED_LANGS = ["en", "th"]

TRANSLATIONS = {
    "app_title": {"en": "Cut Planning Model", "th": "โมเดลวางแผนการตัด"},

    # ---- Tabs ----
    "tab1_label": {"en": "Tab 1: Buffer Cutting Order Form", "th": "แท็บ 1: ใบสั่งตัดบัฟเฟอร์"},
    "tab2_label": {"en": "Tab 2: WIP Upload", "th": "แท็บ 2: อัปโหลด WIP"},
    "tab3_label": {"en": "Tab 3: Cut Plan", "th": "แท็บ 3: แผนการตัด"},

    # ---- Language switcher ----
    "language_label": {"en": "Language", "th": "ภาษา"},

    # ---- Tab 1 ----
    "tab1_desc": {
        "en": "Upload the Buffer Cutting Order Form Excel file and it's extracted into the 17-field "
              "schema below, right on this tab. Only rows whose Sewing Line starts with VS are kept — "
              "everything else is filtered out automatically. If the file is missing any of the 17 "
              "fields, that field is simply left blank in the extracted output.",
        "th": "อัปโหลดไฟล์ Excel ใบสั่งตัดบัฟเฟอร์ ระบบจะดึงข้อมูลลงในแบบฟอร์ม 17 ช่องด้านล่างในแท็บนี้เลย "
              "จะเก็บเฉพาะแถวที่ Sewing Line ขึ้นต้นด้วย VS เท่านั้น ส่วนที่เหลือจะถูกกรองออกโดยอัตโนมัติ "
              "หากไฟล์ขาดข้อมูลในช่องใดช่องหนึ่งจาก 17 ช่อง ช่องนั้นจะถูกปล่อยว่างในผลลัพธ์ที่ดึงออกมา",
    },
    "tab1_file_label": {"en": "Buffer Cutting Order Form (.xlsx)", "th": "ไฟล์ใบสั่งตัดบัฟเฟอร์ (.xlsx)"},
    "use_sample_data": {"en": "Use sample data", "th": "ใช้ข้อมูลตัวอย่าง"},
    "use_sample_wip_data": {"en": "Use sample data", "th": "ใช้ข้อมูลตัวอย่าง"},
    "extract_data_btn": {"en": "Extract data", "th": "ดึงข้อมูล"},
    "target_fields_label": {"en": "Target fields (extracted in this order):", "th": "ช่องข้อมูลเป้าหมาย (เรียงตามลำดับที่ดึง):"},
    "fields_not_found": {
        "en": "These fields were not found in the input file and are shown blank:",
        "th": "ไม่พบช่องข้อมูลเหล่านี้ในไฟล์ต้นฉบับ จึงแสดงเป็นค่าว่าง:",
    },
    "extracted_at": {"en": "Extracted", "th": "ดึงข้อมูลเมื่อ"},
    "rows_kept": {"en": "row(s) kept", "th": "แถวที่เก็บไว้"},
    "rows_filtered_out": {
        "en": 'row(s) filtered out (Sewing Line not starting with "VS")',
        "th": 'แถวที่ถูกกรองออก (Sewing Line ไม่ได้ขึ้นต้นด้วย "VS")',
    },
    "all_fields_found": {"en": "all 17 fields found", "th": "พบครบทั้ง 17 ช่อง"},
    "download_excel": {"en": "Download Excel", "th": "ดาวน์โหลด Excel"},
    "no_rows_extracted": {
        "en": "No rows were extracted from this input file (after filtering).",
        "th": "ไม่มีแถวข้อมูลที่ดึงได้จากไฟล์นี้ (หลังจากกรองแล้ว)",
    },

    # ---- Tab 2 ----
    "tab2_desc": {
        "en": "Upload a Work-in-Process (WIP) buffer report here. It's extracted into a fixed set of "
              "fields matching the known WIP template (per-sewing-line targets, actuals, WIP "
              "quantities, lead times, and shortfall reasons). If a file doesn't match that template, "
              "it falls back to a raw, no-schema preview instead of failing outright.",
        "th": "อัปโหลดรายงานบัฟเฟอร์ Work-in-Process (WIP) ที่นี่ ระบบจะดึงข้อมูลตามชุดฟิลด์คงที่ที่ตรงกับเทมเพลต WIP "
              "ที่รู้จัก (เป้าหมาย ค่าจริง ปริมาณ WIP เวลานำ และสาเหตุที่ขาดเป้า ต่อไลน์เย็บ) "
              "หากไฟล์ไม่ตรงกับเทมเพลตนั้น ระบบจะแสดงตัวอย่างข้อมูลดิบแบบไม่มีโครงสร้างแทนการแจ้งข้อผิดพลาด",
    },
    "wip_file_label": {"en": "WIP file (.xlsx)", "th": "ไฟล์ WIP (.xlsx)"},
    "upload_wip_btn": {"en": "Upload WIP file", "th": "อัปโหลดไฟล์ WIP"},
    "wip_template_mismatch": {
        "en": "This file didn't match the known WIP template (couldn't find the header marker in "
              "column B), so it's shown here as a raw, unmapped preview instead.",
        "th": "ไฟล์นี้ไม่ตรงกับเทมเพลต WIP ที่รู้จัก (ไม่พบหัวข้อในคอลัมน์ B) "
              "จึงแสดงเป็นตัวอย่างข้อมูลดิบแบบไม่มีการแม็ปข้อมูลแทน",
    },
    "uploaded_at": {"en": "Uploaded", "th": "อัปโหลดเมื่อ"},
    "columns_count": {"en": "column(s)", "th": "คอลัมน์"},
    "showing_first_rows": {"en": "showing first", "th": "แสดง"},
    "rows_word": {"en": "rows", "th": "แถวแรก"},
    "matched_wip_template": {"en": "matched the WIP template", "th": "ตรงกับเทมเพลต WIP"},
    "no_rows_in_file": {"en": "This file has no rows.", "th": "ไฟล์นี้ไม่มีข้อมูลแถวใดเลย"},

    # ---- Tab 3 ----
    "tab3_desc": {
        "en": "Plans which cutting tables to work on first, per Sewing Line / JobCut - Suffix / Mark "
              "Type. Completed tables are skipped. Colorways sharing the same Table ID are combined "
              "into one quantity. Tables are planned smallest Table ID first, adding tables until the "
              "running quantity reaches that Mark Type's daily sewing target. Output is split into two "
              "tables by building, based on Sewing Line. For Mark Type 101, if a JobCut's Table No. "
              "(Mark Type 101) field has a value (e.g. \"23-47\"), only cutting tables inside that range "
              "are used for that Mark Type — tables outside it are never selected automatically.",
        "th": "วางแผนว่าจะตัดโต๊ะไหนก่อน ตามกลุ่ม Sewing Line / JobCut - Suffix / Mark Type "
              "โต๊ะที่เสร็จแล้วจะถูกข้าม สีเดียวกันที่อยู่ในหมายเลขโต๊ะเดียวกันจะถูกรวมเป็นจำนวนเดียว "
              "จะวางแผนโต๊ะที่มีหมายเลขน้อยที่สุดก่อน แล้วเพิ่มโต๊ะไปเรื่อยๆ จนกว่าจำนวนสะสมจะถึงเป้าหมายการเย็บต่อวันของ Mark Type นั้น "
              "ผลลัพธ์จะถูกแบ่งเป็นสองตารางตามอาคาร โดยอิงจาก Sewing Line "
              "สำหรับ Mark Type 101 หากช่อง Table No. (Mark Type 101) ของ JobCut นั้นมีค่า (เช่น \"23-47\") "
              "จะใช้เฉพาะโต๊ะตัดที่อยู่ในช่วงนั้นสำหรับ Mark Type นี้ — โต๊ะที่อยู่นอกช่วงจะไม่ถูกเลือกโดยอัตโนมัติ",
    },
    "generate_cut_plan_btn": {"en": "Generate cut plan", "th": "สร้างแผนการตัด"},
    "flag_notice_title": {"en": "Reminder:", "th": "ข้อควรทราบ:"},
    "flag_notice_text": {
        "en": "if a JobCut has more than one Mark Type that needs decoration, you don't have to plan "
              "every Mark Type, and you don't have to plan enough tables to fully reach the Sewing "
              "target per day — it's fine to plan only some tables and leave the rest for later. For "
              "example, you can stop after a few cutting tables instead of planning enough to hit the "
              "full target.",
        "th": "หาก JobCut หนึ่งมี Mark Type ที่ต้องปักมากกว่า 1 แบบ คุณไม่จำเป็นต้องวางแผนให้ครบทุก Mark Type "
              "และไม่จำเป็นต้องวางแผนโต๊ะให้ถึงเป้าหมายการเย็บต่อวันแบบเต็มจำนวน "
              "สามารถวางแผนเพียงบางโต๊ะแล้วเหลือไว้ทีหลังได้ ตัวอย่างเช่น หยุดหลังจากวางแผนไปไม่กี่โต๊ะโดยไม่ต้องวางแผนให้ถึงเป้าเต็มจำนวนก็ได้",
    },
    "decoration_flag_text": {"en": "\u2691 Multiple Mark Types need decoration", "th": "\u2691 มีหลาย Mark Type ที่ต้องปัก"},
    "decoration_flag_title": {
        "en": "This JobCut has more than one Mark Type needing decoration — you don't have to plan "
              "every Mark Type, or plan enough tables to fully reach the Sewing target per day.",
        "th": "JobCut นี้มี Mark Type ที่ต้องปักมากกว่า 1 แบบ — ไม่จำเป็นต้องวางแผนให้ครบทุก Mark Type "
              "หรือวางแผนโต๊ะให้ถึงเป้าหมายการเย็บต่อวันแบบเต็มจำนวน",
    },
    "edit_note": {
        "en": "Every cell below is editable — including Cut Plan Qty and Diff. Editing a row's Cut Plan "
              "Qty updates that same row's Morning/Afternoon to match. Editing a row's Table No. looks "
              "up that table's real quantity from Tab 1's data. Editing Cut Plan Qty or Sewing target "
              "per day re-checks which tables are needed — more tables get pulled in if the target now "
              "needs them, and tables beyond what's needed get dropped. Use + Add row to plan an extra "
              "table by hand, or the \u2715 next to a row to drop it. Note starts blank. Nothing is "
              "saved until you click Save changes & Download Excel — that saves both buildings' tables "
              "together into one workbook.",
        "th": "ทุกช่องด้านล่างสามารถแก้ไขได้ — รวมถึง Cut Plan Qty และ Diff การแก้ไข Cut Plan Qty ของแถวใดจะอัปเดต "
              "Morning/Afternoon ของแถวนั้นให้ตรงกันด้วย การแก้ไข Table No. จะดึงจำนวนจริงของโต๊ะนั้นจากข้อมูลแท็บ 1 "
              "การแก้ไข Cut Plan Qty หรือ Sewing target per day จะตรวจสอบใหม่ว่าต้องใช้โต๊ะไหนบ้าง — "
              "จะดึงโต๊ะเพิ่มเข้ามาถ้าจำเป็นต่อเป้าหมาย และจะตัดโต๊ะที่เกินความจำเป็นออก ใช้ปุ่ม + Add row เพื่อเพิ่มโต๊ะเอง "
              "หรือกด \u2715 ข้างแถวเพื่อลบ Note จะเริ่มต้นเป็นค่าว่าง จะยังไม่มีการบันทึกจนกว่าจะกด Save changes & "
              "Download Excel ซึ่งจะบันทึกตารางของทั้งสองอาคารรวมกันเป็นไฟล์เดียว",
    },
    "generated_label": {"en": "Generated", "th": "สร้างเมื่อ"},
    "table_rows_planned_total": {"en": "table row(s) planned total", "th": "แถวโต๊ะที่วางแผนทั้งหมด"},
    "wip_override_note": {
        "en": "Sewing Target Per Day for these Sewing Lines came from Tab 2's WIP data (not Tab 1):",
        "th": "เป้าหมายการเย็บต่อวันของ Sewing Line เหล่านี้มาจากข้อมูล WIP ในแท็บ 2 (ไม่ใช่แท็บ 1):",
    },
    "save_changes_btn": {"en": "Save changes & Download Excel", "th": "บันทึกการเปลี่ยนแปลง & ดาวน์โหลด Excel"},
    "table_rows_planned": {"en": "table row(s) planned", "th": "แถวโต๊ะที่วางแผน"},
    "add_row_btn": {"en": "+ Add row", "th": "+ เพิ่มแถว"},
    "filter_placeholder": {"en": "Filter\u2026", "th": "กรอง\u2026"},
    "clear_filters_btn": {"en": "Clear filters", "th": "ล้างตัวกรอง"},
    "filter_btn_label": {"en": "Filter \u25be", "th": "กรอง \u25be"},
    "filter_select_all": {"en": "Select All", "th": "เลือกทั้งหมด"},
    "filter_clear_all": {"en": "Clear All", "th": "ล้างทั้งหมด"},
    "filter_no_values": {"en": "No values", "th": "ไม่มีค่า"},
    "filter_blank_value": {"en": "(Blanks)", "th": "(ว่าง)"},
    "rows_shown_of_total": {"en": "shown (filtered)", "th": "ที่แสดง (กรองแล้ว)"},
    "search_placeholder": {"en": "Search this table\u2026", "th": "ค้นหาในตารางนี้\u2026"},
    "nothing_to_plan": {"en": "Nothing to plan for", "th": "ไม่มีอะไรต้องวางแผนสำหรับ"},
    "no_incomplete_tables": {"en": "no incomplete tables found.", "th": "ไม่พบโต๊ะที่ยังทำไม่เสร็จ"},
    "unassigned_note": {
        "en": "These Sewing Lines don't match either building's list — kept here so nothing from the "
              "plan is lost. Update the building lists in cutplan2/model.py if they should belong to "
              "Building 1 or 2.",
        "th": "Sewing Line เหล่านี้ไม่ตรงกับรายชื่ออาคารใดเลย — เก็บไว้ที่นี่เพื่อไม่ให้ข้อมูลในแผนหายไป "
              "แก้ไขรายชื่ออาคารในไฟล์ cutplan2/model.py หากควรจัดอยู่ในอาคาร 1 หรือ 2",
    },
    "building1_label": {"en": "Building 1", "th": "อาคาร 1"},
    "building2_label": {"en": "Building 2", "th": "อาคาร 2"},
    "unassigned_label": {"en": "Unassigned", "th": "ไม่ระบุอาคาร"},
    "of_word": {"en": "of", "th": "ของวันที่"},

    # ---- Flash / status messages ----
    "please_choose_wip_file": {
        "en": "Please choose a WIP .xlsx file to upload, or check 'use bundled sample data'.",
        "th": "กรุณาเลือกไฟล์ WIP .xlsx เพื่ออัปโหลด หรือเลือก 'ใช้ข้อมูลตัวอย่างที่แนบมา'",
    },
    "please_choose_input_file": {
        "en": "Please choose an input .xlsx file, or check 'use bundled sample data'.",
        "th": "กรุณาเลือกไฟล์ .xlsx เพื่ออัปโหลด หรือเลือก 'ใช้ข้อมูลตัวอย่างที่แนบมา'",
    },
    "run_tab1_first": {
        "en": "Run Tab 1 (Buffer Cutting Order Form) first — the cut plan is built from that extracted data.",
        "th": "กรุณารันแท็บ 1 (ใบสั่งตัดบัฟเฟอร์) ก่อน — แผนการตัดจะสร้างจากข้อมูลที่ดึงมาจากแท็บนั้น",
    },

    # ---- JS status strings ----
    "js_saving": {"en": "Saving\u2026", "th": "กำลังบันทึก\u2026"},
    "js_recalculating": {"en": "Recalculating\u2026", "th": "กำลังคำนวณใหม่\u2026"},
    "js_plan_updated": {"en": "Plan updated.", "th": "อัปเดตแผนแล้ว"},
    "js_looking_up_table": {"en": "Looking up table\u2026", "th": "กำลังค้นหาโต๊ะ\u2026"},
    "js_table_qty_updated": {"en": "Table quantity updated.", "th": "อัปเดตจำนวนของโต๊ะแล้ว"},
    "js_error_prefix": {"en": "Error:", "th": "ข้อผิดพลาด:"},
    "js_saved_downloading": {"en": "Saved. Downloading:", "th": "บันทึกแล้ว กำลังดาวน์โหลด:"},
    "js_saved_nothing": {
        "en": "Saved (nothing to download - both buildings are empty).",
        "th": "บันทึกแล้ว (ไม่มีไฟล์ให้ดาวน์โหลด — ทั้งสองอาคารไม่มีข้อมูล)",
    },
    "js_undo_edit": {"en": "Undo this edit", "th": "เลิกทำการแก้ไขนี้"},
    "js_swap_shift": {"en": "Move to the other shift (Morning \u2194 Afternoon)", "th": "ย้ายไปกะตรงข้าม (เช้า \u2194 บ่าย)"},
    "js_remove_row": {"en": "Remove row", "th": "ลบแถว"},
    "js_insert_row_below": {"en": "Insert a new row below this one", "th": "แทรกแถวใหม่ใต้แถวนี้"},
    "js_rows_shown_of_total": {"en": "shown (filtered)", "th": "ที่แสดง (กรองแล้ว)"},
    "js_filter_select_all": {"en": "Select All", "th": "เลือกทั้งหมด"},
    "js_filter_clear_all": {"en": "Clear All", "th": "ล้างทั้งหมด"},
    "js_filter_no_values": {"en": "No values", "th": "ไม่มีค่า"},
    "js_filter_search_placeholder": {"en": "Search values\u2026", "th": "ค้นหาค่า\u2026"},
    "js_filter_no_matches": {"en": "No matching values", "th": "ไม่พบค่าที่ตรงกัน"},
    "js_filter_blank_value": {"en": "(Blanks)", "th": "(ว่าง)"},
    "js_dismiss_notice": {"en": "Dismiss this notice", "th": "ปิดข้อความนี้"},
    "js_dismiss_flag": {"en": "Dismiss this flag", "th": "ปิดป้ายนี้"},
    "js_original_prefix": {"en": "Original:", "th": "ค่าเดิม:"},
    "js_blank": {"en": "(blank)", "th": "(ว่าง)"},
}


COLUMN_LABELS = {
    # Tab 1 (Buffer Cutting Order Form extraction)
    "Sewing Line": {"en": "Sewing Line", "th": "สายการเย็บ"},
    "Sewing Line Code": {"en": "Sewing Line Code", "th": "รหัสสายการเย็บ"},
    "JobCut - Suffix": {"en": "JobCut - Suffix", "th": "JobCut - ต่อท้าย"},
    "Table ID": {"en": "Table ID", "th": "หมายเลขโต๊ะ"},
    "Colorway": {"en": "Colorway", "th": "สี"},
    "Mark Type": {"en": "Mark Type", "th": "ชนิดมาร์ค"},
    "Layer": {"en": "Layer", "th": "ชั้น"},
    "Qty": {"en": "Qty", "th": "จำนวน"},
    "Qty Complete": {"en": "Qty Complete", "th": "จำนวนที่เสร็จ"},
    "Difference": {"en": "Difference", "th": "ส่วนต่าง"},
    "% Complete": {"en": "% Complete", "th": "% เสร็จสิ้น"},
    "Status": {"en": "Status", "th": "สถานะ"},
    "Sewing Target Per Day (with OT)": {"en": "Sewing Target Per Day (with OT)", "th": "เป้าหมายการเย็บต่อวัน (รวม OT)"},
    "Sewing Target Per Day (no OT)": {"en": "Sewing Target Per Day (no OT)", "th": "เป้าหมายการเย็บต่อวัน (ไม่รวม OT)"},
    "Table No. (Mark Type 101)": {"en": "Table No. (Mark Type 101)", "th": "หมายเลขโต๊ะ (Mark Type 101)"},
    "Decoration": {"en": "Decoration", "th": "การปัก/ตกแต่ง"},
    "FG Start Date": {"en": "FG Start Date", "th": "วันที่เริ่ม FG"},
    "Start Cut": {"en": "Start Cut", "th": "วันที่เริ่มตัด"},

    # Tab 3 (Cut Plan)
    "Sewing target per day": {"en": "Sewing target per day", "th": "เป้าหมายการเย็บต่อวัน"},
    "Cut Plan Qty": {"en": "Cut Plan Qty", "th": "จำนวนที่วางแผนตัด"},
    "Diff": {"en": "Diff", "th": "ส่วนต่าง"},
    "Table No.": {"en": "Table No.", "th": "หมายเลขโต๊ะ"},
    "Cut Plan Morning": {"en": "Cut Plan Morning", "th": "แผนตัดช่วงเช้า"},
    "Cut Plan Afternoon": {"en": "Cut Plan Afternoon", "th": "แผนตัดช่วงบ่าย"},
    "Note": {"en": "Note", "th": "หมายเหตุ"},

    # Tab 2 (WIP) - a representative subset; any column not listed here falls
    # back to showing its own (English) name unchanged, so nothing breaks if
    # the WIP template's fields change.
    "Target Hours (OTP 100 Plan)": {"en": "Target Hours (OTP 100 Plan)", "th": "ชั่วโมงเป้าหมาย (แผน OTP 100)"},
    "OTP Sewing Ratio": {"en": "OTP Sewing Ratio", "th": "อัตราส่วน OTP การเย็บ"},
    "Target per Hour": {"en": "Target per Hour", "th": "เป้าหมายต่อชั่วโมง"},
    "Target for Day (with OT)": {"en": "Target for Day (with OT)", "th": "เป้าหมายต่อวัน (รวม OT)"},
    "Target for Day (without OT)": {"en": "Target for Day (without OT)", "th": "เป้าหมายต่อวัน (ไม่รวม OT)"},
    "Target Morning": {"en": "Target Morning", "th": "เป้าหมายช่วงเช้า"},
    "Target Afternoon": {"en": "Target Afternoon", "th": "เป้าหมายช่วงบ่าย"},
    "Target OT": {"en": "Target OT", "th": "เป้าหมาย OT"},
    "Actual Morning": {"en": "Actual Morning", "th": "ยอดจริงช่วงเช้า"},
    "Actual Morning (Extra)": {"en": "Actual Morning (Extra)", "th": "ยอดจริงช่วงเช้า (เพิ่มเติม)"},
    "Actual Afternoon": {"en": "Actual Afternoon", "th": "ยอดจริงช่วงบ่าย"},
    "Actual Afternoon (Extra)": {"en": "Actual Afternoon (Extra)", "th": "ยอดจริงช่วงบ่าย (เพิ่มเติม)"},
    "Actual OT": {"en": "Actual OT", "th": "ยอดจริง OT"},
    "Morning Target Shortfall": {"en": "Morning Target Shortfall", "th": "ขาดเป้าช่วงเช้า"},
    "WIP in Sewing Line": {"en": "WIP in Sewing Line", "th": "WIP ในไลน์เย็บ"},
    "Sewn Yesterday": {"en": "Sewn Yesterday", "th": "ยอดเย็บเมื่อวาน"},
    "Bundled Waiting for Sewing (Full Day)": {"en": "Bundled Waiting for Sewing (Full Day)", "th": "มัดรอเบิกเย็บ (ทั้งวัน)"},
    "WIP Cut Waiting to be Bundled (4-12 hrs)": {"en": "WIP Cut Waiting to be Bundled (4-12 hrs)", "th": "WIP ตัดรอจัดงาน (4-12 ชม.)"},
    "WIP Cut Waiting Health": {"en": "WIP Cut Waiting Health", "th": "สถานะ WIP ตัดรอจัดงาน"},
    "Waiting to Cut": {"en": "Waiting to Cut", "th": "รอตัด"},
    "Lead Time Morning (hrs)": {"en": "Lead Time Morning (hrs)", "th": "Lead Time ช่วงเช้า (ชม.)"},
    "Lead Time Morning Health": {"en": "Lead Time Morning Health", "th": "สถานะ Lead Time ช่วงเช้า"},
    "Lead Time Morning Status": {"en": "Lead Time Morning Status", "th": "สถานะ (Lead Time เช้า)"},
    "Lead Time Afternoon (hrs)": {"en": "Lead Time Afternoon (hrs)", "th": "Lead Time ช่วงบ่าย (ชม.)"},
    "Lead Time Afternoon Health": {"en": "Lead Time Afternoon Health", "th": "สถานะ Lead Time ช่วงบ่าย"},
    "Lead Time Afternoon Status": {"en": "Lead Time Afternoon Status", "th": "สถานะ (Lead Time บ่าย)"},
    "Lead Time OT (hrs)": {"en": "Lead Time OT (hrs)", "th": "Lead Time OT (ชม.)"},
    "Lead Time OT Status": {"en": "Lead Time OT Status", "th": "สถานะ (Lead Time OT)"},
    "Reason Missed Morning Target": {"en": "Reason Missed Morning Target", "th": "สาเหตุที่ไม่ได้เป้าเช้า"},
    "Detail Morning": {"en": "Detail Morning", "th": "รายละเอียด (เช้า)"},
    "Afternoon Target Shortfall": {"en": "Afternoon Target Shortfall", "th": "ขาดเป้าบ่าย"},
    "Reason Missed Afternoon Target": {"en": "Reason Missed Afternoon Target", "th": "สาเหตุที่ไม่ได้เป้าบ่าย"},
    "Detail Afternoon": {"en": "Detail Afternoon", "th": "รายละเอียด (บ่าย)"},
    "OT Target Shortfall (Full Day)": {"en": "OT Target Shortfall (Full Day)", "th": "ขาดเป้า OT (ทั้งวัน)"},
    "Reason Missed OT Target": {"en": "Reason Missed OT Target", "th": "สาเหตุที่ไม่ได้เป้า OT"},
    "Detail OT": {"en": "Detail OT", "th": "รายละเอียด (OT)"},
    "Cause Category": {"en": "Cause Category", "th": "หมวดหมู่สาเหตุ"},
    "Cut Waiting to be Bundled": {"en": "Cut Waiting to be Bundled", "th": "ตัดรอจัดงาน"},
    "Add: Cut Fully Bundled": {"en": "Add: Cut Fully Bundled", "th": "เพิ่ม: ตัดรอจัดงานครบตัว"},
    "Add: WIP Bundled Waiting for Sewing": {"en": "Add: WIP Bundled Waiting for Sewing", "th": "เพิ่ม: WIP จัดงานรอเบิกเย็บ"},
    "OTP Sewing (Reference Date)": {"en": "OTP Sewing (Reference Date)", "th": "OTP เย็บ (วันที่อ้างอิง)"},
    "OT Total": {"en": "OT Total", "th": "รวม OT"},
    "Total without OT": {"en": "Total without OT", "th": "รวมไม่มี OT"},
}


def translate(key: str, lang: str = DEFAULT_LANG) -> str:
    """Look up a UI string by key for the given language, falling back to
    English (and then the key itself) if missing."""
    entry = TRANSLATIONS.get(key)
    if entry is None:
        return key
    return entry.get(lang) or entry.get(DEFAULT_LANG) or key


def column_label(col: str, lang: str = DEFAULT_LANG) -> str:
    """Display label for an internal column name, in the given language.
    Falls back to the column's own (English) name if no translation exists,
    so a WIP template change never breaks display."""
    entry = COLUMN_LABELS.get(col)
    if entry is None:
        return col
    return entry.get(lang) or entry.get(DEFAULT_LANG) or col


def js_translations(lang: str = DEFAULT_LANG) -> dict:
    """Bundle every 'js_*' key into a flat {shortName: translatedText} dict
    for embedding into the page as a JS object (see index.html), so
    client-side status messages can be translated too, e.g. T.saving."""
    result = {}
    for key, entry in TRANSLATIONS.items():
        if key.startswith("js_"):
            short_name = key[len("js_"):]
            result[short_name] = entry.get(lang) or entry.get(DEFAULT_LANG) or key
    return result
