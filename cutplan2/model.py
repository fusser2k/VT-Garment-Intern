"""
Cut Planning Model v2 — Page 1: Data Extraction
=================================================
This is a NEW, separate model from the original cut_plan_model project.
Page 1's job is to correctly read an input Excel file and extract it into
the 14-field schema below, regardless of what the source file's actual
column headers happen to be.

Later pages (e.g. an actual cut-planning/table-selection page) can build on
top of the DataFrame this module produces.
"""

from datetime import datetime, timedelta, date
from typing import Optional, List, Tuple

import json
import os

import pandas as pd
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

FONT_NAME = "Arial"

# The 14 fields Page 1 must extract, in display order.
OUTPUT_COLUMNS = [
    "Sewing Line",
    "JobCut - Suffix",
    "Table ID",
    "Colorway",
    "Mark Type",
    "Layer",
    "Qty",
    "Qty Complete",
    "Difference",
    "% Complete",
    "Status",
    "Sewing Target Per Day",
    "Table No. (Mark Type 101)",
    "Decoration",
]

# Direct aliases: source header -> one of the OUTPUT_COLUMNS above.
# Add more entries here as real production files use different header text.
DIRECT_ALIASES = {
    "Sewing Line": "Sewing Line",
    "No. of sewing line": "Sewing Line",
    "JobCut - Suffix": "JobCut - Suffix",
    "Table ID": "Table ID",
    "Table No.": "Table ID",
    "Colorway": "Colorway",
    "Mark Type": "Mark Type",
    "Layer": "Layer",
    "Qty": "Qty",
    "Total": "Qty",
    "Qty Complete": "Qty Complete",
    "Difference": "Difference",
    "% Complete": "% Complete",
    "Status": "Status",
    "Sewing Target Per Day": "Sewing Target Per Day",
    "Sewing target per day": "Sewing Target Per Day",
    "Table No. (Mark Type 101)": "Table No. (Mark Type 101)",
    "Decoration": "Decoration",
}

# Columns that, together, can build "JobCut - Suffix" when the source file
# keeps Job Cut and Suffix as two separate columns instead of one combined one.
JOBCUT_COL_CANDIDATES = ["Job Cut", "JobCut", "Job Order"]
SUFFIX_COL_CANDIDATES = ["Suffix"]


def load_input(path_or_buffer) -> Tuple[pd.DataFrame, List[str], int]:
    """Read an input Excel file and extract it into the OUTPUT_COLUMNS schema.

    Unlike a strict schema check, this does NOT raise an error if some of
    the 14 fields aren't present in the source file — those fields are
    simply returned blank, and their names are included in the returned
    `missing_columns` list so the caller (web page / CLI) can tell the user
    which fields the source file didn't provide.

    Rows whose Sewing Line does not start with "VS" are filtered out (the
    planning team only tracks VS-prefixed sewing lines, e.g. VSEW012,
    VS02+06). The number of rows dropped this way is returned as
    `filtered_out_count`.

    Returns (extracted_df, missing_columns, filtered_out_count).
    """
    raw = pd.read_excel(path_or_buffer)
    raw_cols = set(raw.columns)
    n = len(raw)

    extracted = {}
    missing_columns = []

    # --- JobCut - Suffix: combined column, or built from two separate ones ---
    if "JobCut - Suffix" in raw_cols:
        extracted["JobCut - Suffix"] = raw["JobCut - Suffix"].astype(str)
    else:
        jobcut_col = next((c for c in JOBCUT_COL_CANDIDATES if c in raw_cols), None)
        suffix_col = next((c for c in SUFFIX_COL_CANDIDATES if c in raw_cols), None)
        if jobcut_col is not None:
            if suffix_col is not None:
                extracted["JobCut - Suffix"] = (
                    raw[jobcut_col].astype(str) + "-" + raw[suffix_col].astype(str)
                )
            else:
                extracted["JobCut - Suffix"] = raw[jobcut_col].astype(str)
        else:
            extracted["JobCut - Suffix"] = pd.Series([""] * n)
            missing_columns.append("JobCut - Suffix")

    # --- Everything else: direct alias lookup ---
    for target_col in OUTPUT_COLUMNS:
        if target_col == "JobCut - Suffix":
            continue  # handled above

        source_col = next(
            (src for src, tgt in DIRECT_ALIASES.items() if tgt == target_col and src in raw_cols),
            None,
        )
        if source_col is not None:
            extracted[target_col] = raw[source_col]
        else:
            extracted[target_col] = pd.Series([None] * n)
            missing_columns.append(target_col)

    df = pd.DataFrame(extracted, columns=OUTPUT_COLUMNS)

    # --- Filter: only keep rows whose Sewing Line starts with "VS" ---
    before_count = len(df)
    sewing_line_str = df["Sewing Line"].astype(str).str.strip()
    df = df[sewing_line_str.str.startswith("VS")].reset_index(drop=True)
    filtered_out_count = before_count - len(df)

    return df, missing_columns, filtered_out_count


TITLE_FONT_COLOR = "1F4E78"
WARN_FONT_COLOR = "C00000"


def write_extracted_workbook(
    df: pd.DataFrame,
    missing_columns: List[str],
    output_path: str,
    generated_at: Optional[datetime] = None,
    filtered_out_count: int = 0,
) -> None:
    """Write the formatted Page 1 output workbook: Extracted Data + a
    Column Mapping Notes sheet documenting what was / wasn't found."""
    generated_at = generated_at or datetime.now()

    wb = Workbook()

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=10)
    cell_font = Font(name=FONT_NAME, size=10)
    title_font = Font(name=FONT_NAME, size=10, bold=True, color=TITLE_FONT_COLOR)
    warn_font = Font(name=FONT_NAME, size=10, bold=True, color=WARN_FONT_COLOR)
    thin = Side(style="thin", color="B7B7B7")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")

    # ---- Sheet 1: Extracted Data ----
    ws = wb.active
    ws.title = "Extracted Data"

    title_text = f"Extracted: {generated_at.strftime('%Y-%m-%d %H:%M')}"
    if filtered_out_count:
        title_text += f"   |   Filtered out {filtered_out_count} row(s) whose Sewing Line did not start with \"VS\""
    ws.cell(row=1, column=1, value=title_text).font = title_font
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(OUTPUT_COLUMNS))

    header_row = 2
    if missing_columns:
        ws.cell(
            row=2, column=1,
            value=f"Fields not found in the input file (left blank): {', '.join(missing_columns)}",
        ).font = warn_font
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(OUTPUT_COLUMNS))
        header_row = 3

    headers = list(df.columns)
    for j, h in enumerate(headers, start=1):
        c = ws.cell(row=header_row, column=j, value=h)
        c.font = header_font
        c.fill = header_fill
        c.alignment = center
        c.border = border

    for i, row in enumerate(df.itertuples(index=False), start=header_row + 1):
        for j, val in enumerate(row, start=1):
            c = ws.cell(row=i, column=j, value=val if pd.notna(val) else None)
            c.font = cell_font
            c.alignment = center
            c.border = border

    widths = [12, 16, 10, 12, 10, 8, 8, 12, 11, 11, 10, 18, 20, 12]
    for j, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(j)].width = w

    ws.freeze_panes = f"A{header_row + 1}"
    last_row = max(len(df) + header_row, header_row)
    ws.auto_filter.ref = f"A{header_row}:{get_column_letter(len(headers))}{last_row}"

    # ---- Sheet 2: Column Mapping Notes ----
    ws2 = wb.create_sheet("Column Mapping Notes")
    ws2.column_dimensions["A"].width = 35
    ws2.column_dimensions["B"].width = 70

    notes = [
        ("CUT PLANNING MODEL v2 — PAGE 1: DATA EXTRACTION", ""),
        ("", ""),
        ("Filter applied", f"Only rows with Sewing Line starting with \"VS\" are kept. {filtered_out_count} row(s) filtered out."),
        ("", ""),
        ("Field", "Status"),
    ]
    for col in OUTPUT_COLUMNS:
        status = "Not found in input file — left blank" if col in missing_columns else "Found and extracted"
        notes.append((col, status))

    for i, (a, b) in enumerate(notes, start=1):
        ca = ws2.cell(row=i, column=1, value=a)
        cb = ws2.cell(row=i, column=2, value=b)
        if i == 1:
            ca.font = Font(name=FONT_NAME, size=11, bold=True, color=TITLE_FONT_COLOR)
        elif i == 3:
            ca.font = Font(name=FONT_NAME, size=10, bold=True)
            cb.font = cell_font
        elif i == 5:
            ca.font = header_font
            cb.font = header_font
            ca.fill = header_fill
            cb.fill = header_fill
        else:
            ca.font = cell_font
            cb.font = warn_font if "Not found" in b else cell_font

    wb.save(output_path)


def run(input_path: str, output_path: str):
    """Convenience entry point: extract, write, return (df, missing_columns, filtered_out_count)."""
    df, missing_columns, filtered_out_count = load_input(input_path)
    write_extracted_workbook(df, missing_columns, output_path, filtered_out_count=filtered_out_count)
    return df, missing_columns, filtered_out_count


WIP_HEADER_MARKER = "ไลน์เย็บ"
WIP_STOP_MARKER = "ไลน์"  # marks the start of the unrelated table-status mini-table further down the sheet

# Maps this WIP template's Excel column letters to friendly English field
# names. The sheet has TWO header blocks (one per set of sewing lines) that
# repeat the exact same column layout, so one static mapping covers both.
WIP_COLUMN_MAP = {
    "B": "Sewing Line",
    "C": "Target Hours (OTP 100 Plan)",
    "D": "OTP Sewing Ratio",
    "E": "Target per Hour",
    "F": "Target for Day (with OT)",
    "G": "Target for Day (without OT)",
    "H": "Target Morning",
    "I": "Target Afternoon",
    "J": "Target OT",
    "K": "Actual Morning",
    "L": "Actual Morning (Extra)",
    "M": "Actual Afternoon",
    "N": "Actual Afternoon (Extra)",
    "O": "Actual OT",
    "Q": "Morning Target Shortfall",
    "R": "WIP in Sewing Line",
    "S": "Sewn Yesterday",
    "T": "Bundled Waiting for Sewing (Full Day)",
    "U": "WIP Cut Waiting to be Bundled (4-12 hrs)",
    "V": "Waiting to Cut",
    "W": "Lead Time Morning (hrs)",
    "X": "Lead Time Morning Status",
    "Y": "Lead Time Afternoon (hrs)",
    "Z": "Lead Time Afternoon Status",
    "AA": "Lead Time OT (hrs)",
    "AB": "Lead Time OT Status",
    "AC": "Reason Missed Morning Target",
    "AD": "Detail Morning",
    "AE": "Afternoon Target Shortfall",
    "AF": "Reason Missed Afternoon Target",
    "AG": "Detail Afternoon",
    "AH": "OT Target Shortfall (Full Day)",
    "AI": "Reason Missed OT Target",
    "AJ": "Detail OT",
    "AK": "Cause Category",
    "AL": "Cut Waiting to be Bundled",
    "AM": "Add: Cut Fully Bundled",
    "AN": "Add: WIP Bundled Waiting for Sewing",
    "AO": "OTP Sewing (Reference Date)",
    "AP": "OT Total",
    "AQ": "Total without OT",
}

WIP_COLUMNS = list(WIP_COLUMN_MAP.values())

# Column U's conditional-formatting thresholds in the source file (Excel's
# own red/green/orange coloring on that cell) - used to derive a text field
# since, unlike the Lead Time columns, this one has no adjacent status
# column already spelling out what the color means.
WIP_CUT_WAITING_LOW = 3.5
WIP_CUT_WAITING_HIGH = 12.5


def classify_wip_cut_waiting(value) -> str:
    """Mirror the source file's conditional formatting on column U ("WIP
    ตัดรอจัดงาน 4-12 ชม"): red = underproduction (too little cut fabric
    queued up waiting to be bundled), green = neutral/healthy, orange =
    overproduction (too much queued up)."""
    try:
        v = float(value)
    except (TypeError, ValueError):
        return ""
    if v < WIP_CUT_WAITING_LOW:
        return "Underproduction"
    if v > WIP_CUT_WAITING_HIGH:
        return "Overproduction"
    return "Neutral"


# Insert the derived health field right after the raw value it's based on.
_insert_at = WIP_COLUMNS.index("WIP Cut Waiting to be Bundled (4-12 hrs)") + 1
WIP_COLUMNS.insert(_insert_at, "WIP Cut Waiting Health")

# Lead Time (W/Y/AA) conditional-formatting thresholds. The source file uses
# TWO different threshold sets depending on which block of sewing lines a
# row belongs to: individually-named lines (the first block, tighter
# 9.5-16.5 "healthy" band) vs. merged lines like "VS02+06" (the second
# block, wider 15.5-32.5 band, since a merged line naturally carries more
# backlog). load_wip() tracks which block each row came from and picks the
# matching thresholds automatically.
LEAD_TIME_INDIVIDUAL_LOW = 9.5
LEAD_TIME_INDIVIDUAL_HIGH = 16.5
LEAD_TIME_MERGED_LOW = 15.5
LEAD_TIME_MERGED_HIGH = 32.5


def classify_lead_time(value, is_merged_block: bool = False) -> str:
    """Mirror the source file's conditional formatting on the Lead Time
    Morning/Afternoon columns (W/Y): red = underproduction (line is running
    low on queued work - at risk of sitting idle), green =
    neutral/healthy, orange = overproduction (too much backlog queued up)."""
    try:
        v = float(value)
    except (TypeError, ValueError):
        return ""
    low, high = (LEAD_TIME_MERGED_LOW, LEAD_TIME_MERGED_HIGH) if is_merged_block else (
        LEAD_TIME_INDIVIDUAL_LOW, LEAD_TIME_INDIVIDUAL_HIGH
    )
    if v < low:
        return "Underproduction"
    if v > high:
        return "Overproduction"
    return "Neutral"


for _col in ["Lead Time Morning (hrs)", "Lead Time Afternoon (hrs)"]:
    _pos = WIP_COLUMNS.index(_col) + 1
    WIP_COLUMNS.insert(_pos, _col.replace("(hrs)", "Health").replace("  ", " ").strip())


def load_wip(path_or_buffer) -> pd.DataFrame:
    """Read a Work-in-Process (WIP) buffer report and extract it into the
    fixed WIP_COLUMNS schema.

    This template repeats its header row once per group of sewing lines
    (e.g. once for the individually-named lines, again for the merged
    lines like "VS02+06"). Rather than assuming a fixed row range, this
    scans column B for the literal header text ("ไลน์เย็บ") to find where
    each block starts, and collects every non-blank row under it as data -
    so it keeps working even if a future day's file has more or fewer
    sewing lines, or an extra block.

    Extraction stops the moment column B hits "ไลน์" (a different, unrelated
    mini-table further down the same sheet that tracks individual cutting
    tables, not sewing lines - the WIP_STOP_MARKER) - so that table is
    correctly left out rather than misread as more sewing-line rows.
    """
    wb = openpyxl.load_workbook(path_or_buffer, data_only=True)
    ws = wb[wb.sheetnames[0]]

    col_indices = {letter: get_column_letter_index(letter) for letter in WIP_COLUMN_MAP}

    rows_out = []
    collecting = False
    block_index = -1  # increments each time a new header block starts
    for r in range(1, ws.max_row + 1):
        b_val = ws.cell(row=r, column=2).value
        b_str = str(b_val).strip() if b_val is not None else ""

        if b_str == WIP_HEADER_MARKER:
            collecting = True
            block_index += 1
            continue
        if b_str == WIP_STOP_MARKER:
            break
        if collecting and b_str:
            row_data = {}
            for letter, field_name in WIP_COLUMN_MAP.items():
                row_data[field_name] = ws.cell(row=r, column=col_indices[letter]).value
            row_data["WIP Cut Waiting Health"] = classify_wip_cut_waiting(
                row_data.get("WIP Cut Waiting to be Bundled (4-12 hrs)")
            )
            # First block = individually-named lines (tighter thresholds);
            # any block after that = merged lines like "VS02+06" (wider
            # thresholds), matching this template's actual layout.
            is_merged_block = block_index > 0
            row_data["Lead Time Morning Health"] = classify_lead_time(
                row_data.get("Lead Time Morning (hrs)"), is_merged_block
            )
            row_data["Lead Time Afternoon Health"] = classify_lead_time(
                row_data.get("Lead Time Afternoon (hrs)"), is_merged_block
            )
            rows_out.append(row_data)

    return pd.DataFrame(rows_out, columns=WIP_COLUMNS)


def get_column_letter_index(letter: str) -> int:
    from openpyxl.utils import column_index_from_string
    return column_index_from_string(letter)


def write_wip_workbook(df: pd.DataFrame, output_path: str, generated_at: Optional[datetime] = None) -> None:
    """Write a formatted, downloadable version of the extracted WIP data."""
    generated_at = generated_at or datetime.now()
    wb = Workbook()
    ws = wb.active
    ws.title = "WIP Data"

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=10)
    cell_font = Font(name=FONT_NAME, size=10)
    title_font = Font(name=FONT_NAME, size=10, bold=True, color=TITLE_FONT_COLOR)
    thin = Side(style="thin", color="B7B7B7")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")

    ws.cell(row=1, column=1, value=f"Extracted: {generated_at.strftime('%Y-%m-%d %H:%M')}").font = title_font
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(WIP_COLUMNS))

    header_row = 2
    for j, h in enumerate(WIP_COLUMNS, start=1):
        c = ws.cell(row=header_row, column=j, value=h)
        c.font = header_font
        c.fill = header_fill
        c.alignment = center
        c.border = border

    for i, row in enumerate(df.itertuples(index=False), start=header_row + 1):
        for j, val in enumerate(row, start=1):
            c = ws.cell(row=i, column=j, value=val if pd.notna(val) else None)
            c.font = cell_font
            c.alignment = center
            c.border = border

    for j in range(1, len(WIP_COLUMNS) + 1):
        ws.column_dimensions[get_column_letter(j)].width = 20

    ws.freeze_panes = f"A{header_row + 1}"
    last_row = max(len(df) + header_row, header_row)
    ws.auto_filter.ref = f"A{header_row}:{get_column_letter(len(WIP_COLUMNS))}{last_row}"

    wb.save(output_path)


def load_wip_raw(path_or_buffer) -> pd.DataFrame:
    """Read ANY Excel file with NO assumed schema - kept as a fallback for
    WIP files that don't match the known WIP_COLUMN_MAP template (e.g. a
    different sheet layout). load_wip() above should be tried first."""
    df = pd.read_excel(path_or_buffer)
    return df


# ---------------------------------------------------------------------------
# Tab 3: Cut Plan
# ---------------------------------------------------------------------------

CUT_PLAN_COLUMNS = [
    "Sewing Line",
    "JobCut - Suffix",
    "Sewing target per day",
    "Cut Plan Qty",
    "Diff",
    "Mark Type",
    "Table No.",
    "Cut Plan Morning",
    "Cut Plan Afternoon",
    "Note",
]


def split_morning_afternoon(selected: list) -> dict:
    """Decide, for one (Sewing Line, JobCut - Suffix, Mark Type) group's
    already-selected tables (in ascending Table ID order), which of Cut
    Plan Morning / Cut Plan Afternoon each table's quantity goes into.

    Rule:
      - Exactly one table selected -> its whole quantity goes to Morning,
        by default.
      - More than one table selected -> split into a "top" part (Morning)
        and a "bottom" part (Afternoon): walk the tables in order, keeping
        each one in the top/Morning part and accumulating its quantity,
        until the running total reaches HALF OF THE GROUP'S CUT PLAN QTY
        (the total quantity actually being allocated across this same
        Mark Type + JobCut - NOT half of Sewing Target Per Day, since the
        selected tables don't always add up anywhere near the target - the
        split still has to happen even when the group is nowhere near its
        target). The table whose addition reaches that halfway point is
        still part of the top/Morning part; every table AFTER that point
        goes to the bottom/Afternoon part. When the total can't be split
        into two exactly equal halves, this naturally puts the LARGER
        portion in Morning and the smaller one in Afternoon, since the
        crossing table stays in the top/Morning part.
      - Whenever there's more than one table, SOME split always happens -
        if one large table (usually the last one in Table ID order) is big
        enough that the halfway point isn't crossed until it's added,
        leaving nothing for Afternoon, that table is moved to Afternoon
        instead so the group isn't left entirely in Morning. (This is the
        one case where the larger side can end up in Afternoon rather than
        Morning - guaranteeing an actual split takes priority here.)

    `selected` is a list of (table_id, qty) tuples, already in the order
    they were selected (ascending Table ID). Returns {table_id: (morning, afternoon)}.
    """
    if len(selected) <= 1:
        return {tid: (qty, 0) for tid, qty in selected}

    total_qty = sum(qty for _, qty in selected)
    half_total = total_qty / 2
    running = 0
    crossed_half = False
    result = {}
    for tid, qty in selected:
        if not crossed_half:
            result[tid] = (qty, 0)
            running += qty
            if running >= half_total:
                crossed_half = True
        else:
            result[tid] = (0, qty)

    # Guarantee an actual split whenever there's more than one table: if the
    # halfway point wasn't crossed until the very last table, everything
    # would otherwise land in Morning with nothing in Afternoon.
    if all(afternoon == 0 for _, afternoon in result.values()):
        last_tid, last_qty = selected[-1]
        result[last_tid] = (0, last_qty)

    return result


def rows_to_cutplan_dataframe(rows) -> pd.DataFrame:
    """Convert a list of row dicts (as submitted from the editable Cut Plan
    table in the browser) back into a properly-typed plan DataFrame, ready
    to be written out with write_cut_plan_workbook()."""
    int_cols = {"Sewing target per day", "Cut Plan Qty", "Diff", "Table No.", "Cut Plan Morning", "Cut Plan Afternoon"}

    def to_int(v):
        try:
            if v is None or v == "":
                return 0
            return int(float(v))
        except (TypeError, ValueError):
            return 0

    cleaned = []
    for row in rows:
        out = {}
        for col in CUT_PLAN_COLUMNS:
            val = row.get(col, "")
            out[col] = to_int(val) if col in int_cols else ("" if val is None else str(val))
        cleaned.append(out)

    return pd.DataFrame(cleaned, columns=CUT_PLAN_COLUMNS)


def _candidate_table_pool(df: pd.DataFrame) -> pd.DataFrame:
    """Shared step for build_cut_plan() and recalc_cut_plan(): drop completed
    colorway rows, then combine Qty across colorways sharing a Table ID.
    Returns one row per (Sewing Line, JobCut - Suffix, Mark Type, Table ID)."""
    work = df.copy()
    work["Qty"] = pd.to_numeric(work["Qty"], errors="coerce").fillna(0)
    work["Table ID"] = pd.to_numeric(work["Table ID"], errors="coerce")
    work["Sewing Target Per Day"] = pd.to_numeric(work["Sewing Target Per Day"], errors="coerce").fillna(0)

    status_norm = work["Status"].astype(str).str.strip().str.lower()
    work = work[status_norm != "completed"]

    group_cols = ["Sewing Line", "JobCut - Suffix", "Mark Type"]
    table_cols = group_cols + ["Table ID"]

    table_level = (
        work.groupby(table_cols, dropna=False)
        .agg(Combined_Qty=("Qty", "sum"), Colorway_Count=("Colorway", "nunique"), Target=("Sewing Target Per Day", "max"))
        .reset_index()
    )
    return table_level


MERGE_COLUMNS = ["Sewing Line", "JobCut - Suffix", "Mark Type", "Sewing target per day", "Cut Plan Qty", "Diff"]

# Which physical building each Sewing Line's cutting work happens in. Used to
# split the Cut Plan into two separate output tables/sheets: "Building 1"
# and "Building 2". A Sewing Line not listed here falls into "Unassigned"
# (kept, not dropped, so nothing silently disappears from the plan).
BUILDING_1_LINES = ["VSEW001", "VSEW002", "VSEW003", "VSEW004", "VSEW005", "VSEW007", "VSEW015", "VS02+06"]
BUILDING_2_LINES = ["VSEW006", "VSEW009", "VSEW010", "VSEW012", "VSEW013", "VSEW014", "VSEW016", "VSEW017", "VSEW018", "VS08+11"]


def get_building(sewing_line) -> str:
    """Map a Sewing Line to its building label."""
    s = str(sewing_line).strip()
    if s in BUILDING_1_LINES:
        return "Building 1"
    if s in BUILDING_2_LINES:
        return "Building 2"
    return "Unassigned"


def compute_multi_decoration_jobcuts(source_df: pd.DataFrame) -> set:
    """Find every (Sewing Line, JobCut - Suffix) that has MORE THAN ONE
    distinct Mark Type needing decoration (Decoration == "Yes" in the
    source data).

    These are the JobCuts where the planner does NOT have to plan every
    Mark Type, and does NOT have to plan enough tables to fully reach the
    Sewing Target Per Day - partial planning is a legitimate choice when a
    JobCut's decoration work is already split across multiple Mark Types.

    Returns a set of (sewing_line, jobcut_suffix) string tuples.
    """
    if "Decoration" not in source_df.columns:
        return set()

    needs_decoration = source_df[source_df["Decoration"].astype(str).str.strip().str.lower() == "yes"]
    if len(needs_decoration) == 0:
        return set()

    counts = needs_decoration.groupby(["Sewing Line", "JobCut - Suffix"])["Mark Type"].nunique()
    flagged = counts[counts > 1]
    return {(str(sl), str(jc)) for sl, jc in flagged.index}


def split_plan_by_building(plan_df: pd.DataFrame) -> "dict[str, pd.DataFrame]":
    """Split a Cut Plan DataFrame into per-building DataFrames, preserving
    row order within each. Only includes "Unassigned" if it's non-empty, so
    a fully-mapped plan never shows a stray empty section - but a Sewing
    Line that isn't in either building list is still kept, never dropped."""
    if len(plan_df) == 0:
        return {"Building 1": plan_df, "Building 2": plan_df}

    building_series = plan_df["Sewing Line"].map(get_building)
    result = {
        "Building 1": plan_df[building_series == "Building 1"].reset_index(drop=True),
        "Building 2": plan_df[building_series == "Building 2"].reset_index(drop=True),
    }
    unassigned = plan_df[building_series == "Unassigned"].reset_index(drop=True)
    if len(unassigned) > 0:
        result["Unassigned"] = unassigned
    return result


def compute_continuation_flags(rows) -> list:
    """For each row, figure out - per merge-eligible column - whether it
    repeats the row above and should therefore be visually merged with it
    (blank/merged instead of shown again).

    Nesting rule:
      - Sewing Line and JobCut - Suffix are shown/hidden TOGETHER: both only
        merge with the row above when the JobCut - Suffix is unchanged. That
        way Sewing Line re-displays at the start of every JobCut, not just
        once for the whole Sewing Line block - easier to read which JobCut
        a given block belongs to.
      - Mark Type merges with the row above only within the same JobCut.
      - Sewing target per day, Cut Plan Qty and Diff merge with the row
        above only within the same JobCut AND the same Mark Type - so they
        re-display at the start of every Mark Type sub-block, even if the
        value happens to coincide with the previous Mark Type's value.

    rows: list of dicts (same shape as CUT_PLAN_COLUMNS rows), already
    sorted by Sewing Line -> JobCut - Suffix -> Mark Type -> Table No.

    Returns a list (same length as rows) of {column: bool} - True means
    "this cell repeats the row above, merge/blank it".
    """
    flags = []
    prev = None
    for row in rows:
        row_flags = {}
        if prev is None:
            for col in MERGE_COLUMNS:
                row_flags[col] = False
        else:
            same_sewing_line = str(row.get("Sewing Line", "")) == str(prev.get("Sewing Line", ""))
            same_jobcut = same_sewing_line and str(row.get("JobCut - Suffix", "")) == str(prev.get("JobCut - Suffix", ""))
            same_mark_type = same_jobcut and str(row.get("Mark Type", "")) == str(prev.get("Mark Type", ""))

            row_flags["Sewing Line"] = same_jobcut
            row_flags["JobCut - Suffix"] = same_jobcut
            row_flags["Mark Type"] = same_mark_type
            for col in ["Sewing target per day", "Cut Plan Qty", "Diff"]:
                row_flags[col] = same_mark_type and str(row.get(col, "")) == str(prev.get(col, ""))
        flags.append(row_flags)
        prev = row
    return flags



def build_cut_plan(df: pd.DataFrame, run_datetime: Optional[datetime] = None):
    """Build the Tab 3 cut plan from Tab 1's extracted DataFrame.

    Rules:
      1. Completed tables are excluded from planning. Status is recorded
         per colorway row, so a table with mixed colorway statuses only
         drops its already-completed colorway(s) — the remaining
         (Pending / In Progress) colorway(s) still get planned.
      2. Multiple colorways can share the same Table ID within a
         (Sewing Line, JobCut - Suffix, Mark Type) group — their Qty is
         combined into one figure per table before planning.
      3. Within each (Sewing Line, JobCut - Suffix, Mark Type) group,
         tables are planned smallest Table ID first, adding tables in that
         order until the running combined Qty reaches (or first exceeds)
         that group's Sewing Target Per Day.
      4. Cut Plan Morning / Cut Plan Afternoon is decided per group, from
         the selected tables' quantities alone - see split_morning_afternoon():
         a single selected table goes entirely to Morning; with more than
         one, tables are split into a "top" (Morning) part and a "bottom"
         (Afternoon) part based on where the running quantity crosses half
         of the group's own Cut Plan Qty (the total quantity actually being
         allocated across the selected tables - not half of Sewing Target
         Per Day, which the selected tables don't always add up anywhere
         near).

    Returns (plan_df, run_info).
    """
    table_level = _candidate_table_pool(df)

    group_cols = ["Sewing Line", "JobCut - Suffix", "Mark Type"]
    # Sort candidates by Sewing Line -> JobCut - Suffix -> Mark Type -> Table ID
    # up front, so the plan naturally comes out grouped/ordered that way too
    # (matches the reference planning sheet's layout).
    table_level = table_level.sort_values(group_cols + ["Table ID"]).reset_index(drop=True)
    output_rows = []

    for keys, g in table_level.groupby(group_cols, sort=False, dropna=False):
        sewing_line, jobcut_suffix, mark_type = keys
        target = g["Target"].max()
        # Rule 3: smallest Table ID first.
        g_sorted = g.sort_values("Table ID", ascending=True).reset_index(drop=True)

        cum = 0
        selected = []
        for _, row in g_sorted.iterrows():
            if cum >= target:
                break
            cum += row["Combined_Qty"]
            table_id = int(row["Table ID"]) if pd.notna(row["Table ID"]) else ""
            selected.append((table_id, int(row["Combined_Qty"])))

        cut_plan_qty = cum
        diff = cut_plan_qty - target

        split = split_morning_afternoon(selected)

        for table_id, qty in selected:
            morning, afternoon = split[table_id]

            output_rows.append(
                {
                    "Sewing Line": sewing_line,
                    "JobCut - Suffix": jobcut_suffix,
                    "Sewing target per day": int(target),
                    "Cut Plan Qty": int(cut_plan_qty),
                    "Diff": int(diff),
                    "Mark Type": mark_type,
                    "Table No.": table_id,
                    "Cut Plan Morning": morning,
                    "Cut Plan Afternoon": afternoon,
                    "Note": "",  # left blank - purely for the user's own manual comments
                }
            )

    plan_df = pd.DataFrame(output_rows, columns=CUT_PLAN_COLUMNS)
    run_info = {
        "run_datetime": run_datetime or datetime.now(),
        "plan_date": (run_datetime or datetime.now()).date(),
    }
    return plan_df, run_info


def lookup_table_qty(source_df: pd.DataFrame, sewing_line: str, jobcut_suffix: str, mark_type, table_id: int):
    """Look up a single table's TRUE combined quantity (colorways combined,
    completed colorways excluded) straight from the source data - used when
    the user directly edits a row's Table No., so that row's quantity gets
    replaced with the real number for whichever table they typed in,
    instead of keeping the previous table's leftover value.

    Returns the quantity (int) if that exact (Sewing Line, JobCut - Suffix,
    Mark Type, Table ID) combination exists in the source data, else None.
    """
    table_level = _candidate_table_pool(source_df)
    match = table_level[
        (table_level["Sewing Line"].astype(str) == str(sewing_line))
        & (table_level["JobCut - Suffix"].astype(str) == str(jobcut_suffix))
        & (table_level["Mark Type"].astype(str) == str(mark_type))
        & (table_level["Table ID"] == table_id)
    ]
    if len(match) == 0:
        return None
    return int(match.iloc[0]["Combined_Qty"])


def recalc_cut_plan(source_df: pd.DataFrame, current_rows: list, run_info: dict):
    """Re-run table selection after the user has edited Cut Plan Qty and/or
    Sewing target per day on the (already generated) Tab 3 table.

    For each (Sewing Line, JobCut - Suffix, Mark Type) group present in
    current_rows:
      - The group's target is whatever the user's rows currently show for
        Sewing target per day (max across that group's rows).
      - Each candidate table's quantity is the user's edited Cut Plan Qty for
        that table if that table is currently on the page, otherwise the
        original combined Qty from the source data (tables the user never
        saw/touched).
      - Tables are re-selected smallest Table ID first, accumulating until
        the (possibly new) target is reached/exceeded — exactly like
        build_cut_plan(), just against the edited numbers. This can both
        ADD tables (target raised, or an edited qty lowered the running
        total) and DROP tables (target lowered, or an edited qty raised the
        running total past it sooner).
      - Manual comments already typed into Note are preserved for tables
        that remain selected.
      - Rows for tables the user manually added (a Table No. that doesn't
        exist in the source data for that group) are always kept as-is
        (their own Morning/Afternoon split is never recomputed), and their
        Cut Plan Qty is added into the group's total/Diff.
      - Cut Plan Morning / Cut Plan Afternoon for the REAL (non-manual)
        selected tables is always freshly recomputed for the whole group
        using split_morning_afternoon() - the same deterministic rule
        build_cut_plan() uses - since the split boundary (half the target)
        can shift whenever the target or a table's quantity changes.

    Returns a fresh plan_df (same columns/shape as build_cut_plan's output).
    """
    table_level = _candidate_table_pool(source_df)

    # Index candidates for fast lookup: (Sewing Line, JobCut - Suffix, Mark Type) -> DataFrame
    group_cols = ["Sewing Line", "JobCut - Suffix", "Mark Type"]
    candidates_by_group = {
        keys: g.reset_index(drop=True) for keys, g in table_level.groupby(group_cols, sort=False, dropna=False)
    }

    def to_num(v, default=0.0):
        try:
            if v is None or v == "":
                return default
            return float(v)
        except (TypeError, ValueError):
            return default

    # Organize the user's current rows by group.
    rows_by_group = {}
    for row in current_rows:
        key = (row.get("Sewing Line", ""), row.get("JobCut - Suffix", ""), row.get("Mark Type", ""))
        rows_by_group.setdefault(key, []).append(row)

    output_rows = []

    for key, user_rows in rows_by_group.items():
        sewing_line, jobcut_suffix, mark_type = key
        target = max((to_num(r.get("Sewing target per day")) for r in user_rows), default=0.0)

        candidates = candidates_by_group.get(key)

        # Build override map: Table ID (int) -> user's current Cut Plan Qty for
        # that row, and preserve any Note text already on that row. Also collect
        # rows whose Table No. isn't a real candidate for this group — manually
        # added extras. A row only counts as an override if it has an EXPLICIT
        # Morning/Afternoon value; if both are blank (e.g. the user just
        # changed which Table No. this row points to, so its old
        # Morning/Afternoon no longer means anything), it's treated as having
        # no known quantity yet, and falls back to that table's real quantity
        # from the source data instead of incorrectly reusing a stale value.
        qty_override = {}
        note_by_table = {}
        manual_rows = []
        for r in user_rows:
            table_no_raw = r.get("Table No.", "")
            try:
                table_id = int(float(table_no_raw))
            except (TypeError, ValueError):
                table_id = None

            is_real_candidate = (
                candidates is not None and table_id is not None
                and (candidates["Table ID"] == table_id).any()
            )

            morning_raw = r.get("Cut Plan Morning", "")
            afternoon_raw = r.get("Cut Plan Afternoon", "")
            has_known_qty = not (
                (morning_raw is None or str(morning_raw).strip() == "")
                and (afternoon_raw is None or str(afternoon_raw).strip() == "")
            )

            if is_real_candidate:
                note_by_table[table_id] = r.get("Note", "")
                if has_known_qty:
                    qty_override[table_id] = to_num(morning_raw) + to_num(afternoon_raw)
            else:
                manual_rows.append(r)

        selected_real_rows = []
        real_total = 0

        if candidates is not None and len(candidates) > 0:
            g_sorted = candidates.sort_values("Table ID", ascending=True).reset_index(drop=True)
            cum = 0
            for _, crow in g_sorted.iterrows():
                if cum >= target:
                    break
                table_id = int(crow["Table ID"]) if pd.notna(crow["Table ID"]) else None
                qty = qty_override.get(table_id, crow["Combined_Qty"])
                cum += qty
                selected_real_rows.append((table_id, int(qty)))
            real_total = cum

        manual_total = sum(to_num(r.get("Cut Plan Morning")) + to_num(r.get("Cut Plan Afternoon")) for r in manual_rows)
        cut_plan_qty = real_total + manual_total
        diff = cut_plan_qty - target

        split = split_morning_afternoon(selected_real_rows)

        for table_id, qty in selected_real_rows:
            morning, afternoon = split[table_id]
            note = note_by_table.get(table_id, "")

            output_rows.append(
                {
                    "Sewing Line": sewing_line,
                    "JobCut - Suffix": jobcut_suffix,
                    "Sewing target per day": int(target),
                    "Cut Plan Qty": int(cut_plan_qty),
                    "Diff": int(diff),
                    "Mark Type": mark_type,
                    "Table No.": table_id,
                    "Cut Plan Morning": int(morning),
                    "Cut Plan Afternoon": int(afternoon),
                    "Note": note,
                }
            )

        for r in manual_rows:
            output_rows.append(
                {
                    "Sewing Line": sewing_line,
                    "JobCut - Suffix": jobcut_suffix,
                    "Sewing target per day": int(target),
                    "Cut Plan Qty": int(cut_plan_qty),
                    "Diff": int(diff),
                    "Mark Type": mark_type,
                    "Table No.": r.get("Table No.", ""),
                    "Cut Plan Morning": int(to_num(r.get("Cut Plan Morning"))),
                    "Cut Plan Afternoon": int(to_num(r.get("Cut Plan Afternoon"))),
                    "Note": r.get("Note", ""),
                }
            )

    result_df = pd.DataFrame(output_rows, columns=CUT_PLAN_COLUMNS)
    if len(result_df) > 0:
        # Sort by Sewing Line -> JobCut - Suffix -> Mark Type, stable so each
        # group's already-correct internal Table No. order is preserved.
        result_df = result_df.sort_values(group_cols, kind="stable").reset_index(drop=True)
    return result_df


CUT_PLAN_ASSUMPTIONS_TEXT = [
    "CUT PLANNING MODEL v2 — TAB 3: CUT PLAN",
    "",
    "1. INPUT",
    "Built directly from Tab 1's extracted Buffer Cutting Order Form data (already filtered to",
    "Sewing Lines starting with \"VS\"). Tab 2 (WIP Upload) is NOT used yet — see the note on the",
    "Cut Plan tab.",
    "",
    "2. COMPLETED TABLES ARE EXCLUDED",
    "Status is recorded per colorway row. Any row with Status = \"Completed\" is dropped before",
    "planning. A table whose colorways have mixed status (e.g. one colorway Completed, another",
    "still Pending) is NOT excluded entirely — only its completed colorway's quantity is dropped;",
    "the remaining colorway(s) are still planned.",
    "",
    "3. COMBINING COLORWAYS",
    "A single Table ID can have more than one colorway. Their Qty is summed into one combined",
    "figure per table before planning.",
    "",
    "3b. NOTE COLUMN",
    "Left blank by the model. It's only filled in if the user types something into it manually",
    "on the editable Cut Plan tab (e.g. a manual adjustment reason) before saving.",
    "",
    "4. TABLE SELECTION RULE",
    "Within each (Sewing Line, JobCut - Suffix, Mark Type) group, tables are planned SMALLEST",
    "Table ID first. Tables are added in that order until the running combined Qty reaches (or",
    "first exceeds) the group's Sewing Target Per Day. Only the selected tables appear in the",
    "output.",
    "",
    "5. Cut Plan Qty / Diff",
    "Cut Plan Qty = sum of combined Qty across all tables selected for that group (repeated on",
    "each row of the group for readability). Diff = Cut Plan Qty - Sewing Target Per Day.",
    "",
    "6. MORNING / AFTERNOON SPLIT",
    "Decided per group, from the selected tables' quantities alone - no wall-clock time or user",
    "choice is involved:",
    "  - If a group has exactly ONE selected table, its whole quantity goes to Cut Plan Morning.",
    "  - If a group has MORE THAN ONE selected table, they're split into a \"top\" part (Morning)",
    "    and a \"bottom\" part (Afternoon): walking the tables in Table ID order, each table stays",
    "    in the top/Morning part and its quantity accumulates, until the running total reaches",
    "    half of the group's own CUT PLAN QTY (the total quantity actually being allocated across",
    "    the selected tables - NOT half of Sewing Target Per Day). The table whose addition",
    "    crosses that halfway point is still counted in the top/Morning part; every table AFTER",
    "    that point goes to the bottom/Afternoon part. When the total can't split into two exactly",
    "    equal halves, the larger portion lands in Morning and the smaller one in Afternoon.",
    "See split_morning_afternoon() in cutplan2/model.py for the exact logic.",
]


def _write_cut_plan_sheet(ws, plan_df: pd.DataFrame, title_text: str, edited_at: Optional[datetime] = None) -> None:
    """Write one Cut Plan table (one building's worth of rows) into the given
    worksheet: title row(s), header, data with merged repeated values,
    grouping borders, column widths, freeze panes and auto-filter."""
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=10)
    cell_font = Font(name=FONT_NAME, size=10)
    note_font = Font(name=FONT_NAME, size=10, italic=True, color=WARN_FONT_COLOR)
    group_fill_a = PatternFill("solid", fgColor="DCE6F1")
    group_fill_b = PatternFill("solid", fgColor="FFFFFF")
    thin = Side(style="thin", color="B7B7B7")
    thick_top = Side(style="thick", color="1F4E78")
    medium_top = Side(style="medium", color="7C99B8")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    border_new_sewing_line = Border(left=thin, right=thin, top=thick_top, bottom=thin)
    border_new_jobcut = Border(left=thin, right=thin, top=medium_top, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")
    title_font = Font(name=FONT_NAME, size=10, bold=True, color=TITLE_FONT_COLOR)
    edited_font = Font(name=FONT_NAME, size=10, bold=True, color=WARN_FONT_COLOR)

    ws.cell(row=1, column=1, value=title_text).font = title_font
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(CUT_PLAN_COLUMNS))

    next_row = 2
    if edited_at is not None:
        edited_text = f"Manually edited on {edited_at.strftime('%Y-%m-%d %H:%M')} — values below may differ from the automatic plan"
        ws.cell(row=2, column=1, value=edited_text).font = edited_font
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(CUT_PLAN_COLUMNS))
        next_row = 3

    header_row = next_row
    headers = list(plan_df.columns)
    for j, h in enumerate(headers, start=1):
        c = ws.cell(row=header_row, column=j, value=h)
        c.font = header_font
        c.fill = header_fill
        c.alignment = center
        c.border = border

    if len(plan_df) > 0:
        rows_as_dicts = plan_df.to_dict(orient="records")
        continuation_flags = compute_continuation_flags(rows_as_dicts)

        group_key_series = (
            plan_df["Sewing Line"].astype(str)
            + "|"
            + plan_df["JobCut - Suffix"].astype(str)
            + "|"
            + plan_df["Mark Type"].astype(str)
        )
        sewing_line_series = plan_df["Sewing Line"].astype(str)
        jobcut_series = plan_df["Sewing Line"].astype(str) + "|" + plan_df["JobCut - Suffix"].astype(str)
        current_key = None
        fill_toggle = False
        prev_sewing_line = None
        prev_jobcut = None

        col_index = {h: j for j, h in enumerate(headers, start=1)}

        for i, row in enumerate(plan_df.itertuples(index=False), start=header_row + 1):
            idx = i - (header_row + 1)
            key = group_key_series.iloc[idx]
            if key != current_key:
                fill_toggle = not fill_toggle
                current_key = key
            fill = group_fill_a if fill_toggle else group_fill_b

            sewing_line_val = sewing_line_series.iloc[idx]
            jobcut_val = jobcut_series.iloc[idx]
            if sewing_line_val != prev_sewing_line and prev_sewing_line is not None:
                row_border = border_new_sewing_line
            elif jobcut_val != prev_jobcut and prev_jobcut is not None:
                row_border = border_new_jobcut
            else:
                row_border = border
            prev_sewing_line = sewing_line_val
            prev_jobcut = jobcut_val

            row_flags = continuation_flags[idx]
            values = list(row)
            for j, val in enumerate(values, start=1):
                col_name = headers[j - 1]
                is_continuation = col_name in MERGE_COLUMNS and row_flags.get(col_name, False)
                c = ws.cell(row=i, column=j, value=(None if is_continuation else val))
                c.border = row_border
                c.fill = fill
                if col_name == "Note":
                    c.font = note_font
                    c.alignment = left
                else:
                    c.font = cell_font
                    c.alignment = center

        # ---- Apply true cell merges for repeated values in MERGE_COLUMNS ----
        for col_name in MERGE_COLUMNS:
            j = col_index[col_name]
            run_start = None
            for idx, flags in enumerate(continuation_flags):
                excel_row = header_row + 1 + idx
                if flags.get(col_name, False):
                    continue
                if run_start is not None:
                    run_end = header_row + idx
                    if run_end > run_start:
                        ws.merge_cells(start_row=run_start, start_column=j, end_row=run_end, end_column=j)
                run_start = excel_row
            if run_start is not None:
                run_end = header_row + len(continuation_flags)
                if run_end > run_start:
                    ws.merge_cells(start_row=run_start, start_column=j, end_row=run_end, end_column=j)

    widths = [12, 16, 18, 12, 8, 10, 10, 15, 16, 46]
    for j, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(j)].width = w

    ws.freeze_panes = f"A{header_row + 1}"
    last_row = max(len(plan_df) + header_row, header_row)
    ws.auto_filter.ref = f"A{header_row}:{get_column_letter(len(headers))}{last_row}"


def write_single_building_workbook(
    building_df: pd.DataFrame,
    run_info: dict,
    building_label: str,
    downloaded_date,
    output_path: str,
    edited_at: Optional[datetime] = None,
) -> None:
    """Write ONE standalone workbook for a single building: its Cut Plan
    sheet (titled "<building_label> of <downloaded_date>") plus a Logic &
    Assumptions sheet, so each building's file is self-contained.

    downloaded_date should be the date the file is actually being generated
    for download (i.e. "today"), not necessarily the plan's target date -
    pass a date object.
    """
    wb = Workbook()
    wb.remove(wb.active)

    generated_at = run_info["run_datetime"].strftime("%Y-%m-%d %H:%M")
    ws = wb.create_sheet("Cut Plan")
    title_text = f"{building_label} of {downloaded_date.isoformat()}   |   Generated: {generated_at}"
    _write_cut_plan_sheet(ws, building_df, title_text, edited_at=edited_at)

    ws2 = wb.create_sheet("Logic & Assumptions")
    ws2.column_dimensions["A"].width = 100
    for i, text in enumerate(CUT_PLAN_ASSUMPTIONS_TEXT, start=1):
        c = ws2.cell(row=i, column=1, value=text)
        c.font = Font(name=FONT_NAME, size=10, bold=(i == 1))
        c.alignment = Alignment(wrap_text=True, vertical="top")

    wb.save(output_path)


def write_cut_plan_workbook(
    plan_df: pd.DataFrame,
    run_info: dict,
    output_path: str,
    edited_at: Optional[datetime] = None,
) -> None:
    """Write the formatted Tab 3 output workbook: one sheet per building
    ("Building 1", "Building 2", and "Unassigned" if any Sewing Line doesn't
    map to either) plus a Logic & Assumptions sheet.

    Pass edited_at when this call is regenerating the file after the user
    manually adjusted the plan table in the browser - it's noted in each
    sheet's title row so it's clear the numbers no longer come straight from
    the automatic selection rule.
    """
    wb = Workbook()
    wb.remove(wb.active)  # replaced by the per-building sheets below

    generated_at = run_info["run_datetime"].strftime("%Y-%m-%d %H:%M")
    plan_date_str = run_info["plan_date"].isoformat()

    by_building = split_plan_by_building(plan_df)
    # Keep a stable, sensible sheet order even when a building is empty.
    ordered_keys = [k for k in ["Building 1", "Building 2", "Unassigned"] if k in by_building]

    for key in ordered_keys:
        building_df = by_building[key]
        ws = wb.create_sheet(key)
        title_text = f"{key} of {plan_date_str}   |   Generated: {generated_at}"
        _write_cut_plan_sheet(ws, building_df, title_text, edited_at=edited_at)

    # ---- Logic & Assumptions ----
    ws2 = wb.create_sheet("Logic & Assumptions")
    ws2.column_dimensions["A"].width = 100
    for i, text in enumerate(CUT_PLAN_ASSUMPTIONS_TEXT, start=1):
        c = ws2.cell(row=i, column=1, value=text)
        c.font = Font(name=FONT_NAME, size=10, bold=(i == 1))
        c.alignment = Alignment(wrap_text=True, vertical="top")

    wb.save(output_path)
