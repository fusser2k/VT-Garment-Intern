"""
Cut Planning Model v2 — local web server
==========================================
Three tabs:
  Tab 1: Buffer Cutting Order Form  — upload + extract into the 17-field schema
  Tab 2: WIP Upload                 — upload a Work-in-Process file (no fixed
                                       template yet, so this just stores/previews it)
  Tab 3: Cut Plan                   — plans which (incomplete) tables to cut
                                       first, smallest Table ID first, using
                                       Tab 1's extracted data. NOT wired to
                                       Tab 2 yet, since the WIP file's
                                       template isn't known.

Run locally:
    pip install -r requirements.txt
    python app.py

Then open http://127.0.0.1:5001 (see README.md for local-network access).
"""

import os
import traceback
import uuid

from flask import Flask, render_template, request, send_file, flash, redirect, url_for, session, jsonify

from cutplan2 import (
    load_input,
    write_extracted_workbook,
    OUTPUT_COLUMNS,
    now_th,
    load_wip_raw,
    load_wip,
    write_wip_workbook,
    build_cut_plan,
    compute_wip_session_targets,
    compute_wip_jobcut_restrictions,
    recalc_cut_plan,
    lookup_table_qty,
    write_cut_plan_workbook,
    write_single_building_workbook,
    CUT_PLAN_COLUMNS,
    MERGE_COLUMNS,
    compute_continuation_flags,
    rows_to_cutplan_dataframe,
    get_building,
    split_plan_by_building,
    compute_multi_decoration_jobcuts,
    BUILDING_1_LINES,
    BUILDING_2_LINES,
    translate,
    column_label,
    js_translations,
    DEFAULT_LANG,
    SUPPORTED_LANGS,
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")
SAMPLE_FILE = os.path.join(BASE_DIR, "sample_data", "BufferCuttingOrderForm_2026-08-21.xlsx")
WIP_SAMPLE_FILE = os.path.join(BASE_DIR, "sample_data", "WIP_18-08-2569.xlsx")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

app = Flask(__name__)
app.secret_key = "cut-planning-model-v2-local-dev"  # only used for session cookie + flash; change if deploying beyond localhost
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # 25 MB upload limit

# In-memory store, keyed by a per-browser session id: holds each tab's last
# result so switching tabs (or reloading) doesn't lose what's already there.
# This is a local single-user tool, not built for a hosted multi-user service —
# restarting the server clears everything (just re-upload).
SESSIONS = {}

WIP_PREVIEW_ROW_LIMIT = 200


def _get_lang() -> str:
    lang = session.get("lang", DEFAULT_LANG)
    return lang if lang in SUPPORTED_LANGS else DEFAULT_LANG


# Jinja globals: {{ t('key') }} and {{ col_label(col) }}, always bound to the
# CURRENT request's language (read from the session on every call, so a
# language switch takes effect immediately without needing a fresh login).
app.jinja_env.globals["t"] = lambda key: translate(key, _get_lang())
app.jinja_env.globals["col_label"] = lambda col: column_label(col, _get_lang())
app.jinja_env.filters["col_label"] = lambda col: column_label(col, _get_lang())


@app.route("/set-language/<lang>", methods=["GET"])
def set_language(lang):
    if lang in SUPPORTED_LANGS:
        session["lang"] = lang
    try:
        active_tab = int(request.args.get("tab", 1))
    except ValueError:
        active_tab = 1
    return redirect(url_for("index", tab=active_tab))


def _get_sid():
    sid = session.get("sid")
    if not sid:
        sid = uuid.uuid4().hex[:12]
        session["sid"] = sid
    SESSIONS.setdefault(sid, {"extraction": None, "wip": None, "cutplan": None})
    return sid


def _cutplan_for_display(cutplan, extraction):
    """Return a copy of the cutplan session state, split into per-building
    row lists (each annotated with '_cont' continuation flags computed
    WITHIN that building only, and '_multi_decoration' marking JobCuts that
    have more than one Mark Type needing decoration), used purely for
    rendering Tab 3's two tables - never stored back into SESSIONS."""
    if not cutplan:
        return cutplan
    rows = cutplan["rows"]

    flagged_jobcuts = set()
    if extraction and "df" in extraction:
        try:
            flagged_jobcuts = compute_multi_decoration_jobcuts(extraction["df"])
        except Exception:
            flagged_jobcuts = set()

    buildings = {"Building 1": [], "Building 2": [], "Unassigned": []}
    for row in rows:
        buildings[get_building(row.get("Sewing Line", ""))].append(row)

    display_buildings = {}
    for key, building_rows in buildings.items():
        flags = compute_continuation_flags(building_rows)
        display_rows = []
        for row, row_flags in zip(building_rows, flags):
            display_row = dict(row)
            display_row["_cont"] = row_flags
            display_row["_multi_decoration"] = (
                str(row.get("Sewing Line", "")), str(row.get("JobCut - Suffix", ""))
            ) in flagged_jobcuts
            display_rows.append(display_row)
        display_buildings[key] = {"rows": display_rows, "row_count": len(display_rows)}

    display_cutplan = dict(cutplan)
    display_cutplan["buildings"] = display_buildings
    display_cutplan["multi_decoration_keys"] = [f"{sl}|{jc}" for sl, jc in flagged_jobcuts]
    return display_cutplan


def _render(active_tab: int):
    sid = _get_sid()
    state = SESSIONS[sid]
    lang = _get_lang()
    return render_template(
        "index.html",
        now=now_th(),
        active_tab=active_tab,
        extraction=state["extraction"],
        wip=state["wip"],
        cutplan=_cutplan_for_display(state["cutplan"], state["extraction"]),
        merge_columns=MERGE_COLUMNS,
        OUTPUT_COLUMNS=OUTPUT_COLUMNS,
        building_1_lines=BUILDING_1_LINES,
        building_2_lines=BUILDING_2_LINES,
        current_lang=lang,
        supported_langs=SUPPORTED_LANGS,
        js_t=js_translations(lang),
    )


@app.route("/", methods=["GET"])
def index():
    try:
        active_tab = int(request.args.get("tab", 1))
    except ValueError:
        active_tab = 1
    return _render(active_tab=active_tab)


# ---------------------------------------------------------------------------
# Tab 1: Buffer Cutting Order Form
# ---------------------------------------------------------------------------
@app.route("/extract", methods=["POST"])
def extract():
    sid = _get_sid()
    use_sample = request.form.get("use_sample") == "1"
    file = request.files.get("input_file")

    if use_sample:
        input_path = SAMPLE_FILE
    elif file and file.filename:
        job_id = uuid.uuid4().hex[:8]
        input_path = os.path.join(UPLOAD_DIR, f"{job_id}_{file.filename}")
        file.save(input_path)
    else:
        flash(translate("please_choose_input_file", _get_lang()))
        return redirect(url_for("index", tab=1))

    run_datetime = now_th()

    try:
        df, missing_columns, filtered_out_count = load_input(input_path)
        job_id = uuid.uuid4().hex[:8]
        output_path = os.path.join(OUTPUT_DIR, f"Extracted_Data_{job_id}.xlsx")
        write_extracted_workbook(
            df, missing_columns, output_path,
            generated_at=run_datetime, filtered_out_count=filtered_out_count,
        )
    except Exception as e:
        flash(f"Error while processing the Buffer Cutting Order Form file: {e}")
        traceback.print_exc()
        return redirect(url_for("index", tab=1))

    rows = df.where(df.notna(), "").to_dict(orient="records")

    SESSIONS[sid]["extraction"] = {
        "now": run_datetime,
        "rows": rows,
        "columns": OUTPUT_COLUMNS,
        "missing_columns": missing_columns,
        "filtered_out_count": filtered_out_count,
        "download_filename": os.path.basename(output_path),
        "row_count": len(rows),
        "df": df,  # full extracted+filtered DataFrame, reused by Tab 3's cut plan
    }
    return _render(active_tab=1)


@app.route("/download/<filename>", methods=["GET"])
def download(filename):
    safe_path = os.path.join(OUTPUT_DIR, os.path.basename(filename))
    if not os.path.exists(safe_path):
        flash("That output file is no longer available — please extract again.")
        return redirect(url_for("index", tab=1))
    download_name = f"Extracted Data of {now_th().date().isoformat()}.xlsx"
    return send_file(safe_path, as_attachment=True, download_name=download_name)


# ---------------------------------------------------------------------------
# Tab 2: WIP Upload
# ---------------------------------------------------------------------------
@app.route("/wip-upload", methods=["POST"])
def wip_upload():
    sid = _get_sid()
    use_sample = request.form.get("use_sample") == "1"
    file = request.files.get("wip_file")

    if use_sample:
        saved_path = WIP_SAMPLE_FILE
        display_filename = os.path.basename(WIP_SAMPLE_FILE)
    elif file and file.filename:
        upload_job_id = uuid.uuid4().hex[:8]
        saved_path = os.path.join(UPLOAD_DIR, f"wip_{upload_job_id}_{file.filename}")
        file.save(saved_path)
        display_filename = file.filename
    else:
        flash(translate("please_choose_wip_file", _get_lang()))
        return redirect(url_for("index", tab=2))

    job_id = uuid.uuid4().hex[:8]

    structured = True
    try:
        df = load_wip(saved_path)
        if len(df) == 0:
            # Template not recognized (header marker not found) - fall back
            # to a raw, no-schema preview rather than showing nothing.
            structured = False
            df = load_wip_raw(saved_path)
    except Exception as e:
        flash(f"Error while reading the WIP file: {e}")
        traceback.print_exc()
        return redirect(url_for("index", tab=2))

    download_filename = None
    if structured:
        try:
            output_path = os.path.join(OUTPUT_DIR, f"WIP_Extracted_{job_id}.xlsx")
            write_wip_workbook(df, output_path)
            download_filename = os.path.basename(output_path)
        except Exception:
            traceback.print_exc()  # extraction still succeeded; just skip offering a download

    preview_df = df.head(WIP_PREVIEW_ROW_LIMIT)
    rows = preview_df.where(preview_df.notna(), "").to_dict(orient="records")

    # Compute the Sewing Line -> (Target Morning, Target Afternoon, Target
    # OT) mapping right away, against the FULL df (not the possibly-
    # truncated preview) - see compute_wip_session_targets() in
    # cutplan2/model.py. Tab 3 requires this to generate a plan at all
    # (its planning logic is now built around these three per-session
    # targets), so it's computed once here rather than on every /cut-plan
    # request.
    session_targets = compute_wip_session_targets(df) if structured else {}
    # Same idea for which JobCuts each Sewing Line's own Detail Morning /
    # Detail Afternoon / Detail OT cells actually mention - see
    # compute_wip_jobcut_restrictions(). Tab 3 only plans a JobCut a line's
    # own Detail cells list, when this restriction is available.
    jobcut_restrictions = compute_wip_jobcut_restrictions(df) if structured else {}

    SESSIONS[sid]["wip"] = {
        "now": now_th(),
        "filename": display_filename,
        "columns": list(df.columns),
        "rows": rows,
        "row_count": len(df),
        "preview_row_count": len(rows),
        "truncated": len(df) > WIP_PREVIEW_ROW_LIMIT,
        "structured": structured,
        "download_filename": download_filename,
        "session_targets": session_targets,
        "jobcut_restrictions": jobcut_restrictions,
    }
    return _render(active_tab=2)


@app.route("/download-wip/<filename>", methods=["GET"])
def download_wip(filename):
    safe_path = os.path.join(OUTPUT_DIR, os.path.basename(filename))
    if not os.path.exists(safe_path):
        flash("That output file is no longer available — please upload again.")
        return redirect(url_for("index", tab=2))
    return send_file(safe_path, as_attachment=True, download_name="WIP_Extracted.xlsx")


# ---------------------------------------------------------------------------
# Tab 3: Cut Plan
# ---------------------------------------------------------------------------
@app.route("/cut-plan", methods=["POST"])
def cut_plan():
    sid = _get_sid()
    extraction = SESSIONS[sid]["extraction"]

    if not extraction or "df" not in extraction:
        flash(translate("run_tab1_first", _get_lang()))
        return redirect(url_for("index", tab=3))

    # Tab 3's planning is now built entirely around Tab 2's per-Sewing-Line
    # Morning/Afternoon/OT targets (see compute_wip_session_targets()) -
    # unlike before, Tab 2 is now a REQUIRED input, not an optional
    # override on top of Tab 1 alone.
    wip_session = SESSIONS[sid].get("wip")
    if not wip_session or not wip_session.get("structured"):
        flash(translate("run_tab2_first", _get_lang()))
        return redirect(url_for("index", tab=3))
    wip_session_targets = wip_session.get("session_targets") or {}
    wip_jobcut_restrictions = wip_session.get("jobcut_restrictions") or {}

    run_datetime = now_th()

    try:
        plan_df, run_info = build_cut_plan(
            extraction["df"], wip_session_targets, run_datetime, wip_jobcut_restrictions=wip_jobcut_restrictions
        )
        job_id = uuid.uuid4().hex[:8]
    except Exception as e:
        flash(f"Error while building the cut plan: {e}")
        traceback.print_exc()
        return redirect(url_for("index", tab=3))

    rows = plan_df.to_dict(orient="records")

    SESSIONS[sid]["cutplan"] = {
        "run_info": run_info,
        "rows": rows,
        "columns": CUT_PLAN_COLUMNS,
        "row_count": len(rows),
        "job_id": job_id,
        "building_paths": {},  # populated / refreshed each time Save is clicked
    }
    return _render(active_tab=3)


@app.route("/recalc-cut-plan/<job_id>", methods=["POST"])
def recalc_cut_plan_route(job_id):
    """Re-run table selection using the currently edited Cut Plan table
    (triggered when the user changes Cut Plan Qty, Cut Plan Morning/Afternoon,
    or Sewing target per day). Returns fresh JSON rows — no file is written
    here; that only happens on Save."""
    sid = _get_sid()
    extraction = SESSIONS[sid]["extraction"]
    cutplan = SESSIONS[sid]["cutplan"]

    if not extraction or "df" not in extraction:
        return jsonify({"error": "Tab 1 data is no longer available in this session."}), 410
    if not cutplan or cutplan.get("job_id") != job_id:
        return jsonify({"error": "This cut plan is no longer available in this session."}), 410

    payload = request.get_json(silent=True) or {}
    current_rows = payload.get("rows", [])

    try:
        fresh_plan_df = recalc_cut_plan(extraction["df"], current_rows, cutplan["run_info"])
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": f"Could not recalculate the plan: {e}"}), 400

    fresh_rows = fresh_plan_df.to_dict(orient="records")
    SESSIONS[sid]["cutplan"]["rows"] = fresh_rows
    SESSIONS[sid]["cutplan"]["row_count"] = len(fresh_rows)

    return jsonify({"rows": fresh_rows, "row_count": len(fresh_rows)})


@app.route("/lookup-table-qty/<job_id>", methods=["POST"])
def lookup_table_qty_route(job_id):
    """Look up a single table's real quantity from the source (Tab 1) data -
    used when the user edits a row's Table No. directly, so that row's
    quantity gets replaced with the correct number for the table they typed,
    rather than keeping the previous table's leftover quantity. Does NOT
    re-run the table-selection algorithm - this is a plain lookup."""
    sid = _get_sid()
    extraction = SESSIONS[sid]["extraction"]
    cutplan = SESSIONS[sid]["cutplan"]

    if not extraction or "df" not in extraction:
        return jsonify({"error": "Tab 1 data is no longer available in this session."}), 410
    if not cutplan or cutplan.get("job_id") != job_id:
        return jsonify({"error": "This cut plan is no longer available in this session."}), 410

    payload = request.get_json(silent=True) or {}
    sewing_line = payload.get("sewing_line", "")
    jobcut_suffix = payload.get("jobcut_suffix", "")
    mark_type = payload.get("mark_type", "")
    table_no_raw = payload.get("table_no", "")

    try:
        table_id = int(float(table_no_raw))
    except (TypeError, ValueError):
        return jsonify({"error": "Table No. must be a number."}), 400

    try:
        qty = lookup_table_qty(extraction["df"], sewing_line, jobcut_suffix, mark_type, table_id)
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": f"Could not look up that table: {e}"}), 400

    if qty is None:
        return jsonify({
            "error": f"Table {table_id} was not found for {sewing_line} / {jobcut_suffix} / Mark Type {mark_type} "
                     f"(it may not exist, or may already be fully Completed)."
        }), 404

    return jsonify({"qty": qty})


@app.route("/update-cut-plan/<job_id>", methods=["POST"])
def update_cut_plan(job_id):
    """Regenerate the Cut Plan as SEPARATE workbooks - one per building
    ("Building 1 of <date>.xlsx", "Building 2 of <date>.xlsx", and
    "Unassigned of <date>.xlsx" if that section has any rows) - from the
    (possibly hand-edited) table posted from Tab 3. Returns JSON describing
    each file; the browser downloads them individually via
    /download-cutplan-file/<job_id>/<building_key>."""
    sid = _get_sid()
    cutplan = SESSIONS[sid]["cutplan"]

    if not cutplan or cutplan.get("job_id") != job_id:
        return jsonify({"error": "This cut plan is no longer available in this session. Please generate it again."}), 410

    payload = request.get_json(silent=True) or {}
    rows = payload.get("rows", [])

    try:
        edited_plan_df = rows_to_cutplan_dataframe(rows)
        edited_at = now_th()
        by_building = split_plan_by_building(edited_plan_df)

        building_paths = {}
        files_info = []
        for key in ["Building 1", "Building 2", "Unassigned"]:
            building_df = by_building.get(key)
            if building_df is None or len(building_df) == 0:
                continue
            slug = key.lower().replace(" ", "-")
            output_path = os.path.join(OUTPUT_DIR, f"CutPlan_{job_id}_{slug}.xlsx")
            write_single_building_workbook(
                building_df, cutplan["run_info"], key, edited_at.date(), output_path, edited_at=edited_at
            )
            building_paths[key] = output_path
            files_info.append({
                "building": key,
                "url": url_for("download_cutplan_file", job_id=job_id, building_key=slug),
                "filename": f"{key} of {edited_at.date().isoformat()}.xlsx",
            })

        SESSIONS[sid]["cutplan"]["rows"] = edited_plan_df.to_dict(orient="records")
        SESSIONS[sid]["cutplan"]["row_count"] = len(edited_plan_df)
        SESSIONS[sid]["cutplan"]["building_paths"] = building_paths
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": f"Could not save edits: {e}"}), 400

    return jsonify({"files": files_info})


@app.route("/download-cutplan-file/<job_id>/<building_key>", methods=["GET"])
def download_cutplan_file(job_id, building_key):
    """Serve one building's saved Cut Plan file. The download name always
    reflects TODAY's date (the moment of download), e.g.
    'Building 1 of 2026-08-18.xlsx' - even if the file was saved earlier."""
    sid = _get_sid()
    cutplan = SESSIONS[sid]["cutplan"]

    if not cutplan or cutplan.get("job_id") != job_id:
        flash("This cut plan is no longer available in this session. Please generate it again.")
        return redirect(url_for("index", tab=3))

    key_by_slug = {k.lower().replace(" ", "-"): k for k in ["Building 1", "Building 2", "Unassigned"]}
    building_key_full = key_by_slug.get(building_key)
    output_path = (cutplan.get("building_paths") or {}).get(building_key_full) if building_key_full else None

    if not output_path or not os.path.exists(output_path):
        flash("That file is no longer available — please save again.")
        return redirect(url_for("index", tab=3))

    download_name = f"{building_key_full} of {now_th().date().isoformat()}.xlsx"
    return send_file(output_path, as_attachment=True, download_name=download_name)


if __name__ == "__main__":
    # Railway (and most hosting platforms) assign their own port via the
    # PORT environment variable, and require the app to listen on 0.0.0.0
    # (not 127.0.0.1, which only accepts connections from the same
    # machine - that's why a hardcoded 127.0.0.1 causes "Application
    # failed to respond" on Railway). Locally this still defaults to port
    # 5001 exactly as before if PORT isn't set.
    port = int(os.environ.get("PORT", 5001))
    debug_mode = os.environ.get("FLASK_DEBUG", "1") == "1"
    app.run(host="0.0.0.0", port=port, debug=debug_mode)
