# Cut Planning Model v2

A **new, separate** project from the original `cut_plan_model` — a fresh
build for cut planning. It currently has three tabs, all on one page:

- **Tab 1: Buffer Cutting Order Form** — upload the cutting order Excel file
  and it's extracted into a fixed 14-field schema.
- **Tab 2: WIP Upload** — upload a Work-in-Process (WIP) buffer report and
  it's extracted into a fixed set of per-sewing-line fields (targets,
  actuals, WIP quantities, lead times, shortfall reasons). Falls back to a
  raw, no-schema preview for a file that doesn't match the known template.
- **Tab 3: Cut Plan** — plans which cutting tables to work on first, using
  Tab 1's extracted data. **Not wired to Tab 2 yet** — it's being built
  independently for now. Once it's ready to factor in WIP levels, this tab
  can be wired in.

Switching tabs is instant (client-side) and doesn't lose what's on the other
tabs — upload something on Tab 1, switch to Tab 2, upload something there,
and Tab 1's results are still exactly as you left them.

## Tab 1: the 14 extracted fields

```
Sewing Line | JobCut - Suffix | Table ID | Colorway | Mark Type | Layer |
Qty | Qty Complete | Difference | % Complete | Status |
Sewing Target Per Day | Table No. (Mark Type 101) | Decoration
```

**Filter applied automatically:** only rows whose **Sewing Line** starts with
`VS` are kept (e.g. `VSEW012`, `VS02+06`). Anything else (e.g. sample/test
rows like `SAMPL02`) is dropped before the data is shown or saved. The
results banner on the page — and the title row of the downloaded workbook —
always states how many rows were filtered out this way.

The bundled sample file (`BufferCuttingOrderForm_2026-08-17.xlsx`) already
uses these exact header names, so it extracts with zero missing fields. If
you point the model at a file that uses different header text, or is missing
some of these columns entirely, those fields are simply left blank in the
output — the tool tells you exactly which ones.

## What's in this folder

```
cut_plan_model_v2/
├── app.py                 # Flask local web server (3-tab single page)
├── run_cli.py               # Command-line alternative — Tab 1's extraction only, no server needed
├── cutplan2/
│   ├── __init__.py
│   └── model.py             # Extraction logic (Tab 1) + generic WIP reader (Tab 2)
├── templates/
│   └── index.html           # All 3 tabs live in this one template
├── sample_data/
│   └── BufferCuttingOrderForm_2026-08-17.xlsx   # Sample input for Tab 1
├── uploads/                 # Uploaded files land here (web server mode)
├── outputs/                 # Generated Extracted_Data.xlsx lands here
├── data/
│   └── allocation_memory.json  # Persistent Tab 3 shift-allocation memory (created on first use)
└── requirements.txt
```

This is a completely separate folder/package from the original
`cut_plan_model` project (different package name `cutplan2`, different port
`5001`), so you can have both running side by side without conflicts.

## Setup (VS Code)

1. Open this folder in VS Code (`File > Open Folder…`) — open **this**
   folder specifically, not a parent folder that also contains the old model.
2. Create and activate a virtual environment (recommended):
   ```bash
   python -m venv venv
   # macOS/Linux:
   source venv/bin/activate
   # Windows:
   venv\Scripts\activate
   ```
3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Option A — Run as a local web server

```bash
python app.py
```

By default this starts on **http://127.0.0.1:5001** (note the different port
from the old model, so both can run at once). Open that URL in your browser
— you'll see three tabs at the top:

- **Tab 1: Buffer Cutting Order Form** — upload your input `.xlsx` (or check
  "use bundled sample data") and click **Extract data**. The extracted table
  appears right below the form, with a **Download Excel** button.
- **Tab 2: WIP Upload** — upload a WIP `.xlsx` file (or check "use bundled
  sample WIP data"). It's extracted into a fixed set of fields (see below)
  with a **Download Excel** button, or falls back to a raw, unmapped
  preview if the file doesn't match the known template.
- **Tab 3: Cut Plan** — click **Generate cut plan** (uses Tab 1's extracted
  data, so run Tab 1 first). The plan is shown as an **editable table** —
  adjust any cell, add/remove rows — then **Save changes & Download Excel**
  saves your edits and downloads one Excel file per building.

Results from each tab are kept in memory (per-browser session) as you
navigate between tabs, so uploading on one tab doesn't clear what's on
another.

### How the WIP file (Tab 2) is extracted

The known WIP template repeats its header row once per group of sewing
lines (individually-named lines, then again for merged lines like
`VS02+06`). `load_wip()` in `cutplan2/model.py` doesn't assume a fixed row
range — it scans column B for the literal header text (`ไลน์เย็บ`) to find
where each block starts, and collects every non-blank row under it,
stopping the moment it hits the unrelated table-status mini-table further
down the same sheet (marked by column B reading `ไลน์`). This keeps working
even if a future day's file has a different number of sewing lines.

The ~40 columns extracted cover: targets (per hour, morning, afternoon,
OT — with/without OT), actuals for each shift, WIP quantities at each stage
(cut-waiting-to-be-bundled, bundled-waiting-for-sewing, WIP-in-line),
lead times and their status per shift, and the shortfall reasons/details
recorded for each shift. Add or rename fields via `WIP_COLUMN_MAP` at the
top of `cutplan2/model.py` if the template's column layout changes.

**Conditional-formatting colors, not just cell values, are accounted for.**
The source file color-codes several cells (red/green/orange) using Excel
conditional formatting rather than a fixed fill. Three derived "Health"
fields mirror those colors as explicit text, using the same semantics
throughout — **red → Underproduction, green → Neutral, orange →
Overproduction**:

- `WIP Cut Waiting Health` — derived from "WIP Cut Waiting to be Bundled
  (4-12 hrs)" (< 3.5 → Underproduction, 3.5–12.5 → Neutral, > 12.5 →
  Overproduction). This one has no adjacent status column of its own in the
  source file, so without this field its color-coded meaning wasn't
  captured as text anywhere.
- `Lead Time Morning Health` / `Lead Time Afternoon Health` — derived from
  "Lead Time Morning/Afternoon (hrs)". These use **different thresholds
  depending on which block of sewing lines a row came from**: individually-
  named lines use 9.5–16.5 as the Neutral band, while merged lines like
  "VS02+06" use a wider 15.5–32.5 band (matching the source file's own
  distinct conditional-formatting ranges for each block).
  `load_wip()` tracks which header block each row was read from
  automatically, so the right threshold set is applied without needing the
  block boundaries hardcoded by row number.

Each derived field sits immediately to the right of the raw value it's
computed from. See `classify_wip_cut_waiting()` and `classify_lead_time()`
in `cutplan2/model.py` to adjust the thresholds if the source file's
conditional formatting changes.

If a file doesn't match the template (the header marker isn't found),
`wip_upload()` in `app.py` automatically falls back to `load_wip_raw()` — a
plain, no-schema read — so the tab still shows *something* useful instead
of failing outright.

### How the cut plan (Tab 3) is built

For each (Sewing Line, JobCut - Suffix, Mark Type) group in Tab 1's data:

1. **Completed tables are skipped.** Status is recorded per colorway row —
   a table with mixed colorway status (e.g. one colorway Completed, another
   still Pending) isn't skipped entirely, only its completed colorway's
   quantity is dropped.
2. **Colorways sharing a Table ID are combined** into one quantity for that
   table before planning.
3. Tables are planned **smallest Table ID first**, adding tables in that
   order until the running combined quantity reaches that group's
   **Sewing Target Per Day**.
4. **You choose Morning or Afternoon when you click "Generate cut plan"** —
   a radio button on Tab 3, not the wall-clock. Choosing **Morning** means
   the planned quantity goes into **Cut Plan Afternoon of the same day**;
   choosing **Afternoon** means it goes into **Cut Plan Morning of the next
   day**. (The old automatic clock-based behavior still exists as
   `determine_shift()` and is used as a fallback for the CLI, which has no
   UI to ask the question — but the web app always uses your explicit
   choice.)
5. **The model remembers which table/JobCut has already been planned
   before, across every session — not just within one browser tab.** This
   is stored in a plain JSON file at `data/allocation_memory.json`, keyed
   by (Sewing Line, JobCut - Suffix, Mark Type, Table No.), and survives
   server restarts. If a table you already planned shows up again in a
   later "Generate cut plan" (meaning it's still incomplete — it wasn't
   marked Completed in Tab 1's data), it is **not** reset to the shift you
   just chose. Instead, it advances exactly **one step forward** from
   wherever it was last allocated (Morning → Afternoon same day; Afternoon
   → Morning next day) — so a table that keeps not getting finished keeps
   sliding through the schedule instead of resetting every time someone
   reruns the model. This happens silently in the background — it does
   **not** write anything into Note (see below); the only visible effect is
   which of Cut Plan Morning/Afternoon ends up holding the quantity. See
   `_assign_table_shift()` in `cutplan2/model.py` for the exact logic, used
   by both the initial generation and any later recalculation (e.g. raising
   a target pulls in a table that turns out to have been planned before —
   it goes through the same check).
6. **Note always starts blank, for every table — new or carried over.** The
   model never writes anything into it automatically; it's reserved
   entirely for the user's own manual comments, added on the
   editable table before saving.
7. **Sorted by Sewing Line → JobCut - Suffix → Mark Type → Table No.**, both
   on the page and in the downloaded Excel — matching the reference
   planning sheet's layout.
8. **Repeated values are merged, not repeated** — Sewing Line and JobCut -
   Suffix show/hide together, re-displaying at the start of every JobCut
   (not merged across a whole Sewing Line block). Mark Type re-displays at
   the start of every Mark Type sub-block within a JobCut. Sewing target
   per day, Cut Plan Qty, and Diff re-display at the start of every Mark
   Type sub-block too — even if the value happens to coincide with the
   previous Mark Type's value, they're shown again rather than silently
   merged across the boundary. In the downloaded Excel these are **true
   merged cells**. On the page, the repeated value is shown blank/faded
   instead — click into it to reveal and edit it (there's no hover-reveal;
   it only becomes visible when you actually focus the cell) — while a
   thick border marks a new Sewing Line and a medium border marks a new
   JobCut - Suffix, so each block is easy to pick out at a glance. Table
   No., Cut Plan Morning, Cut Plan Afternoon, and Note are never merged —
   they stay one row per table. This sort/merge/border grouping is
   re-applied after every recalculation too, not just on the initial
   "Generate cut plan".
9. **Split into two tables by building**, based on Sewing Line — "Building 1
   of `<date>`" and "Building 2 of `<date>`", shown as two separate tables
   on the page and downloaded as two **separate Excel files** (plus a third
   "Unassigned" file if any Sewing Line doesn't match either list — kept,
   never silently dropped). The date in each filename is the date you
   actually click Save/Download, not the plan's target date. The building
   lists live in `BUILDING_1_LINES` / `BUILDING_2_LINES` at the top of
   `cutplan2/model.py` — edit those if a Sewing Line's building changes.
10. **A dismissible reminder banner** on Tab 3 notes that if a JobCut has
   more than one Mark Type needing decoration, you don't have to plan every
   Mark Type or plan enough tables to fully reach the Sewing target per
   day — partial planning is fine. Click the ✕ on the banner to dismiss it;
   this is purely a one-time on-page reminder, not saved or synced anywhere.
11. **Which specific JobCuts that applies to is flagged directly in the
    table.** Using Tab 1's `Decoration` field, any JobCut with more than one
    Mark Type marked `Decoration = Yes` gets a small amber
    "⚑ Multiple Mark Types need decoration" badge under its JobCut - Suffix
    cell (shown once per JobCut, same place the merged value itself shows).
    Click the ✕ on the badge to dismiss just that one JobCut's flag — like
    the banner, this is a one-time on-page marker, not saved anywhere.

### Adjusting the plan on Tab 3

- Each building has its **own "+ Add row" button** (adds a blank row to
  that building's table only), but there's a single, shared
  **Save changes & Download Excel** button at the top that gathers rows
  from **both** tables together, then downloads **one file per building** —
  e.g. `Building 1 of 2026-08-18.xlsx` and `Building 2 of 2026-08-18.xlsx` —
  each a self-contained workbook (its own Cut Plan sheet + a Logic &
  Assumptions sheet). Your browser may ask permission the first time a page
  triggers more than one download at once; allow it to get both files.
- Click into **any** cell to edit it — including **Cut Plan Qty** and
  **Diff** — quantities, table numbers, notes, even Sewing Line /
  JobCut - Suffix / Mark Type.
- **Every edited cell shows its original value, with a one-click undo.**
  "Original" is pinned to the plan as it was right when you last clicked
  **Generate cut plan** — it does **not** shift if the plan gets
  recalculated afterward (see below), so you can always get back to the
  true starting point. As soon as a cell's value differs from that
  baseline, a small "Original: …" note appears under it with a **✕** button
  next to it — click that ✕ to instantly revert just that one cell back to
  its original value. The cell also gets a light highlight while edited.
  This applies to every row and every column. Reverting **Cut Plan Qty** or
  **Sewing target per day** also re-derives that row's **Diff** (Cut Plan
  Qty − Sewing target per day) to match — Diff itself stays freely editable
  otherwise, it just gets put back in sync specifically when you revert one
  of the two values it's derived from.
- **Editing a row's Cut Plan Qty syncs that row's Morning/Afternoon** —
  whichever shift is currently active for that row (nonzero) is updated to
  match the new quantity.
- **Editing a row's Table No. looks up that specific table's real quantity
  from Tab 1's data** and applies it to that row (Morning/Afternoon,
  whichever shift was active), then recomputes Cut Plan Qty and Diff for
  every row sharing that Sewing Line + JobCut - Suffix + Mark Type. This is
  a direct lookup, not a re-run of table selection — it won't add or drop
  other tables, it just corrects that one row's quantity for whichever
  Table No. you typed. If that table doesn't exist for this group (wrong
  number, or it's already fully Completed), the status area shows an error
  and the row is left as it was.
- **Editing Cut Plan Morning or Cut Plan Afternoon directly also
  recomputes Cut Plan Qty and Diff** for every row in that same group — so
  hand-adjusting a shift split always keeps the totals honest, whether or
  not a Table No. edit happened first.
- **Reverting Table No. also reverts that row's Morning/Afternoon** back to
  their own original values (not just the Table No. itself), then
  recomputes the group's Cut Plan Qty/Diff from the reverted numbers — so
  undoing a Table No. change fully undoes its knock-on effects, not just
  the number in that one cell.
- **Reverting Cut Plan Morning or Cut Plan Afternoon recomputes the
  group's Cut Plan Qty/Diff too.** If that was the only outstanding edit in
  the group, this brings Qty/Diff fully back to their original totals; if
  another edit is still present elsewhere in the group, it recalculates
  against whatever's still there instead of forcing a full revert.
- **Merged/blank cells never show an "edited" highlight.** A cell that's
  part of a merged block (see below) always looks blank/untouched, even if
  a group recalculation updates its value behind the scenes — the "real"
  edited indicator only ever appears on the top cell of that merged block.
- **Editing Cut Plan Qty or Sewing target per day re-checks which tables are
  needed.** This fires automatically once you finish editing the cell
  (on blur): the model re-runs table selection for that group using the
  full pool of available (non-completed) tables — pulling in more tables if
  the new target/quantities need them, or dropping tables once the planned
  quantity already reaches the target. Any Notes already typed on tables
  that stay selected are preserved, and each table's "Original" comparison
  value stays pinned to its true baseline throughout.
- **+ Add row** adds a blank row (e.g. to plan one more table by hand).
  Manually added rows (a Table No. that isn't one of the real candidate
  tables) are always kept as-is and don't get dropped by recalculation.
- **✕** next to a row removes it.
- Nothing is saved until you click **Save changes & Download Excel** — this
  regenerates the workbook with your edits and downloads it. The workbook's
  title row is stamped "Manually edited on …" once you've saved edits, so
  it's clear it no longer matches the automatic plan exactly.

Edits are held in the server's memory for the life of that session (this is
a local single-user tool, not a hosted multi-user service). If you restart
the server mid-edit, generate the cut plan again from Tab 3.

Stop the server with `Ctrl+C` in the terminal.

## Option B — Run from the command line (no server)

```bash
# Using the bundled sample data:
python run_cli.py

# Using your own file:
python run_cli.py path/to/your_input.xlsx path/to/Extracted_Data.xlsx
```

Prints a preview of the extracted data and which (if any) fields were missing.

## Running on your local network (accessible from other devices/computers)

By default the server only answers on `127.0.0.1`, meaning only the same
computer can open it. To let other devices on the same office/home network
(e.g. a planner's laptop, a shop-floor tablet) reach it too:

1. Open `app.py` and find the last line:
   ```python
   app.run(host="127.0.0.1", port=5001, debug=True)
   ```
   Change `host="127.0.0.1"` to `host="0.0.0.0"`:
   ```python
   app.run(host="0.0.0.0", port=5001, debug=True)
   ```
2. Find this computer's local network IP address:
   - **Windows:** open Command Prompt, run `ipconfig`, look for "IPv4 Address" (e.g. `192.168.1.42`)
   - **macOS:** System Settings → Network → Wi-Fi/Ethernet → look for the IP address, or run `ipconfig getifaddr en0` in Terminal
   - **Linux:** run `hostname -I` or `ip addr` in a terminal
3. Run the server as usual:
   ```bash
   python app.py
   ```
4. From another device **on the same network**, open a browser and go to:
   ```
   http://<that-ip-address>:5001
   ```
   e.g. `http://192.168.1.42:5001`

**Security note:** this makes the tool reachable by anyone on the same
network, with no login. That's fine on a trusted office/home LAN; don't do
this on a public or shared Wi-Fi network, and don't expose it to the open
internet without adding proper authentication first. Also turn `debug=True`
off (`debug=False`) once you're not actively developing — debug mode exposes
a code-execution console if something crashes.

## Extending Tab 3 to also use the WIP file, once its template is known

Tab 3 currently plans purely from Tab 1's data. When the WIP file's template
is confirmed:

1. Replace `load_wip_raw()` in `cutplan2/model.py` with a proper
   `load_wip()` that extracts a fixed set of fields from the WIP file, the
   same way `load_input()` does for the Buffer Cutting Order Form.
2. Have `wip_upload()` in `app.py` store the *parsed* WIP data (not just the
   raw preview) in `SESSIONS[sid]["wip"]`.
3. Update `build_cut_plan()` in `cutplan2/model.py` to take the WIP
   DataFrame as a second input and factor it into the planning logic (e.g.
   cross-checking against actual WIP levels).
4. Update the `cut_plan()` route in `app.py` to pass `SESSIONS[sid]["wip"]`
   through alongside the extraction data.

Keeping each tab's logic in its own function/section like this means Tab 3
can be extended without touching Tabs 1 or 2.
