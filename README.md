# Cut Planning Model v2

A **new, separate** project from the original `cut_plan_model` — a fresh
build for cut planning. It currently has three tabs, all on one page:

- **Tab 1: Buffer Cutting Order Form** — upload the cutting order Excel file
  and it's extracted into a fixed 17-field schema.
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

## Language: English / Thai

**EN** / **ไทย** buttons in the top-right switch the entire UI — labels,
buttons, descriptions, flash messages, table column headers on all three
tabs, and the JavaScript-driven status text (Saving…, Plan updated…, etc.).
The choice is stored in your browser session and applies across all three
tabs immediately.

Internally, only the *display* text changes — every internal column name
used in the code (`Sewing Line`, `Cut Plan Qty`, `data-col` attributes in
the HTML, Python dict keys, etc.) always stays in English, so translations
never risk breaking any of the app's logic. All translated strings live in
one place: `cutplan2/i18n.py` — `TRANSLATIONS` for UI text and
`COLUMN_LABELS` for table headers. To add a language, add a third language
code there (e.g. `"vi"`) to every entry, add it to `SUPPORTED_LANGS`, and
add a corresponding button next to the EN/ไทย ones in
`templates/index.html`'s topbar.

## Tab 1: the 17 extracted fields

```
Sewing Line | JobCut - Suffix | Table ID | Colorway | Mark Type | Layer |
Qty | Qty Complete | Difference | % Complete | Status |
Sewing Target Per Day (with OT) | Sewing Target Per Day (no OT) |
Table No. (Mark Type 101) | Decoration | FG Start Date | Start Cut
```

**Sewing Target Per Day (no OT) is computed, not extracted from the file** —
it's always calculated as `Sewing Target Per Day (with OT) ÷ 10.75 × 7.75`
(scaling the with-overtime daily target down to a regular, no-overtime
10.75h → 7.75h workday), so it never shows up in the "fields not found"
list even if the source file has no column for it — it only ends up blank
if the WITH-OT value it's computed from is itself missing.

**Filter applied automatically:** only rows whose **Sewing Line** starts with
`VS` are kept (e.g. `VSEW012`, `VS02+06`). Anything else (e.g. sample/test
rows like `SAMPL02`) is dropped before the data is shown or saved. The
results banner on the page — and the title row of the downloaded workbook —
always states how many rows were filtered out this way.

The bundled sample file (`BufferCuttingOrderForm_2026-08-21.xlsx`) already
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
│   └── BufferCuttingOrderForm_2026-08-21.xlsx   # Sample input for Tab 1
├── uploads/                 # Uploaded files land here (web server mode)
├── outputs/                 # Generated Extracted_Data.xlsx lands here
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
  appears right below the form, with a **search box** above it (matches
  against any column) and a **Download Excel** button.
- **Tab 2: WIP Upload** — upload a WIP `.xlsx` file (or check "use bundled
  sample WIP data"). It's extracted into a fixed set of fields (see below)
  with the same **search box** and a **Download Excel** button, or falls
  back to a raw, unmapped preview if the file doesn't match the known
  template.
- **Tab 3: Cut Plan** — click **Generate cut plan** (uses Tab 1's extracted
  data, so run Tab 1 first). The plan is shown as an **editable table** —
  adjust any cell, add/remove rows, **filter by any column** — then
  **Save changes & Download Excel** saves your edits and downloads one
  Excel file per building.

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
   **Sewing Target Per Day (with OT)** — Tab 3's planning always uses the
   with-OT figure, not the computed no-OT one (which is purely informational
   on Tab 1).
4. **Cut Plan Morning / Cut Plan Afternoon is decided per group, purely
   from the selected tables' quantities** — no wall-clock time, no user
   choice, no cross-session memory involved:
   - If a group has **exactly one** selected table, its whole quantity goes
     to **Morning** by default.
   - If a group has **more than one** selected table, they're split into a
     "top" part (Morning) and a "bottom" part (Afternoon): walking the
     tables in Table ID order, each stays in the top/Morning part and its
     quantity accumulates, until the running total reaches **half of the
     group's own Cut Plan Qty** — the total quantity actually being
     allocated across the selected tables, **not** half of Sewing Target
     Per Day (the selected tables often don't add up anywhere near the
     target, and the split still needs to happen either way). The table
     whose addition crosses that halfway point is still counted in the
     top/Morning part; every table after that point goes to the
     bottom/Afternoon part — so when the total can't split into two exactly
     equal halves, the larger portion normally lands in Morning.
   - **A split always happens whenever there's more than one table.** If
     one large table (typically the last one in Table ID order) is big
     enough that the halfway point isn't crossed until it's added — which
     would otherwise leave Afternoon empty — that table is moved to
     Afternoon instead. This is the one case where the larger side can end
     up in Afternoon rather than Morning; guaranteeing an actual split
     takes priority.
   - This is fully deterministic and gets **recomputed fresh** every time
     the plan is generated or recalculated (e.g. editing a target or a
     table's quantity can shift where the halfway point falls, moving
     tables between Morning and Afternoon accordingly). See
     `split_morning_afternoon()` in `cutplan2/model.py` for the exact logic.
5. **Note always starts blank, for every table.** The
   model never writes anything into it automatically; it's reserved
   entirely for the user's own manual comments, added on the
   editable table before saving.
6. **Sorted by Sewing Line → JobCut - Suffix → Mark Type → Table No.**, both
   on the page and in the downloaded Excel — matching the reference
   planning sheet's layout.
7. **Repeated values are merged, not repeated** — Sewing Line and JobCut -
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
8. **Split into two tables by building**, based on Sewing Line — "Building 1
   of `<date>`" and "Building 2 of `<date>`", shown as two separate tables
   on the page and downloaded as two **separate Excel files** (plus a third
   "Unassigned" file if any Sewing Line doesn't match either list — kept,
   never silently dropped). The date in each filename is the date you
   actually click Save/Download, not the plan's target date. The building
   lists live in `BUILDING_1_LINES` / `BUILDING_2_LINES` at the top of
   `cutplan2/model.py` — edit those if a Sewing Line's building changes.
9. **A dismissible reminder banner** on Tab 3 notes that if a JobCut has
   more than one Mark Type needing decoration, you don't have to plan every
   Mark Type or plan enough tables to fully reach the Sewing target per
   day — partial planning is fine. Click the ✕ on the banner to dismiss it;
   this is purely a one-time on-page reminder, not saved or synced anywhere.
10. **Which specific JobCuts that applies to is flagged directly in the
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
- **A small ⇄ button on each Cut Plan Morning/Afternoon cell moves that
  row's quantity to the opposite shift** — click it and Morning becomes
  Afternoon (or vice versa) for that one row. Since this is just another
  edit, it gets the same **"Original: …" + ✕ revert** treatment as any
  other cell — click revert on *either* side afterward and both cells
  return to their true original values together (not just the one you
  clicked, which would otherwise leave the same quantity duplicated in
  both cells).
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
- **+ Add row** adds a blank row **at the top** of that building's table —
  easy to find without scrolling. Fill in its Sewing Line, JobCut - Suffix,
  Mark Type, and Table No. (in that order, tabbing through as normal), and
  once you finish editing **Table No.**, the row automatically relocates to
  its correct sorted position within the matching group, exactly as if it
  had been part of the original plan. Manually added rows (a Table No. that
  isn't one of the real candidate tables) are always kept as-is and don't
  get dropped by recalculation. Editing Table No. on an existing row
  relocates it the same way.
- **Hover over any row to reveal a small ✚ button** (next to the ✕) that
  inserts a blank row directly below it — handy when you want to start
  filling in a new table right next to a specific one you're already
  looking at, instead of scrolling to the top. It relocates the same way
  as above once you fill in its Table No.
- **✕** next to a row removes it.
- **Each column has an Excel-style filter dropdown**, right under the
  headers — click **Filter ▾** to see every unique value currently in that
  column (with a count for each), check/uncheck individual values, or use
  **Select All** / **Clear All**. A **search box** at the top of the
  dropdown narrows the checkbox list as you type (case-insensitive partial
  match) — handy for a long list like Sewing Line's — without touching the
  actual table filter itself; Select All / Clear All still act on every
  unique value regardless of what's currently searched for. Multiple
  columns' filters combine with AND (a row must match every active filter).
  The button highlights blue while a filter is active on that column.
  **Filtering is purely visual** — filtered-out rows stay fully part of the
  data; editing, recalculating, and Save changes & Download Excel always
  work against the complete table regardless of what's currently filtered.
  The building's **Clear filters** button resets every column's filter at
  once. Filters stay in effect through recalculation and adding rows (a
  freshly rebuilt table re-applies whatever's currently selected); the
  unique-values list itself is recomputed fresh every time you open a
  dropdown, so it always reflects what's currently in the table.
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

The server already binds to `0.0.0.0` by default (see the note on Railway
deployment below for why), so any device on the same network can already
reach it once you know this computer's address:

1. Find this computer's local network IP address:
   - **Windows:** open Command Prompt, run `ipconfig`, look for "IPv4 Address" (e.g. `192.168.1.42`)
   - **macOS:** System Settings → Network → Wi-Fi/Ethernet → look for the IP address, or run `ipconfig getifaddr en0` in Terminal
   - **Linux:** run `hostname -I` or `ip addr` in a terminal
2. Run the server as usual:
   ```bash
   python app.py
   ```
3. From another device **on the same network**, open a browser and go to:
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

## Deploying to Railway (or a similar host)

`app.py` and `requirements.txt` are already set up for this:

- `app.py`'s startup reads the `PORT` environment variable Railway assigns
  (falling back to `5001` locally) and always binds `0.0.0.0` — the two
  things a hardcoded `127.0.0.1`/fixed-port setup gets wrong, which is what
  causes Railway's **"Application failed to respond"** error.
- A `Procfile` in the project root tells Railway exactly how to start the
  app with `gunicorn` (a production-grade server) instead of Flask's own
  dev server:
  ```
  web: gunicorn app:app --bind 0.0.0.0:$PORT --workers 1 --threads 4 --timeout 120
  ```
- `gunicorn` is included in `requirements.txt`.

**If you still see "Application failed to respond" after this:** open the
**Deployments** tab in Railway, click the failed deployment, and check
**Deploy Logs** — it shows the actual Python exception (missing dependency,
import error, etc.) rather than just the generic error page.

**Important: keep `--workers 1`.** This app keeps each browser's Tab 1/2/3
data in an in-memory `SESSIONS` dict (see `app.py`) — that memory is *not*
shared between separate worker processes. Running with more than one worker
would randomly show a different (or empty) session depending on which
worker handled a given request. If you need more concurrency, increase
`--threads` instead of `--workers`.

**The filesystem (and `SESSIONS`) is reset on every redeploy.** Uploaded
files and generated Excel files live on Railway's container filesystem,
which persists between requests but is wiped clean on every new
deploy/restart. Since none of the app's core planning logic depends on
persisted state between deploys (Tab 3's Morning/Afternoon split is
recomputed fresh from Tab 1's data every time), this doesn't affect
correctness — it just means uploaded files and past sessions don't survive
a redeploy, same as restarting the app locally.

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
