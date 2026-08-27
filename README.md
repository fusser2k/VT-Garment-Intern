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
  Tab 1's extracted data together with each Sewing Line's own
  Morning/Afternoon/OT targets from Tab 2's WIP data (required — a Sewing
  Line with no matching Tab 2 coverage is skipped this run; see below).

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
used in the code (`Sewing Line`, `Cut Plan Morning`, `data-col` attributes in
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
results banner on the page always states how many rows were filtered out
this way — the downloaded workbook itself starts directly with the header
row (no title/notice row above it), so it drops straight into a plain data
table.

The bundled sample file (`BufferCuttingOrderForm_2026-08-21.xlsx`) already
uses these exact header names, so it extracts with zero missing fields. If
you point the model at a file that uses different header text, or is missing
some of these columns entirely, those fields are simply left blank in the
output — the tool tells you exactly which ones.

**Download Excel** downloads as `Extracted Data of <today's date>.xlsx`
(Thailand time), matching the naming convention used for Tab 3's
`Building N of <date>.xlsx` files.

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
  against any column), an **Excel-style filter dropdown on every column
  header** (same as Tab 3's — see below), and a **Download Excel** button.
- **Tab 2: WIP Upload** — upload a WIP `.xlsx` file (or check "use bundled
  sample WIP data"). It's extracted into a fixed set of fields (see below)
  with the same **search box**, **column filter dropdowns**, and a
  **Download Excel** button, or falls back to a raw, unmapped preview if
  the file doesn't match the known template.
- **Tab 3: Cut Plan** — click **Generate cut plan** (uses Tab 1's extracted
  data plus Tab 2's Morning/Afternoon/OT targets, so run both Tab 1 and
  Tab 2 first). The plan is shown as an **editable table** — adjust any
  cell, add/remove rows, **filter by any column** — then **Save changes &
  Download Excel** saves your edits and downloads one Excel file per
  building.

Results from each tab are kept in memory (per-browser session) as you
navigate between tabs, so uploading on one tab doesn't clear what's on
another.

### How the WIP file (Tab 2) is extracted

The known WIP template repeats its header row once per group of sewing
lines (individually-named lines, then again for merged lines like
`VS02+06`). `load_wip()` in `cutplan2/model.py` doesn't assume a fixed row
range — it scans column B for the literal header text (`ไลน์เย็บ`, or the
English `Sewing line` used in some template revisions — matched
case-insensitively) to find where each block starts, and collects every
non-blank row under it, stopping the moment it hits the unrelated
table-status mini-table further down the same sheet (marked by column B
reading `ไลน์`). This keeps working even if a future day's file has a
different number of sewing lines.

Column A of each row (**Sewing Line Code**, e.g. `VSEW012`) is extracted
too — this is what ties a WIP row back to Tab 1's own Sewing Line
identifier for the Tab 3 target override described above. **When column A
is blank** (not every day's WIP file has it filled in), the code is
recognized instead from that row's own **Sewing Line** Thai name/team text
(e.g. "ราตรี /") against a fixed lookup table
(`SEWING_LINE_THAI_ROOTS`/`match_sewing_line_code()` in
`cutplan2/model.py`) — tolerant of the name ending without a trailing "/",
with an extra number, or nothing at all after the name. A genuine code
already present in the file is always kept as-is and never overridden by
the Thai-name guess. Either way, the resulting code is what's shown in
Tab 2's own "Sewing Line Code" column (and its downloaded Excel) and what
Tab 3's target override matches against — there's no separate matching
logic in two places.

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

**Tab 2's WIP data is now required, not optional.** Planning is built around
each Sewing Line's own **Morning / Afternoon / OT targets** from Tab 2 (its
`Target Morning`, `Target Afternoon`, `Target OT` columns), not a single
daily total. If Tab 2 hasn't been uploaded this session (or its file didn't
match the known template), Tab 3 won't generate a plan at all — it'll ask
you to run Tab 2 first.

For each (Sewing Line, JobCut - Suffix, Mark Type) group in Tab 1's data:

1. **Only rows with Status exactly "Pending" are planning candidates.**
   Status is recorded per colorway row — a table with mixed colorway status
   (e.g. one colorway Pending, another Completed) isn't skipped entirely,
   only its non-Pending colorway's quantity is dropped; the Pending
   colorway(s) are still planned. Both "Completed" **and** "In Progress"
   are excluded — this is stricter than just excluding Completed alone.
2. **A Sewing Line with no matching row in Tab 2's WIP data is skipped
   entirely this run** — none of its tables get planned at all, rather than
   guessing with a made-up target. A note on Tab 3's results lists exactly
   which Sewing Lines were skipped this way, so it's never a silent gap.
   **Matching a Tab 1 Sewing Line (e.g. "VSEW012") to a row in Tab 2's WIP
   data** happens via Tab 2's own **"Sewing Line Code"** column (see "How
   the WIP file is extracted" above for exactly how that column gets
   populated, including its Thai-name-recognition fallback). Merged lines
   (two individual lines sharing one combined line, e.g. "นิภาพร +
   ชุติมันต์" → `VS02+06`, since 02 and 06 are the two individual lines' own
   numbers) are handled the same way. Update
   `SEWING_LINE_THAI_ROOTS`/`SEWING_LINE_MERGED_ROOTS` in
   `cutplan2/model.py` if a line's assigned supervisor/team changes.
3. **Colorways sharing a Table ID are combined** into one quantity for that
   table before planning.
4. **Tables are planned smallest Table ID first, one session at a time —
   Morning, then Afternoon, then OT.** A running **Diff** = (total quantity
   planned in the group so far, across every table and session) minus (the
   sum of every session's target up to and including the one currently
   being filled). The moment a table's addition pushes that Diff to zero or
   above, that table is the **last one for the current session** — the next
   table moves into the next session, and the cumulative target immediately
   jumps up by that whole next session's target. Once OT's Diff also
   reaches zero or above, or all three sessions have been used, the group
   stops — any remaining tables for that group stay pending for a future
   run rather than being force-fit into an already-full day. Each table
   belongs to **exactly one session**: its own quantity appears in that
   session's Cut Plan column, with the other two sessions' Cut Plan columns
   left at 0 for that row.

   **Worked example** (targets Morning 144 / Afternoon 154 / OT 115, tables
   2–7 with quantities 133, 60, 65, 36, 12, 8):

   | Table | Session   | Cut Plan | Diff |
   |-------|-----------|---------:|-----:|
   | 2     | Morning   | 133      | −11  |
   | 3     | Morning   | 60       | 49   |
   | 4     | Afternoon | 65       | −40  |
   | 5     | Afternoon | 36       | −4   |
   | 6     | Afternoon | 12       | 8    |
   | 7     | OT        | 8        | −99  |

   Table 3's Diff (49) carries into Table 4's calculation: 49 + 65 − 154 =
   −40. This carry-forward is why Diff can look larger than a single
   table's own quantity — it's cumulative across the whole group, not reset
   per table.

   **A note on rounding:** each session's own target (144/154/115 above) is
   rounded UP individually from Tab 2's fractional value, since the
   cascading selection needs a whole number to compare Diff against for
   each session. **Sewing Target Per Day**, shown as a separate reference
   column, is **not** the sum of those three already-rounded numbers —
   summing individually-rounded values can overshoot the true total by a
   couple of units (e.g. Morning 144.0 + Afternoon 153.6 + OT 115.2 = 412.8,
   which should round up to 413 — but 144 + 154 + 116 = 414). It's always
   computed as the ceiling of the raw fractional total instead, so it
   matches Tab 2's own "Target for Day (with OT)" figure rounded up exactly
   once.
5. **Exception — oversized tables:** if a table's own Qty is more than
   double the group's Morning target, **and** the group hasn't yet moved
   past Morning (no earlier table's Diff has turned non-negative), that
   table isn't split across sessions at all — its full Qty goes into Cut
   Plan Morning alone, and the **entire group** stops planning right there,
   regardless of how many other tables are still pending for that Mark
   Type. An oversized table encountered *after* the group has already moved
   on to Afternoon/OT is planned normally into whichever session it's
   actually in — this exception only applies to the group's still-in-Morning
   phase.
6. **Note always starts blank, for every table.** The
   model never writes anything into it automatically; it's reserved
   entirely for the user's own manual comments, added on the
   editable table before saving.
7. **Sorted by Sewing Line → JobCut - Suffix → Mark Type → Table No.**, both
   on the page and in the downloaded Excel — matching the reference
   planning sheet's layout.
8. **Repeated values are merged, not repeated**, at three different levels
   of granularity:
   - Sewing Line and JobCut - Suffix show/hide together, re-displaying at
     the start of every JobCut (not merged across a whole Sewing Line
     block).
   - Mark Type and Sewing Target Per Day re-display at the start of every
     Mark Type sub-block within a JobCut.
   - Each session's own target (Sewing target Morning/Afternoon/OT)
     re-displays at the start of **its own session's block of rows** — even
     within the same, unchanged Mark Type. In the worked example above, the
     Afternoon target shows again on Table 4's row, even though it's still
     the same Mark Type as Tables 2–3's Morning rows.
   - **Exception:** a Mark Type group with only **one row total** (e.g. an
     oversized table stopping the group right there, or simply a group with
     only one Pending table) shows **all three** session targets on that
     one row, not just the one matching its own session — there's no "next
     row" to show the others on, and seeing all three gives useful context
     (e.g. confirming Morning alone already covers far more than a full
     day's target, so Afternoon/OT genuinely aren't needed for that group).

   In the downloaded Excel these are **true merged cells**. On the page,
   the repeated value is shown blank/faded instead — click into it to
   reveal and edit it (there's no hover-reveal; it only becomes visible
   when you actually focus the cell) — while a thick border marks a new
   Sewing Line and a medium border marks a new JobCut - Suffix, so each
   block is easy to pick out at a glance. Table No., every Cut Plan/Diff
   column, and Note are never merged — they stay one row per table. This
   sort/merge/border grouping is re-applied after every recalculation too,
   not just on the initial "Generate cut plan".
9. **Split into two tables by building**, based on Sewing Line — "Building 1
   of `<date>`" and "Building 2 of `<date>`", shown as two separate tables
   on the page and downloaded as two **separate Excel files** (plus a third
   "Unassigned" file if any Sewing Line doesn't match either list — kept,
   never silently dropped). The date in each filename is the date you
   actually click Save/Download, not the plan's target date. The building
   lists live in `BUILDING_1_LINES` / `BUILDING_2_LINES` at the top of
   `cutplan2/model.py` — edit those if a Sewing Line's building changes.
10. **A dismissible reminder banner** on Tab 3 notes that if a JobCut has
    more than one Mark Type needing decoration, you don't have to plan
    every Mark Type or plan enough tables to fully reach every session's
    target — partial planning is fine. Click the ✕ on the banner to
    dismiss it; this is purely a one-time on-page reminder, not saved or
    synced anywhere.
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
  each a self-contained workbook with just its own Cut Plan sheet. Your
  browser may ask permission the first time a page triggers more than one
  download at once; allow it to get both files.
- Click into **any** cell to edit it — every Cut Plan quantity and every
  Diff, table numbers, notes, even Sewing Line / JobCut - Suffix / Mark
  Type.
- **Every edited cell shows its original value, with a one-click undo.**
  "Original" is pinned to the plan as it was right when you last clicked
  **Generate cut plan** — it does **not** shift if the plan gets
  recalculated afterward (see below), so you can always get back to the
  true starting point. As soon as a cell's value differs from that
  baseline, a small "Original: …" note appears under it with a **✕** button
  next to it — click that ✕ to instantly revert just that one cell back to
  its original value. The cell also gets a light highlight while edited.
  This applies to every row and every column. Reverting a Cut Plan
  quantity, a session target, or Table No. triggers a full recalculation
  (see below) rather than trying to patch just that one cell — the
  three-session cascading selection depends on the whole group's state, so
  it's re-derived properly from the server rather than guessed at in the
  browser.
- **Editing a row's Table No. looks up that specific table's real quantity
  from Tab 1's data** and puts it into whichever of the three Cut Plan
  columns (Morning/Afternoon/OT) that row currently occupies, then triggers
  a full recalculation. This is different from a normal quantity edit in
  one respect: it's a direct lookup for a *specific* table you typed in,
  not a re-run of table selection from scratch — but since Table No. can
  change which table a row represents entirely, letting the group's Diff
  values and session assignments recalculate afterward keeps everything
  consistent. If that table doesn't exist for this group (wrong number, or
  its Status isn't exactly "Pending"), the status area shows an error and
  the row is left as it was, but Table No.'s relocation to the correct
  sorted position still happens either way.
- **Merged/blank cells never show an "edited" highlight.** A cell that's
  part of a merged block (see below) always looks blank/untouched, even if
  a group recalculation updates its value behind the scenes — the "real"
  edited indicator only ever appears on the top cell of that merged block.
- **Editing any Cut Plan quantity or any session target (Sewing target
  Morning/Afternoon/OT) re-checks which tables are needed.** This fires
  automatically once you finish editing the cell (on blur): the model
  re-runs the same three-session cascading selection for that group using
  the full pool of available (Status = "Pending") tables — pulling in more
  tables if a target now needs them, or dropping tables once the planned
  quantity already reaches every session's target. Any Notes already typed
  on tables that stay selected are preserved, and each table's "Original"
  comparison value stays pinned to its true baseline throughout.
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
  dropdown, so it always reflects what's currently in the table. **The
  exact same filter dropdowns are also available on Tab 1's and Tab 2's
  tables** (alongside their own search box), independent of whether a cut
  plan has even been generated — the filtering system works the same way
  everywhere in the app, just applied to read-only tables there instead of
  editable ones.
- **Every column can be resized** — hover the right edge of a column
  header (a thin drag handle appears) and drag to widen or narrow it,
  handy when a value like a longer Sewing Line code gets cut off at the
  default width. **Double-click that same edge to auto-fit the column** to
  its widest currently-visible content in one click. This works the same
  way on **all three tables** (Tab 1, Tab 2, Tab 3), not just Tab 3's.
- Nothing is saved until you click **Save changes & Download Excel** — this
  regenerates the workbook with your edits and downloads it. Downloaded
  files start directly with the header row (no title/notice rows above it,
  and no "Manually edited on …" stamp either) — a plain data table either
  way.

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


