# Prophet Table Change Tool

Consolidate quarterly Prophet table changes and integrate them safely into the production table set.

You drive the tool with an Excel **Control** workbook. The tool compares `before/` vs `after/` tables, writes an auditable Change Log, then validates and (optionally) applies those changes onto production CSVs.

---

## Requirements

- **Anaconda or Miniconda** — day-to-day users do not need a separate Python install
- Shared Conda env `prophet-table` from [`environment.yml`](environment.yml) (Python **3.11**, **polars 1.44.2**, `openpyxl`, `xlsxwriter`)
- [`requirements.txt`](requirements.txt) remains available for pip / pytest workflows (pytest is not in the day-to-day Conda env)

`Setup.bat` / `Run.bat` look for `conda` on PATH and in common install folders. If they still cannot find Conda, open **Anaconda Prompt**, `cd` to this project folder, and run the `.bat` from there.

---

## Quick start

1. **(First time)** Double-click [`Setup.bat`](Setup.bat) to create/update the `prophet-table` Conda env.
2. Set up a working folder (see [Folder layout](#folder-layout)) and fill in `Control.xlsx` (see [Control workbook](#control-workbook)).
3. Set `mode` in the Control `Config` sheet (`generate_changelog`, `validate_only`, or `apply`).
4. Double-click [`Run.bat`](Run.bat), or drag `Control.xlsx` onto it, and confirm the Control path if prompted.
5. For Stage 2 (`validate_only` / `apply`), ensure a Change Log already exists in `Output\` (from Stage 1). When prompted, leave **Change Log** blank to use the newest `ChangeLog_*.xlsx` there, or type a path / filename to pin a specific file. (Stage 1 can leave this blank.) You can also pass the Change Log as a second argument: `Run.bat Control.xlsx ChangeLog_YYYYMMDD_HHMMSS.xlsx`.

Typical sequence: set `mode` to `generate_changelog` → run → review the Change Log → set `mode` to `validate_only` → run → set `mode` to `apply` → run.

### CLI alternative

From the project root, with the `prophet-table` env active (or via `conda run -n prophet-table`):

```bash
# Stage 1 – generate Change Log
python -m prophet_table_tool path\to\Control.xlsx --mode generate_changelog

# Stage 2 – dry run (no new tables written); omit --change-log to use the newest in Output
python -m prophet_table_tool path\to\Control.xlsx --mode validate_only

# Stage 2 – apply a specific Change Log (full path, or filename under Output)
python -m prophet_table_tool path\to\Control.xlsx --mode apply --change-log ChangeLog_YYYYMMDD_HHMMSS.xlsx
```

`--mode` overrides the `mode` value in the Control `Config` sheet. If you omit `--mode` (as `Run.bat` does), the tool uses whatever is set in Control. For Stage 2, `--change-log` is optional: omit it to auto-pick the newest `ChangeLog_*.xlsx` in `Output\`.

---

## Folder layout

Create a working root like this (paths are relative to `Control.xlsx` by default):

```
WorkingRoot/
├── Control.xlsx
├── ChangeRequests/
│   ├── CR_2026Q3_01_MortalityUpdate/     ← folder name = change_request_id
│   │   ├── before/
│   │   │   └── *.csv
│   │   └── after/
│   │       └── *.csv
│   └── CR_2026Q3_02_ExpenseRevision/
│       ├── before/
│       └── after/
├── Production_Tables/                    ← current production snapshot
│   └── *.csv
└── Output/                               ← created/updated by the tool
    ├── ChangeLog_YYYYMMDD_HHMMSS.xlsx
    ├── ChangeLog_YYYYMMDD_HHMMSS_Detail.csv
    ├── ChangeLog_YYYYMMDD_HHMMSS_reviews/  ← per-folder human review workbooks
    │   ├── root.xlsx
    │   ├── BE.xlsx
    │   └── IFRS.xlsx
    ├── IntegrationReport_YYYYMMDD_HHMMSS.xlsx
    ├── New_Production_Tables/            ← written only in apply mode
    └── Audit/
        └── *.log
```

**Tips**

- Put the tables you are changing under each change request’s `before/` and `after/` folders, using the **same filenames** as in production (e.g. `MORT_TABLE.csv`).
- `before/` should match the baseline you are changing from (usually the current production version of those tables).
- Tables that appear only in `after/` are treated as **new tables** (`table_add`).
- Tables that appear only in `before/` produce a **warning** and are skipped (nothing is auto-deleted).

---

## Prophet CSV / FAC format

Tables may be named `*.csv` or `*.FAC` (Prophet export). Content is detected as the block from the first `!N` header through the following `*` data rows:

```csv
Prophet table,,,,,
!4,Age,Duration,Product,Rate,Loading
*,20,1,PROD_A,0.0012,1.05
*,25,1,PROD_A,0.0015,1.05
Edited on 2026,,,,,
```

| Rule | Detail |
|------|--------|
| Header marker | First cell is `!N`, where `N` is the number of **key columns including the marker column itself** |
| Headers | Remaining cells on the `!N` row are column names |
| Data rows | First cell is always `*`, then key values and data values |
| Dummy lines | Lines before `!N` and after the last `*` row are ignored for Change Log comparison; Stage 2 keeps production’s dummy lines as-is |
| Matching | Exact string match only (no floating-point tolerance) |

Example: `!4` means keys are the marker + the next 3 columns (`Age`, `Duration`, `Product`).

---

## Control workbook

`Control.xlsx` is the single source of truth. It needs these sheets:

### `Config`

| Key | Example | Description |
|-----|---------|-------------|
| `working_root` | `D:\ProphetChanges\2026Q3` | Root folder. Leave blank to use the folder that contains `Control.xlsx`. |
| `mode` | `generate_changelog` / `validate_only` / `apply` | Default run mode (can be overridden on the CLI) |
| `production_tables_path` | `Production_Tables` | Relative to working root, or absolute |
| `output_path` | `Output` | Relative to working root, or absolute |
| `run_id` | (optional) | If blank, auto-generated as `YYYYMMDD_HHMMSS` |

### `ChangeRequests`

| Column | Meaning |
|--------|---------|
| `change_request_id` | Must match the folder name under `ChangeRequests/` |
| `order` | Integer; Stage 2 applies CRs in this sequence |
| `include` | `Y` / `N` — whether to process this CR |
| `approved` | `Y` / `N` — required for Stage 2 (`validate_only` and `apply`) |
| `description` | Short description (shown in Change Log Summary) |
| `notes` | Optional |

### `ColumnRenames` (optional)

Column renames are **never auto-detected**. Declare them here or Stage 2 will hard-stop.

| change_request_id | table_name | old_column_name | new_column_name |
|-------------------|------------|-----------------|-----------------|
| CR_2026Q3_01_… | MORT_TABLE | OldRate | Rate |

### `KeyCountApprovals` (optional)

Use when a table’s `!N` changes. Without approval, Stage 2 hard-stops.

| change_request_id | table_name | approved |
|-------------------|------------|----------|
| CR_2026Q3_03_… | EXPENSE_TABLE | Y |

If this sheet has no row for a table, the tool falls back to the CR-level `approved` flag.

---

## Workflow in detail

### Stage 1 — Generate Change Log

Compares `before/` vs `after/` for every included change request (including
tables in **subfolders**; identity is the relative path, e.g. `SubA/MORT_TABLE`)
and writes:

- `Output/ChangeLog_<run_id>.xlsx` — Summary, Conflicts, and a **ReviewFiles** index (Stage 2 uses this file plus the Detail CSV)
- `Output/ChangeLog_<run_id>_Detail.csv` — machine-readable change rows used by Stage 2
- `Output/ChangeLog_<run_id>_reviews/<group>.xlsx` — human-review workbooks split by first-level production-table folder (`root` for tables at the top, `BE` for `BE/...`, `IFRS` for `IFRS/...`, and so on)

Canonical workbook sheets:

- **Summary** — per-CR counts and conflict flags
- **Conflicts** — one row per overlap or gap (see [Conflicts](#conflicts))
- **ReviewFiles** — maps each table to its folder review workbook

Folder review workbooks:

- **Conflicts** — rows whose `table_name` belongs to that folder
- **One sheet per touched table** — human review only (changed rows); Stage 2 ignores these sheets. The `_change` column shows `[change_request_id]` for each changed row, with an `ADD` / `DELETE` prefix when the row was added or removed.

Value comparison is **numeric-aware** (`1.10` equals `1.1`). CSV files are read
with encoding fallback: utf-8-sig → utf-8 → cp1252 → latin-1.

```bash
python -m prophet_table_tool WorkingRoot\Control.xlsx --mode generate_changelog
```

**Review the Change Log before continuing.** Prefer the per-folder review workbooks under `ChangeLog_*_reviews/` for visual checks; use the Detail CSV / Conflicts for the machine-readable record. If the Conflicts sheet has rows, resolve them before applying (see [Conflicts](#conflicts)).

### Stage 2 — Validate only (dry run)

Checks that the Change Log can be applied cleanly against current production, **in Control `order`**:

- Referenced tables exist (except pure `table_add`)
- Each CR is validated against the in-memory table **after earlier CRs** have been applied (simulated; no files written)
- `value_update` / `row_delete` keys must exist in that current table (`old_value` is not required to match)
- `row_add` keys must **not** already exist in that current table (no duplicate rows)
- Column renames are declared; key-count changes are approved
- Unresolved conflicts are listed as FAIL rows (see [Conflicts](#conflicts))

Writes `Output/IntegrationReport_<run_id>.xlsx` and an audit log. **Does not write** any files under `New_Production_Tables/`.

```bash
python -m prophet_table_tool WorkingRoot\Control.xlsx --mode validate_only
python -m prophet_table_tool WorkingRoot\Control.xlsx --mode validate_only --change-log WorkingRoot\Output\ChangeLog_YYYYMMDD_HHMMSS.xlsx
```

### Stage 2 — Apply

Same validation as above. If everything passes, writes updated CSVs to:

`Output/New_Production_Tables/`

Filenames and `!N` / `*` format are preserved. Production input files are not overwritten in place.

```bash
python -m prophet_table_tool WorkingRoot\Control.xlsx --mode apply
python -m prophet_table_tool WorkingRoot\Control.xlsx --mode apply --change-log ChangeLog_YYYYMMDD_HHMMSS.xlsx
```

Stage 2 only processes change requests with **`include = Y` and `approved = Y`**, in `order` sequence.

---

## Conflicts

Conflicts are detected in **Stage 1** after every included change request has been compared (`before/` vs `after/`). They are written to the Change Log **Conflicts** sheet (`resolved` starts as `N`). The Summary sheet sets `has_conflict = Y` and `status = CONFLICT` for each CR involved.

Stage 2 does **not** re-detect conflicts from the Detail CSV. It reads the Conflicts sheet as stored in the Change Log workbook. `apply` **hard-stops** (writes no files under `New_Production_Tables/`) while any conflict that still involves an `include=Y` + `approved=Y` CR has `resolved` other than `Y`. `validate_only` records the same rows as FAIL in the Integration Report so you can see them before applying.

A conflict is ignored in Stage 2 only if you mark `resolved=Y`, delete that Conflicts row, or exclude **every** CR listed on that row (`include=N` or `approved=N`). Excluding just one of the listed CRs is not enough unless you also mark the row resolved.

### When each type is raised

| Type | Raised when | Typical example |
|------|-------------|-----------------|
| `cell_overlap` | Two or more CRs touch the **same cell** (same table + same key combination + same column) with `value_update`, `row_add`, or `row_delete`. Still flagged when both CRs write the same `new_value` (numeric-aware: `1.10` equals `1.1`), and when a `row_add` or `column_add` is followed by a `value_update`. Set `resolved=Y` to apply both in Control `order` (later CR wins). | CR_A and CR_B both change `Rate` for `Age=20, Duration=1, Product=PROD_A`. |
| `structural_collision` | Two or more CRs apply **structural** changes to the **same table**. Structural types: `column_add`, `column_delete`, `column_rename`, `key_count_change`, `table_add`. Different columns still collide if they are on the same table. | CR_A adds a column and CR_B renames another column on `MORT_TABLE`. |
| `missing_row_column_fill` | A `row_add` and a `column_add` on the same table leave their **intersection cell** with no covering `row_add` / `value_update` in the Detail CSV. Can be one CR or several. | CR_A adds a new age row; CR_B adds `NewCol`; neither supplies a value for that new row × `NewCol`. |

Not a conflict (handled elsewhere): a table only in `before/` (warning + skip), an undeclared column rename, or an unapproved key-count change. Those are separate hard-stops — see [Safety rules](#safety-rules-what-will-block-you).

### How to find them

1. Open `Output/ChangeLog_*.xlsx`.
2. **Summary** — any CR with `has_conflict = Y`.
3. **Conflicts** — one row per issue: type, table, key, column, CR ids, old/new values, notes, `resolved`.
4. **ReviewFiles** — which folder workbook holds each table.
5. **Per-folder review workbooks** (`ChangeLog_*_reviews/<group>.xlsx`) — useful for `cell_overlap`: `_change` names the related CR(s); changed cells show `old -> new` and `[change_request_id]` when more than one CR touches the table. These sheets are human review only; Stage 2 does not read them.

### How to resolve them

Prefer fixing the **source** (Control and/or `before/` / `after/`) and **re-running Stage 1**. Regenerating the Change Log rebuilds Conflicts from scratch (`resolved` resets to `N` if the overlap is still present).

**`cell_overlap`** — still listed with `resolved=N`. Conflicts `notes` say whether the `new_value`s agree, and whether this is a sequenced `row_add` / `column_add` plus `value_update`.

Set `resolved = Y` to accept the overlap and apply **both** remaining Detail rows in Control `order`. The **later** CR’s value is the one that remains in the cell. You do not have to delete a Detail row when the values differ.

Typical cases:

- **Same `new_value`** (numeric-aware): keep both Detail rows and set `resolved = Y`.
- **`row_add` then `value_update`** (CR_B’s `before/` already includes the new row): `notes` mention `row_add and value_update`. Set `resolved = Y`. Do **not** model the follow-on as a second `row_add` of the same key: even with `resolved = Y`, Stage 2 fails `row_add` when the key already exists.
- **`column_add` then `value_update`** (CR_B’s `before/` already includes the new column): `notes` mention `column_add and value_update`. Set `resolved = Y`. CR_A adds the column; CR_B’s later value stays.
- **Independent updates with different `new_value`s**: you may still exclude a CR or edit Detail to pick a winner; or just set `resolved = Y` and accept Control-order last-writer-wins.

To drop one CR’s change instead of last-writer-wins:

- Set `include = N` on the losing CR in Control and re-run Stage 1 (cleanest).
- Edit that CR’s `before/` / `after/` so it no longer changes the cell, then re-run Stage 1.
- Edit `ChangeLog_*_Detail.csv`: delete or change the losing CR’s rows for that table + key + column, **and** set `resolved = Y` on the Conflicts row. Editing Detail alone does not clear the Conflicts sheet.

**`structural_collision`** — do not leave two structural CRs on the same table in one run. Merge the structural work into a **single** CR folder (one `before/` / `after/` pair), or exclude one CR, then re-run Stage 1. Mark `resolved = Y` only if you have already merged or sequenced the work by hand and accept applying the remaining Detail rows as-is.

**`missing_row_column_fill`** — supply the missing cell:

- Put the intended value in `after/` (so the covering `row_add` / `value_update` is generated) and re-run Stage 1, **or** add that cell as a row in `ChangeLog_*_Detail.csv` and set `resolved = Y` on the Conflicts row.
- If a **blank** cell is intentional, set `resolved = Y` on that Conflicts row without adding a value. Stage 2 will then write a blank at that intersection.

### Marking `resolved = Y`

On the Conflicts sheet, change `resolved` from `N` to `Y` (also accepted: `YES`, `TRUE`, `1`). Use this only after you have decided the outcome — accepted last-writer-wins in Control `order` (including same-value overlap, sequenced `row_add`/`column_add` then `value_update`, or differing independent updates), edited Detail, excluded CRs, merged folders, or confirmed a blank fill.

Then run `validate_only`, and only then `apply`.

---

## Outputs & audit

| Output | When |
|--------|------|
| `ChangeLog_*.xlsx` | Stage 1 (Summary / Conflicts / ReviewFiles index) |
| `ChangeLog_*_Detail.csv` | Stage 1 (Stage 2 source of truth) |
| `ChangeLog_*_reviews/*.xlsx` | Stage 1 (human review, split by first-level folder) |
| `IntegrationReport_*.xlsx` | Stage 2 |
| `New_Production_Tables/*.csv` | Stage 2 `apply` only |
| `Output/Audit/*.log` | Every run |

Each audit log records timestamp, mode, Control/Change Log hashes, CRs processed, warnings, and final status (`SUCCESS` / `FAILED` / `DRY_RUN_SUCCESS`).

---

## Safety rules (what will block you)

| Situation | Behaviour |
|-----------|-----------|
| Unresolved conflicts in Change Log | Hard stop in `apply` (see [Conflicts](#conflicts)) |
| `row_add` whose key already exists in the current table | Hard stop (including a second CR `row_add` of a key an earlier CR just added) |
| `value_update` / `row_delete` whose key is not in the current table | Hard stop (a follow-on update is valid only after an earlier CR has added that row, and Conflicts `resolved=Y` if they overlap) |
| Key-count change without approval | Hard stop |
| Column rename not listed in `ColumnRenames` | Hard stop |
| Production value ≠ Change Log `old_value` | Allowed — `new_value` is still applied |
| Table only in `before/` | Warning + skip (no auto-delete) |
| `validate_only` | Never writes new production tables |

Matching is **exact string** only. Re-running the same Change Log on the same production snapshot is designed to be **idempotent**.

---

## Typical quarterly checklist

1. Copy current production CSVs into `Production_Tables/`.
2. For each change request, drop baseline tables in `before/` and revised tables in `after/`.
3. Register each CR in Control (`include` / `approved` / `order`).
4. Declare any column renames; approve any key-count changes.
5. Run Stage 1 → review Summary, Conflicts, ReviewFiles, and the per-folder review workbooks. Resolve any Conflicts rows before Stage 2 (see [Conflicts](#conflicts)).
6. Run `validate_only` → fix any validation failures.
7. Run `apply` → take `New_Production_Tables/` as the candidate production set.
8. Keep the Change Log, Integration Report, and Audit logs with the quarter’s records.

---

## Running the acceptance tests

Optional self-check after install (uses pip/`requirements.txt`, including pytest):

```bash
python -m pip install -r requirements.txt
python fixtures/build_acceptance_fixtures.py
python -m pytest tests/ -q
# large-table timing (50k rows by default; override with PROPHET_TIMING_ROWS)
python -m pytest tests/test_large_table_timing.py -s
```

Detailed behaviour and acceptance cases (T01–T12) are documented in [`Prophet_Table_Change_Tool_Function_Doc.md`](Prophet_Table_Change_Tool_Function_Doc.md).

---

## Troubleshooting

| Problem | What to check |
|---------|----------------|
| `conda was not found` / Setup fails | Install Anaconda/Miniconda and ensure `conda` is on PATH |
| Env `prophet-table` missing | Run `Setup.bat` once |
| `Unknown mode` / wrong behaviour | `Config.mode` or pass `--mode` explicitly |
| no Change Log found in Output | Stage 2 needs a `ChangeLog_*.xlsx` in `Output\` (run Stage 1 first), or pass `--change-log` / type a path or filename when prompted |
| Change request not processed | Folder name = `change_request_id`; Stage 2 needs `include=Y` **and** `approved=Y` |
| Empty Change Log for a table | Confirm CSVs are under `before/` and `after/` with matching names |
| Apply refused after conflict | Open the Change Log **Conflicts** sheet and follow [Conflicts](#conflicts). `apply` stops until every remaining row is `resolved=Y` or no longer involves an included+approved CR |
| Rename / key-count hard stop | Fill `ColumnRenames` or `KeyCountApprovals` (or CR `approved`) |

For design-level detail (change types, conflict rules, developer checklist), see the function documentation linked above.
