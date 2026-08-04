# Prophet Table Change Tool

Consolidate quarterly Prophet table changes and integrate them safely into the production table set.

You drive the tool with an Excel **Control** workbook. The tool compares `before/` vs `after/` tables, writes an auditable Change Log, then validates and (optionally) applies those changes onto production CSVs.

---

## Requirements

- **Anaconda or Miniconda** — day-to-day users do not need a separate Python install
- Shared Conda env `prophet-table` from [`environment.yml`](environment.yml) (Python **3.11**, `polars`, `openpyxl`, `xlsxwriter`)
- [`requirements.txt`](requirements.txt) remains available for pip / pytest workflows (pytest is not in the day-to-day Conda env)

`Setup.bat` / `Run.bat` look for `conda` on PATH and in common install folders. If they still cannot find Conda, open **Anaconda Prompt**, `cd` to this project folder, and run the `.bat` from there.

---

## Quick start

1. **(First time)** Double-click [`Setup.bat`](Setup.bat) to create/update the `prophet-table` Conda env.
2. Set up a working folder (see [Folder layout](#folder-layout)) and fill in `Control.xlsx` (see [Control workbook](#control-workbook)).
3. Set `mode` in the Control `Config` sheet (`generate_changelog`, `validate_only`, or `apply`).
4. Double-click [`Run.bat`](Run.bat), or drag `Control.xlsx` onto it, and confirm the Control path if prompted.
5. For Stage 2 (`validate_only` / `apply`), ensure a Change Log already exists in `Output\` (from Stage 1). `Run.bat` auto-picks the newest `ChangeLog_*.xlsx` there.

Typical sequence: set `mode` to `generate_changelog` → run → review the Change Log → set `mode` to `validate_only` → run → set `mode` to `apply` → run.

### CLI alternative

From the project root, with the `prophet-table` env active (or via `conda run -n prophet-table`):

```bash
# Stage 1 – generate Change Log
python -m prophet_table_tool path\to\Control.xlsx --mode generate_changelog

# Stage 2 – dry run (no new tables written)
python -m prophet_table_tool path\to\Control.xlsx --mode validate_only --change-log path\to\Output\ChangeLog_YYYYMMDD_HHMMSS.xlsx

# Stage 2 – apply changes
python -m prophet_table_tool path\to\Control.xlsx --mode apply --change-log path\to\Output\ChangeLog_YYYYMMDD_HHMMSS.xlsx
```

`--mode` overrides the `mode` value in the Control `Config` sheet. If you omit `--mode` (as `Run.bat` does), the tool uses whatever is set in Control.

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

- `Output/ChangeLog_<run_id>.xlsx` — Summary, Conflicts, and per-table review sheets
- `Output/ChangeLog_<run_id>_Detail.csv` — machine-readable change rows used by Stage 2

Workbook sheets:

- **Summary** — per-CR counts and conflict flags
- **Conflicts** — overlapping updates across CRs (same table + key + column, structural collisions, or missing row×column fills)
- **One sheet per touched table** — human review only (changed rows); Stage 2 ignores these sheets

Value comparison is **numeric-aware** (`1.10` equals `1.1`). CSV files are read
with encoding fallback: utf-8-sig → utf-8 → cp1252 → latin-1.

```bash
python -m prophet_table_tool WorkingRoot\Control.xlsx --mode generate_changelog
```

**Review the Change Log before continuing.** Prefer the per-table review tabs for visual checks; use the Detail CSV / Conflicts for the machine-readable record. If the Conflicts sheet has rows, resolve them before applying.

### Stage 2 — Validate only (dry run)

Checks that the Change Log can be applied cleanly against current production:

- Referenced tables exist (except pure `table_add`)
- `value_update` / `row_delete` keys exist in production (`old_value` is not required to match)
- Column renames are declared; key-count changes are approved
- No unresolved conflicts

Writes `Output/IntegrationReport_<run_id>.xlsx` and an audit log. **Does not write** any files under `New_Production_Tables/`.

```bash
python -m prophet_table_tool WorkingRoot\Control.xlsx --mode validate_only --change-log WorkingRoot\Output\ChangeLog_YYYYMMDD_HHMMSS.xlsx
```

### Stage 2 — Apply

Same validation as above. If everything passes, writes updated CSVs to:

`Output/New_Production_Tables/`

Filenames and `!N` / `*` format are preserved. Production input files are not overwritten in place.

```bash
python -m prophet_table_tool WorkingRoot\Control.xlsx --mode apply --change-log WorkingRoot\Output\ChangeLog_YYYYMMDD_HHMMSS.xlsx
```

Stage 2 only processes change requests with **`include = Y` and `approved = Y`**, in `order` sequence.

---

## Outputs & audit

| Output | When |
|--------|------|
| `ChangeLog_*.xlsx` | Stage 1 |
| `IntegrationReport_*.xlsx` | Stage 2 |
| `New_Production_Tables/*.csv` | Stage 2 `apply` only |
| `Output/Audit/*.log` | Every run |

Each audit log records timestamp, mode, Control/Change Log hashes, CRs processed, warnings, and final status (`SUCCESS` / `FAILED` / `DRY_RUN_SUCCESS`).

---

## Safety rules (what will block you)

| Situation | Behaviour |
|-----------|-----------|
| Unresolved conflicts in Change Log | Hard stop in `apply` |
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
5. Run Stage 1 → review Summary, per-table tabs, Detail, and Conflicts.
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
```

Detailed behaviour and acceptance cases (T01–T12) are documented in [`Prophet_Table_Change_Tool_Function_Doc.md`](Prophet_Table_Change_Tool_Function_Doc.md).

---

## Troubleshooting

| Problem | What to check |
|---------|----------------|
| `conda was not found` / Setup fails | Install Anaconda/Miniconda and ensure `conda` is on PATH |
| Env `prophet-table` missing | Run `Setup.bat` once |
| `Unknown mode` / wrong behaviour | `Config.mode` or pass `--mode` explicitly |
| `--change-log is required` / no Change Log found | Stage 2 needs a `ChangeLog_*.xlsx` in `Output\` (run Stage 1 first); CLI needs `--change-log` |
| Change request not processed | Folder name = `change_request_id`; Stage 2 needs `include=Y` **and** `approved=Y` |
| Empty Change Log for a table | Confirm CSVs are under `before/` and `after/` with matching names |
| Apply refused after conflict | Open Change Log **Conflicts** sheet; for overlaps drop/edit the losing Detail rows or exclude a CR; for `missing_row_column_fill` add a Detail value (or set `resolved=Y` if blank is intentional) |
| Rename / key-count hard stop | Fill `ColumnRenames` or `KeyCountApprovals` (or CR `approved`) |

For design-level detail (change types, conflict rules, developer checklist), see the function documentation linked above.
