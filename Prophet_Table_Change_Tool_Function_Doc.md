# Prophet Table Change Consolidation & Integration Tool
**Function Documentation v1.0**  
**Target Language:** Python 3.11+  
**Key Libraries:** `polars`, `openpyxl` / `xlsxwriter`, `pathlib`, `logging`, `hashlib`, `datetime`

---

## 1. Purpose

Automate the consolidation of quarterly actuarial Prophet table changes and their safe integration into the production table set.

**Two main stages:**

1. **Stage 1 – Generate Change Log**  
   Compare `before/` vs `after/` tables for each change request and produce a structured, auditable Change Log.

2. **Stage 2 – Validate & Integrate**  
   Apply the Change Log (in controlled order) onto the current production tables, with full validation, conflict checking, dry-run support, and audit trail.

Excel is the primary human interface (Control workbook).

---

## 2. CSV Table Format (Prophet-style)

Every table follows this exact structure:

```csv
!3,Age,Duration,Product,Rate,Loading
*,20,1,PROD_A,0.0012,1.05
*,25,1,PROD_A,0.0015,1.05
...
```

**Rules:**
- Row 1, Column 1: `!N` where `N` = number of key columns (**including** the first column itself).
- Row 1, Columns 2 onwards: actual column headers.
- Data rows: first cell is always `*`, followed by key values and data values.
- Key columns = first `N` columns (the marker column + the next `N-1` columns).
- All comparisons and matching use **exact string match** (no floating-point tolerance).

When writing output tables, the exact `!N` + `*` format must be preserved.

---

## 3. Folder Structure

```
WorkingRoot/
├── Control.xlsx
├── ChangeRequests/
│   ├── <change_request_id>/          ← folder name = change_request_id
│   │   ├── before/
│   │   │   └── *.csv
│   │   └── after/
│   │       └── *.csv
│   └── ...
├── Production_Tables/                ← current production snapshot
│   └── *.csv
└── Output/
    ├── ChangeLog_YYYYMMDD_HHMM.xlsx
    ├── IntegrationReport_YYYYMMDD_HHMM.xlsx
    ├── New_Production_Tables/        ← final output tables (original filenames)
    └── Audit/
        └── *.log
```

- `change_request_id` = the folder name itself.
- Output table filenames must keep the **original** names.

---

## 4. Control.xlsx Specification

### Sheet: `Config`
| Key                        | Example / Value                          | Description |
|---------------------------|------------------------------------------|-----------|
| `working_root`            | `D:\ProphetChanges\2026Q3`              | Root path |
| `mode`                    | `generate_changelog` / `validate_only` / `apply` | Run mode |
| `production_tables_path`  | `Production_Tables`                     | Relative or absolute |
| `output_path`             | `Output`                                | |
| `run_id`                  | (auto-generated)                        | |

### Sheet: `ChangeRequests`
| change_request_id                  | order | include | approved | description                  | notes |
|------------------------------------|-------|---------|----------|------------------------------|-------|
| CR_2026Q3_01_MortalityUpdate      | 1     | Y       | Y        | Mortality assumption update |       |
| CR_2026Q3_02_ExpenseRevision      | 2     | Y       | Y        | Expense loading revision    |       |

- `order`: integer – determines application sequence in Stage 2.
- `include`: `Y` / `N`
- `approved`: `Y` / `N` (required for key-count changes and major structural changes)

### Sheet: `ColumnRenames` (explicit declaration only)
| change_request_id             | table_name   | old_column_name | new_column_name |
|-------------------------------|--------------|-----------------|-----------------|
| CR_2026Q3_01_MortalityUpdate | MORT_TABLE   | OldRate         | Rate            |

Column renames are **never** auto-detected. They must be declared here.

### Sheet: `KeyCountApprovals` (optional helper)
| change_request_id             | table_name   | approved |
|-------------------------------|--------------|----------|
| CR_2026Q3_03_StructureChange  | EXPENSE_TABLE| Y        |

---

## 5. Stage 1 – Generate Change Log

**Function signature (conceptual):**
```python
def generate_change_log(control_path: Path) -> Path:
    """
    Returns path to the generated ChangeLog_*.xlsx
    """
```

**Logic:**
1. Read `Control.xlsx` → list of included change requests sorted by `order`.
2. For each change request folder:
   - Discover all CSV files in `before/` and `after/`.
   - Tables present in both → compute detailed diff.
   - Tables only in `after/` → record as full table add (`table_add`).
   - Tables only in `before/` → raise warning, record in Summary, **do not** generate any change rows for that table.
3. Detect conflicts across change requests (same table + same key combination + same column, or structural collision).
4. Write multi-sheet Change Log Excel.

**Change Log Sheets:**

**Summary**
- change_request_id, description, tables_touched, n_value_changes, n_row_adds, n_row_deletes, n_column_changes, n_key_count_changes, has_conflict, status

**ChangeLog_Detail**
| change_request_id | table_name | change_type | n_keys_before | n_keys_after | key_tuple | column_name | old_value | new_value | notes |

**Supported `change_type` values:**
- `value_update`
- `row_add`
- `row_delete`
- `column_add`
- `column_delete`
- `column_rename` (only if declared in Control)
- `key_count_change`
- `table_add` (table only exists in after)

**Conflicts**
- Populated only when overlaps are detected.
- User must resolve conflicts (by editing ChangeLog_Detail or marking resolution) before Stage 2 can proceed in `apply` mode.

**Per-table review sheets** (one sheet per touched table name, e.g. `MORT_TABLE`)
- Human review only — Stage 2 does **not** read these sheets.
- Wide Prophet-style layout: `_change` column (`ADD` / `DELETE` / blank) plus table columns.
- Value changes shown in-place as `old -> new` (with `[change_request_id]` when multiple CRs touch the same table).
- Structural changes (`column_add`, `column_delete`, `column_rename`, `key_count_change`, `table_add`) listed as notes above the table header.
- Unchanged rows included for context so the sheet still looks like the real table.

---

## 6. Stage 2 – Validate & Integrate

**Function signature (conceptual):**
```python
def integrate_changes(
    control_path: Path,
    change_log_path: Path,
    mode: Literal["validate_only", "apply"]
) -> Path:
    """
    Returns path to IntegrationReport_*.xlsx
    """
```

**Logic:**
1. Load production tables.
2. Load Change Log and filter only `include=Y` + `approved=Y` change requests, sorted by `order`.
3. Pre-validate:
   - All referenced tables exist in production (except pure `table_add`).
   - For every `value_update` / `row_delete`: exact key + old_value must match current production.
   - For `key_count_change`: only proceed if approved in Control.
   - For `column_rename`: only apply if declared in `ColumnRenames`.
   - No unresolved conflicts.
4. If any validation fails → stop and write detailed Validation_Report (never write tables in `validate_only` or on failure).
5. If `mode == "apply"` and validation passes:
   - Apply changes sequentially in the order defined in Control.
   - Write new CSV files to `Output/New_Production_Tables/` using **original filenames**.
   - Preserve exact `!N` + `*` format.
6. Always produce Integration Report + timestamped audit log.

---

## 7. Audit Trail

Every run writes a `.log` file containing:
- Timestamp (ISO format)
- Run mode
- Control file hash
- Change Log hash
- List of change requests processed
- Number of tables affected
- Any warnings / conflicts / validation failures
- Final status (`SUCCESS` / `FAILED` / `DRY_RUN_SUCCESS`)

---

## 8. Error Handling & Safety Rules

- Exact string match only (no tolerance).
- Tables only in `before` → warning + skip (never delete automatically).
- Unresolved conflicts → hard stop in `apply` mode.
- Key count change without approval → hard stop.
- Column rename without explicit declaration → hard stop.
- Always support pure dry-run (`validate_only`).
- Idempotent: running the same Change Log twice on the same production snapshot produces identical results.

---

## 9. Recommended Implementation Notes

- Use **Polars** for all table reading, joining, and diffing.
- Parse the special first row carefully; never treat the `!N` or `*` rows as normal data.
- Represent keys as a tuple (or struct) of the first `N` columns for reliable matching.
- Keep all intermediate results as Polars DataFrames until final Excel writing.
- Use `openpyxl` for reading Control and writing formatted Change Log / Reports.
- Generate `run_id` as `YYYYMMDD_HHMMSS`.

---

## 10. Developer Checklist & Acceptance Test Cases

### Developer Checklist (Must Pass Before Handover)

- [ ] Correctly parses Prophet CSV format (`!N` on first row, `*` on data rows). `N` includes the marker column.
- [ ] Key matching uses the first `N` columns as a composite key (exact string match only).
- [ ] Stage 1 correctly handles:
  - Value updates
  - Row add / delete
  - Column add / delete
  - Explicit column renames (from Control sheet only)
  - Key-count changes
  - Tables only in `after` → `table_add`
  - Tables only in `before` → warning + skip (no changes generated)
- [ ] Conflict detection works across multiple change requests (same table + same key + same column, or structural collisions).
- [ ] Stage 2 respects `order` column and only processes `include = Y` + `approved = Y` rows.
- [ ] `validate_only` mode never writes any output tables.
- [ ] `apply` mode writes new CSVs with **original filenames** and preserves exact `!N` / `*` format.
- [ ] Key-count change is blocked unless approved in Control.
- [ ] Column rename is blocked unless explicitly declared in the `ColumnRenames` sheet.
- [ ] Audit log is written for every run (timestamp, mode, hashes, status).
- [ ] All intermediate processing uses Polars; Excel I/O only at the boundaries.
- [ ] Control.xlsx is the single source of truth for configuration and sequencing.

### Core Acceptance Test Cases

| ID  | Scenario                                      | Expected Result |
|-----|-----------------------------------------------|-----------------|
| T01 | Two change requests, no overlapping cells     | Change Log generated cleanly; both applied in order; new tables match expected |
| T02 | Two change requests update the **same cell**  | Conflict detected and written to Conflicts sheet; `apply` mode refuses to run |
| T03 | Table exists only in `before/`                | Warning recorded; no change rows generated for that table |
| T04 | Table exists only in `after/`                 | Recorded as `table_add`; appears in new production set |
| T05 | Column rename declared in Control             | Applied correctly; old column name disappears, new name appears |
| T06 | Column rename **not** declared                | Hard stop with clear error |
| T07 | Key-count change (`!N` different) + approved = Y | Applied successfully |
| T08 | Key-count change + approved = N               | Hard stop |
| T09 | `validate_only` mode                          | Full validation report produced; **zero** files written to `New_Production_Tables/` |
| T10 | Value in production does not match `old_value` in Change Log | Validation fails; clear message identifying the mismatched key |
| T11 | Run the same Change Log twice on identical production snapshot | Identical output tables (idempotent) |
| T12 | Empty change request folder (no CSVs)         | Graceful handling; no crash; recorded in Summary |

### Recommended Manual Spot-Checks

- Open the generated Change Log and visually confirm a few `value_update` and `row_add` rows.
- Diff one output table against a manually prepared expected file.
- Confirm the audit log contains the Control file hash and Change Log hash.

---

**End of Function Documentation v1.0**
