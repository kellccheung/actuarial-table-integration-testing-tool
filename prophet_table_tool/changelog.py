"""Stage 1 – Generate Change Log from before/ vs after/ tables."""

from __future__ import annotations

import logging
import re
import shutil
import time
from collections import Counter, defaultdict
from pathlib import Path

import polars as pl
from openpyxl import Workbook
from openpyxl.styles import Font

from .audit import close_audit_logger, setup_audit_logger
from .changelog_review import (
    ReviewWorkbookWriter,
    TableReviewModel,
    add_contribution,
    review_group_for_table,
    sanitize_group_filename,
)
from .control import ControlConfig, read_control
from .diff import (
    ChangeRow,
    concat_cell_changes,
    detect_conflicts,
    diff_table,
    empty_cell_changes,
    rows_to_cell_df,
)
from .prophet_csv import ProphetTable, read_prophet_csv, read_prophet_header, require_unique_keys
from .utils import file_hash, generate_run_id


DETAIL_HEADERS = [
    "change_request_id",
    "table_name",
    "change_type",
    "n_keys_before",
    "n_keys_after",
    "key_tuple",
    "column_name",
    "old_value",
    "new_value",
    "notes",
]

SUMMARY_HEADERS = [
    "change_request_id",
    "description",
    "tables_touched",
    "n_value_changes",
    "n_row_adds",
    "n_row_deletes",
    "n_column_changes",
    "n_key_count_changes",
    "has_conflict",
    "status",
]

CONFLICT_HEADERS = [
    "conflict_type",
    "table_name",
    "key_tuple",
    "column_name",
    "change_request_ids",
    "change_types",
    "old_values",
    "new_values",
    "notes",
    "resolved",
]

REVIEW_FILES_HEADERS = [
    "table_name",
    "group",
    "review_file",
]

SCAN_HEADERS = [
    "change_request_id",
    "table_name",
    "severity",
    "message",
]

_UNSAFE_PATH = re.compile(r'[<>:"|?*]')


class StructureScanError(ValueError):
    """Raised when generate_changelog stops on a fatal structure finding."""

    def __init__(self, path: Path, n_fatal: int) -> None:
        self.path = path
        self.n_fatal = n_fatal
        super().__init__(
            f"Structure scan found {n_fatal} fatal issue(s). See {path}"
        )


def generate_change_log(
    control_path: Path,
    *,
    control: ControlConfig | None = None,
) -> Path:
    """
    Stage 1: Compare before/ vs after/ for each included change request
    and write a ChangeLog_*.xlsx index, ChangeLog_*_Detail.csv sidecar,
    and per-folder review workbooks under ChangeLog_*_reviews/.

    Returns the path to the generated Change Log workbook.
    """
    control_path = Path(control_path)
    control = control or read_control(control_path)
    run_id = control.run_id or generate_run_id()
    control.output_path.mkdir(parents=True, exist_ok=True)

    logger, _log_path = setup_audit_logger(
        control.output_path, run_id, "generate_changelog"
    )
    status = "FAILED"
    out_path: Path | None = None

    try:
        logger.info("control_path=%s", control_path)
        logger.info("control_hash=%s", file_hash(control_path))
        logger.info("working_root=%s", control.working_root)

        included = control.included_requests(require_approved=False)
        logger.info(
            "change_requests_processed=%s",
            [r.change_request_id for r in included],
        )

        findings = scan_structure(control)
        scan_path = _write_structure_scan(control.output_path, run_id, findings)
        n_fatal = sum(1 for row in findings if row["severity"] == "FATAL")
        logger.info("structure_scan=%s fatal=%d findings=%d", scan_path, n_fatal, len(findings))
        if n_fatal:
            raise StructureScanError(scan_path, n_fatal)

        all_structural: list[ChangeRow] = []
        summary_rows: list[dict] = []
        all_warnings: list[str] = [
            row["message"] for row in findings if row["severity"] == "WARN"
        ]
        staging = control.output_path / f".stage1_{run_id}"
        cell_paths: list[Path] = []
        cell_by_table: dict[str, list[tuple[str, Path]]] = defaultdict(list)
        slice_by_table: dict[str, list[tuple[str, Path | None, Path | None]]] = defaultdict(list)
        struct_by_table: dict[str, list[ChangeRow]] = defaultdict(list)

        for cr in included:
            before_tables, after_tables, production_before = control.discover_cr_tables(cr)
            logger.info(
                "[%s] found %d before table(s), %d after table(s)",
                cr.change_request_id,
                len(before_tables),
                len(after_tables),
            )

            if not before_tables and not after_tables:
                msg = (
                    f"[{cr.change_request_id}] Empty change request folder "
                    f"(no CSVs in before/ or after/)."
                )
                logger.warning(msg)
                all_warnings.append(msg)
                summary_rows.append(
                    {
                        "change_request_id": cr.change_request_id,
                        "description": cr.description,
                        "tables_touched": "",
                        "n_value_changes": 0,
                        "n_row_adds": 0,
                        "n_row_deletes": 0,
                        "n_column_changes": 0,
                        "n_key_count_changes": 0,
                        "has_conflict": "N",
                        "status": "EMPTY",
                    }
                )
                continue

            only_before: set[str] = set()
            if not production_before:
                only_before = set(before_tables) - set(after_tables)
            for table_name in sorted(only_before):
                msg = (
                    f"[{cr.change_request_id}] Table '{table_name}' exists only "
                    "in before/ — skipped (no automatic deletes)."
                )
                logger.warning(msg)
                all_warnings.append(msg)

            cr_structural: list[ChangeRow] = []
            cr_cell_rows = 0
            cr_cell_counts: Counter[str] = Counter()
            tables_touched: set[str] = set()

            all_names = set(before_tables) | set(after_tables)
            for table_name in sorted(all_names):
                if table_name in only_before:
                    continue

                before_tbl = (
                    read_prophet_csv(before_tables[table_name])
                    if table_name in before_tables
                    else None
                )
                after_tbl = (
                    read_prophet_csv(after_tables[table_name])
                    if table_name in after_tables
                    else None
                )

                n_before = 0 if before_tbl is None else before_tbl.data.height
                n_after = 0 if after_tbl is None else after_tbl.data.height
                if before_tbl is not None and before_tbl.source_encoding:
                    logger.info(
                        "[%s] read before %s encoding=%s",
                        cr.change_request_id,
                        table_name,
                        before_tbl.source_encoding,
                    )
                if after_tbl is not None and after_tbl.source_encoding:
                    logger.info(
                        "[%s] read after %s encoding=%s",
                        cr.change_request_id,
                        table_name,
                        after_tbl.source_encoding,
                    )
                logger.info(
                    "[%s] comparing %s (%s→%s rows)...",
                    cr.change_request_id,
                    table_name,
                    n_before,
                    n_after,
                )

                diff = diff_table(
                    cr.change_request_id,
                    table_name,
                    before_tbl,
                    after_tbl,
                    control,
                )
                n_change = len(diff.structural_rows) + diff.cell_changes.height
                cr_structural.extend(diff.structural_rows)
                for row in diff.structural_rows:
                    struct_by_table[row.table_name].append(row)
                if diff.cell_changes.height:
                    cell_path = _stage_cell_path(staging, cr.change_request_id, table_name)
                    diff.cell_changes.select(DETAIL_HEADERS).write_parquet(cell_path)
                    cell_paths.append(cell_path)
                    cell_by_table[table_name].append((cr.change_request_id, cell_path))
                    cr_cell_rows += diff.cell_changes.height
                    for rec in diff.cell_changes.get_column("change_type").value_counts().iter_rows():
                        cr_cell_counts[str(rec[0])] += int(rec[1])
                tables_touched |= diff.tables_touched
                logger.info(
                    "[%s] %s done, %d change row(s)",
                    cr.change_request_id,
                    table_name,
                    n_change,
                )
                if n_change or table_name in diff.tables_touched:
                    before_side, after_side = _spill_review_slice(
                        staging,
                        cr.change_request_id,
                        table_name,
                        before_tbl,
                        after_tbl,
                        diff.structural_rows,
                        diff.cell_changes,
                    )
                    slice_by_table[table_name].append(
                        (cr.change_request_id, before_side, after_side)
                    )
                for w in diff.warnings:
                    logger.warning(w)
                    all_warnings.append(w)
                del before_tbl, after_tbl, diff

            all_structural.extend(cr_structural)
            counts: Counter[str] = Counter(r.change_type for r in cr_structural)
            counts.update(cr_cell_counts)
            n_col = (
                counts.get("column_add", 0)
                + counts.get("column_delete", 0)
                + counts.get("column_rename", 0)
            )
            n_change_cr = len(cr_structural) + cr_cell_rows
            summary_rows.append(
                {
                    "change_request_id": cr.change_request_id,
                    "description": cr.description,
                    "tables_touched": ", ".join(sorted(tables_touched)),
                    "n_value_changes": counts.get("value_update", 0),
                    "n_row_adds": counts.get("row_add", 0),
                    "n_row_deletes": counts.get("row_delete", 0),
                    "n_column_changes": n_col,
                    "n_key_count_changes": counts.get("key_count_change", 0),
                    "has_conflict": "N",
                    "status": "OK" if n_change_cr or tables_touched else "NO_CHANGES",
                }
            )

        conflicts = _conflicts_by_table(struct_by_table, cell_by_table)
        conflicted_crs: set[str] = set()
        for c in conflicts:
            conflicted_crs.update(c["change_request_ids"])

        for row in summary_rows:
            if row["change_request_id"] in conflicted_crs:
                row["has_conflict"] = "Y"
                if row["status"] == "OK":
                    row["status"] = "CONFLICT"

        out_path = control.output_path / f"ChangeLog_{run_id}.xlsx"
        detail_csv_path = control.output_path / f"ChangeLog_{run_id}_Detail.csv"

        t0 = time.perf_counter()
        logger.info("writing detail CSV: %s", detail_csv_path)
        n_change_rows = _write_change_log_detail_csv(
            detail_csv_path, all_structural, cell_paths
        )
        logger.info("wrote detail CSV in %.2fs", time.perf_counter() - t0)

        t0 = time.perf_counter()
        review_index = _write_reviews_from_spill(
            control.output_path,
            run_id,
            included_ids=[cr.change_request_id for cr in included],
            slice_by_table=slice_by_table,
            cell_by_table=cell_by_table,
            struct_by_table=struct_by_table,
            conflicts=conflicts,
            logger=logger,
        )
        logger.info("wrote review workbooks in %.2fs", time.perf_counter() - t0)

        t0 = time.perf_counter()
        logger.info(
            "writing Change Log index (Summary/Conflicts/ReviewFiles): %s", out_path
        )
        _write_change_log_excel(out_path, summary_rows, conflicts, review_index)
        logger.info("wrote Change Log index in %.2fs", time.perf_counter() - t0)

        tables_affected = set(struct_by_table) | set(cell_by_table) | set(slice_by_table)
        logger.info("tables_affected=%d", len(tables_affected))
        logger.info("n_change_rows=%d", n_change_rows)
        logger.info("n_conflicts=%d", len(conflicts))
        logger.info(
            "n_review_files=%d", len({row["review_file"] for row in review_index})
        )
        for w in all_warnings:
            logger.info("warning=%s", w)
        logger.info("change_log_path=%s", out_path)
        logger.info("change_log_hash=%s", file_hash(out_path))
        logger.info("change_log_detail_path=%s", detail_csv_path)
        logger.info("change_log_detail_hash=%s", file_hash(detail_csv_path))

        status = "SUCCESS"
        logger.info("final_status=%s", status)
        return out_path

    except Exception as exc:
        logger.exception("generate_change_log failed: %s", exc)
        logger.info("final_status=%s", status)
        raise
    finally:
        if "staging" in locals() and staging.is_dir():
            shutil.rmtree(staging, ignore_errors=True)
        close_audit_logger(logger)


def _write_change_log_detail_csv(
    path: Path,
    structural: list[ChangeRow],
    cell_paths: list[Path],
) -> int:
    """Write structural rows, then each spilled cell file. Returns data-row count."""
    path.parent.mkdir(parents=True, exist_ok=True)
    n_rows = 0
    header = True
    if structural:
        frame = rows_to_cell_df(structural).select(DETAIL_HEADERS)
        _append_detail_csv(path, frame, header=header)
        n_rows += frame.height
        header = False
    for cell_path in cell_paths:
        frame = pl.read_parquet(cell_path).select(DETAIL_HEADERS)
        _append_detail_csv(path, frame, header=header)
        n_rows += frame.height
        header = False
        del frame
    if header:
        empty_cell_changes().select(DETAIL_HEADERS).write_csv(path, null_value="")
    return n_rows


def _append_detail_csv(path: Path, frame: pl.DataFrame, *, header: bool) -> None:
    if header:
        frame.write_csv(path, null_value="")
        return
    with path.open("a", encoding="utf-8", newline="") as handle:
        frame.write_csv(handle, include_header=False, null_value="")


def _conflict_row_values(conflict: dict) -> list:
    return [
        conflict.get("conflict_type", ""),
        conflict.get("table_name", ""),
        conflict.get("key_tuple", ""),
        conflict.get("column_name", ""),
        ", ".join(conflict.get("change_request_ids", [])),
        ", ".join(conflict.get("change_types", [])),
        ", ".join(conflict.get("old_values", [])),
        ", ".join(conflict.get("new_values", [])),
        conflict.get("notes", ""),
        conflict.get("resolved", "N"),
    ]


def _unique_group_filename(base: str, used: set[str]) -> str:
    candidate = f"{base}.xlsx"
    if candidate not in used:
        return candidate
    for i in range(2, 1000):
        candidate = f"{base}_{i}.xlsx"
        if candidate not in used:
            return candidate
    raise ValueError(f"Unable to allocate unique review filename for group {base!r}")


def run_structure_scan(control: ControlConfig) -> tuple[Path, list[dict]]:
    """Write StructureScan_<run_id>.xlsx and return ``(path, findings)``."""
    control.output_path.mkdir(parents=True, exist_ok=True)
    findings = scan_structure(control, check_duplicate_keys=True)
    path = _write_structure_scan(
        control.output_path, control.run_id or "scan", findings
    )
    return path, findings


def scan_structure(
    control: ControlConfig, *, check_duplicate_keys: bool = False
) -> list[dict]:
    """Header check of each included change request. Does not diff cell values.

    ``check_duplicate_keys`` is scan mode only, so changelog generation does not
    read each table twice.
    """
    findings: list[dict] = []
    for cr in control.included_requests(require_approved=False):
        try:
            before_dir, after_dir, production = control.cr_locations(cr)
        except ValueError as exc:
            _finding(findings, cr.change_request_id, "", "FATAL", str(exc))
            continue
        if cr.after_path.strip() and not after_dir.is_dir():
            _finding(
                findings,
                cr.change_request_id,
                "",
                "FATAL",
                f"after_path does not exist: {after_dir}",
            )
        if (
            not production
            and cr.before_path.strip()
            and before_dir is not None
            and not before_dir.is_dir()
        ):
            _finding(
                findings,
                cr.change_request_id,
                "",
                "FATAL",
                f"before_path does not exist: {before_dir}",
            )
        try:
            before_tables, after_tables, production = control.discover_cr_tables(cr)
        except ValueError as exc:
            _finding(findings, cr.change_request_id, "", "FATAL", str(exc))
            continue
        if not before_tables and not after_tables:
            _finding(
                findings,
                cr.change_request_id,
                "",
                "WARN",
                f"[{cr.change_request_id}] No table files for this change request.",
            )
            continue
        if not production:
            for table_name in sorted(set(before_tables) - set(after_tables)):
                _scan_header(
                    findings,
                    cr.change_request_id,
                    table_name,
                    before_tables[table_name],
                    check_duplicate_keys=check_duplicate_keys,
                )
                _finding(
                    findings,
                    cr.change_request_id,
                    table_name,
                    "WARN",
                    f"[{cr.change_request_id}] Table '{table_name}' exists only "
                    "in before/ — skipped (no automatic deletes).",
                )
        for table_name in sorted(set(before_tables) & set(after_tables)):
            _scan_pair(
                findings,
                control,
                cr.change_request_id,
                table_name,
                before_tables[table_name],
                after_tables[table_name],
                check_duplicate_keys=check_duplicate_keys,
            )
        for table_name in sorted(set(after_tables) - set(before_tables)):
            _scan_header(
                findings,
                cr.change_request_id,
                table_name,
                after_tables[table_name],
                check_duplicate_keys=check_duplicate_keys,
            )
    return findings


def _finding(
    findings: list[dict],
    change_request_id: str,
    table_name: str,
    severity: str,
    message: str,
) -> None:
    findings.append(
        {
            "change_request_id": change_request_id,
            "table_name": table_name,
            "severity": severity,
            "message": message,
        }
    )


def _scan_header(
    findings: list[dict],
    change_request_id: str,
    table_name: str,
    path: Path,
    *,
    check_duplicate_keys: bool = False,
) -> tuple[int, list[str]] | None:
    try:
        n_keys, columns = read_prophet_header(path)
    except ValueError as exc:
        _finding(findings, change_request_id, table_name, "FATAL", str(exc))
        return None
    duplicates = _duplicate_names(columns)
    if duplicates:
        # Skip the body read: repeated names collapse in the reader schema.
        _finding(
            findings,
            change_request_id,
            table_name,
            "FATAL",
            f"[{change_request_id}] Table '{table_name}' has duplicate column names: "
            f"{duplicates}",
        )
    elif check_duplicate_keys:
        try:
            table = read_prophet_csv(path)
            require_unique_keys(
                table, f"[{change_request_id}] Table '{table_name}' {path}"
            )
        except ValueError as exc:
            _finding(findings, change_request_id, table_name, "FATAL", str(exc))
    return n_keys, columns


def _scan_pair(
    findings: list[dict],
    control: ControlConfig,
    change_request_id: str,
    table_name: str,
    before_path: Path,
    after_path: Path,
    *,
    check_duplicate_keys: bool = False,
) -> None:
    before = _scan_header(
        findings,
        change_request_id,
        table_name,
        before_path,
        check_duplicate_keys=check_duplicate_keys,
    )
    after = _scan_header(
        findings,
        change_request_id,
        table_name,
        after_path,
        check_duplicate_keys=check_duplicate_keys,
    )
    if before is None or after is None:
        return
    before_n, before_cols = before
    after_n, after_cols = after
    pairs = control.key_column_pairs(change_request_id, table_name)
    if pairs:
        both = [(b, a) for b, a in pairs if b and a]
        if not both:
            _finding(
                findings,
                change_request_id,
                table_name,
                "FATAL",
                f"[{change_request_id}] Table '{table_name}' KeyCountApprovals rows "
                "need a before_column and after_column pair to match rows.",
            )
            return
        if len({b for b, _a in both}) != len(both) or len({a for _b, a in both}) != len(both):
            _finding(
                findings,
                change_request_id,
                table_name,
                "FATAL",
                f"[{change_request_id}] Table '{table_name}' KeyCountApprovals "
                "repeats a before_column or after_column.",
            )
        missing = [b for b, _a in both if b not in before_cols] + [
            a for _b, a in both if a not in after_cols
        ]
        if missing:
            _finding(
                findings,
                change_request_id,
                table_name,
                "FATAL",
                f"[{change_request_id}] Table '{table_name}' KeyCountApprovals "
                f"column(s) {missing} are not on that side of the table.",
            )
        return
    renames = {
        row.old_column_name: row.new_column_name
        for row in control.renames_for(change_request_id, table_name)
    }
    before_mapped = [renames.get(col, col) for col in before_cols]
    after_keys = after_cols[: max(after_n - 1, 0)]
    if before_n == after_n:
        missing_keys = [col for col in after_keys if col not in before_mapped]
        if missing_keys:
            _finding(
                findings,
                change_request_id,
                table_name,
                "FATAL",
                f"[{change_request_id}] Table '{table_name}' key column(s) {missing_keys} "
                "differ between before and after. Declare before_column and "
                "after_column on KeyCountApprovals.",
            )
    else:
        shared = [col for col in after_keys if col in before_mapped]
        if not shared:
            _finding(
                findings,
                change_request_id,
                table_name,
                "FATAL",
                f"[{change_request_id}] Table '{table_name}' has no shared key column "
                f"({before_n} -> {after_n}). Declare before_column and "
                "after_column on KeyCountApprovals.",
            )
        else:
            _finding(
                findings,
                change_request_id,
                table_name,
                "WARN",
                f"[{change_request_id}] Table '{table_name}' !N changed "
                f"{before_n} -> {after_n}; shared key names still match.",
            )
    after_set = set(after_cols)
    before_set = set(before_mapped)
    for col in sorted(after_set - before_set):
        _finding(
            findings,
            change_request_id,
            table_name,
            "WARN",
            f"[{change_request_id}] Table '{table_name}' column '{col}' is only on the after table.",
        )
    for col in before_cols:
        mapped = renames.get(col, col)
        if mapped not in after_set and col not in after_set:
            _finding(
                findings,
                change_request_id,
                table_name,
                "WARN",
                f"[{change_request_id}] Table '{table_name}' column '{col}' is only on the before table.",
            )


def _duplicate_names(columns: list[str]) -> list[str]:
    seen: set[str] = set()
    duplicates: list[str] = []
    for name in columns:
        if name in seen and name not in duplicates:
            duplicates.append(name)
        seen.add(name)
    return duplicates


def _write_structure_scan(output_path: Path, run_id: str, findings: list[dict]) -> Path:
    path = output_path / f"StructureScan_{run_id}.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Structure"
    _write_header(ws, SCAN_HEADERS)
    for i, row in enumerate(findings, start=2):
        for j, key in enumerate(SCAN_HEADERS, start=1):
            ws.cell(i, j, row.get(key, ""))
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def _safe_part(value: str) -> str:
    cleaned = value.replace("\\", "__").replace("/", "__")
    cleaned = _UNSAFE_PATH.sub("_", cleaned).strip(" .")
    return cleaned or "_"


def _stage_cell_path(staging: Path, change_request_id: str, table_name: str) -> Path:
    path = staging / "cells" / _safe_part(change_request_id) / f"{_safe_part(table_name)}.parquet"
    path.parent.mkdir(parents=True, exist_ok=True)
    return path


def _spill_review_slice(
    staging: Path,
    change_request_id: str,
    table_name: str,
    before: ProphetTable | None,
    after: ProphetTable | None,
    structural_rows: list[ChangeRow],
    cell_changes: pl.DataFrame,
) -> tuple[Path | None, Path | None]:
    """Write the changed-key slice the review sheet needs. Returns side paths."""
    scratch: dict[str, TableReviewModel] = {}
    add_contribution(
        scratch,
        table_name,
        change_request_id,
        before,
        after,
        structural_rows,
        cell_changes=cell_changes,
    )
    model = scratch.get(table_name)
    if model is None or not model.contributions:
        return None, None
    contrib = model.contributions[-1]
    folder = staging / "slices" / _safe_part(change_request_id)
    stem = _safe_part(table_name)
    before_path = _write_side(folder / f"{stem}__before.parquet", contrib.before)
    after_path = _write_side(folder / f"{stem}__after.parquet", contrib.after)
    return before_path, after_path


def _write_side(path: Path, table: ProphetTable | None) -> Path | None:
    if table is None:
        return None
    path.parent.mkdir(parents=True, exist_ok=True)
    table.data.select(table.columns).write_parquet(path)
    path.with_suffix(".nkeys").write_text(str(table.n_keys), encoding="utf-8")
    return path


def _read_side(path: Path | None) -> ProphetTable | None:
    if path is None or not path.is_file():
        return None
    data = pl.read_parquet(path).cast(pl.Utf8).fill_null("")
    n_keys = int(path.with_suffix(".nkeys").read_text(encoding="utf-8"))
    return ProphetTable(n_keys=n_keys, columns=list(data.columns), data=data)


def _conflicts_by_table(
    struct_by_table: dict[str, list[ChangeRow]],
    cell_by_table: dict[str, list[tuple[str, Path]]],
) -> list[dict]:
    conflicts: list[dict] = []
    for table_name in sorted(set(struct_by_table) | set(cell_by_table)):
        frames = [pl.read_parquet(path) for _cr, path in cell_by_table.get(table_name, [])]
        cells = concat_cell_changes(frames)
        conflicts.extend(detect_conflicts(struct_by_table.get(table_name, []), cells))
        del frames, cells
    return conflicts


def _write_reviews_from_spill(
    output_path: Path,
    run_id: str,
    *,
    included_ids: list[str],
    slice_by_table: dict[str, list[tuple[str, Path | None, Path | None]]],
    cell_by_table: dict[str, list[tuple[str, Path]]],
    struct_by_table: dict[str, list[ChangeRow]],
    conflicts: list[dict],
    logger: logging.Logger,
) -> list[dict]:
    tables = sorted(
        set(slice_by_table) | set(cell_by_table) | set(struct_by_table),
        key=lambda name: (review_group_for_table(name), name),
    )
    if not tables:
        return []
    conflicts_by_group: dict[str, list[dict]] = defaultdict(list)
    for conflict in conflicts:
        group = review_group_for_table(str(conflict.get("table_name", "")))
        conflicts_by_group[group].append(conflict)

    reviews_dir = output_path / f"ChangeLog_{run_id}_reviews"
    writers: dict[str, tuple[str, ReviewWorkbookWriter]] = {}
    used_filenames: set[str] = set()
    index_rows: list[dict] = []
    try:
        for table_name in tables:
            model = _review_model_for_table(
                table_name,
                included_ids,
                slice_by_table.get(table_name, []),
                cell_by_table.get(table_name, []),
                struct_by_table.get(table_name, []),
            )
            if model is None:
                continue
            group = review_group_for_table(table_name)
            if group not in writers:
                filename = _unique_group_filename(
                    sanitize_group_filename(group), used_filenames
                )
                used_filenames.add(filename)
                writers[group] = (
                    filename,
                    ReviewWorkbookWriter(
                        reviews_dir / filename,
                        CONFLICT_HEADERS,
                        [_conflict_row_values(row) for row in conflicts_by_group.get(group, [])],
                    ),
                )
            filename, writer = writers[group]
            writer.write_table(model)
            index_rows.append(
                {
                    "table_name": table_name,
                    "group": group,
                    "review_file": f"ChangeLog_{run_id}_reviews/{filename}",
                }
            )
            del model
    finally:
        for _filename, writer in writers.values():
            writer.close()
    logger.info("wrote %d review workbook(s)", len(writers))
    return index_rows


def _review_model_for_table(
    table_name: str,
    included_ids: list[str],
    slices: list[tuple[str, Path | None, Path | None]],
    cells: list[tuple[str, Path]],
    structural: list[ChangeRow],
) -> TableReviewModel | None:
    slice_by_cr = {cr_id: (before, after) for cr_id, before, after in slices}
    cell_by_cr = {cr_id: path for cr_id, path in cells}
    struct_by_cr: dict[str, list[ChangeRow]] = defaultdict(list)
    for row in structural:
        struct_by_cr[row.change_request_id].append(row)
    present = set(slice_by_cr) | set(cell_by_cr) | set(struct_by_cr)
    reviews: dict[str, TableReviewModel] = {}
    for cr_id in included_ids:
        if cr_id not in present:
            continue
        before_path, after_path = slice_by_cr.get(cr_id, (None, None))
        cell_path = cell_by_cr.get(cr_id)
        cell_df = pl.read_parquet(cell_path) if cell_path is not None else empty_cell_changes()
        add_contribution(
            reviews,
            table_name,
            cr_id,
            _read_side(before_path),
            _read_side(after_path),
            struct_by_cr.get(cr_id, []),
            cell_changes=cell_df,
        )
    return reviews.get(table_name)


def _write_change_log_excel(
    path: Path,
    summary_rows: list[dict],
    conflicts: list[dict],
    review_index: list[dict] | None = None,
) -> None:
    """Write slim workbook: Summary + Conflicts + ReviewFiles (no Detail/review tabs)."""
    wb = Workbook()

    ws = wb.active
    ws.title = "Summary"
    _write_header(ws, SUMMARY_HEADERS)
    for i, row in enumerate(summary_rows, start=2):
        for j, key in enumerate(SUMMARY_HEADERS, start=1):
            ws.cell(i, j, row.get(key, ""))

    ws_conf = wb.create_sheet("Conflicts")
    _write_header(ws_conf, CONFLICT_HEADERS)
    for i, c in enumerate(conflicts, start=2):
        for j, val in enumerate(_conflict_row_values(c), start=1):
            ws_conf.cell(i, j, val)

    ws_idx = wb.create_sheet("ReviewFiles")
    _write_header(ws_idx, REVIEW_FILES_HEADERS)
    for i, row in enumerate(review_index or [], start=2):
        for j, key in enumerate(REVIEW_FILES_HEADERS, start=1):
            ws_idx.cell(i, j, row.get(key, ""))

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def _write_header(ws, headers: list[str]) -> None:
    bold = Font(bold=True)
    for j, h in enumerate(headers, start=1):
        cell = ws.cell(1, j, h)
        cell.font = bold
