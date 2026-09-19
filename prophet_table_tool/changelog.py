"""Stage 1 – Generate Change Log from before/ vs after/ tables."""

from __future__ import annotations

import logging
import time
from collections import Counter
from pathlib import Path

import polars as pl
from openpyxl import Workbook
from openpyxl.styles import Font

from .audit import close_audit_logger, setup_audit_logger
from .changelog_review import (
    TableReviewModel,
    add_contribution,
    group_table_reviews,
    review_group_for_table,
    sanitize_group_filename,
    write_review_workbook,
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
from .prophet_csv import discover_csv_tables, read_prophet_csv
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

        all_structural: list[ChangeRow] = []
        cell_frames: list[pl.DataFrame] = []
        summary_rows: list[dict] = []
        all_warnings: list[str] = []
        table_reviews: dict[str, TableReviewModel] = {}

        for cr in included:
            cr_dir = control.change_request_dir(cr.change_request_id)
            before_dir = cr_dir / "before"
            after_dir = cr_dir / "after"

            logger.info("[%s] discovering tables in before/ and after/", cr.change_request_id)
            before_tables = discover_csv_tables(before_dir)
            after_tables = discover_csv_tables(after_dir)
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

            only_before = set(before_tables) - set(after_tables)
            for table_name in sorted(only_before):
                msg = (
                    f"[{cr.change_request_id}] Table '{table_name}' exists only "
                    "in before/ — skipped (no automatic deletes)."
                )
                logger.warning(msg)
                all_warnings.append(msg)

            cr_structural: list[ChangeRow] = []
            cr_cells: list[pl.DataFrame] = []
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
                if diff.cell_changes.height:
                    cr_cells.append(diff.cell_changes)
                tables_touched |= diff.tables_touched
                logger.info(
                    "[%s] %s done, %d change row(s)",
                    cr.change_request_id,
                    table_name,
                    n_change,
                )
                if n_change or table_name in diff.tables_touched:
                    add_contribution(
                        table_reviews,
                        table_name,
                        cr.change_request_id,
                        before_tbl,
                        after_tbl,
                        diff.structural_rows,
                        cell_changes=diff.cell_changes,
                    )
                for w in diff.warnings:
                    logger.warning(w)
                    all_warnings.append(w)

            cr_cell_df = concat_cell_changes(cr_cells)
            all_structural.extend(cr_structural)
            if cr_cell_df.height:
                cell_frames.append(cr_cell_df)

            counts: Counter[str] = Counter(r.change_type for r in cr_structural)
            if cr_cell_df.height:
                vc = cr_cell_df.get_column("change_type").value_counts()
                for rec in vc.iter_rows():
                    counts[str(rec[0])] += int(rec[1])
            n_col = (
                counts.get("column_add", 0)
                + counts.get("column_delete", 0)
                + counts.get("column_rename", 0)
            )
            n_change_cr = len(cr_structural) + cr_cell_df.height
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

        all_cells = concat_cell_changes(cell_frames)
        conflicts = detect_conflicts(all_structural, all_cells)
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

        n_change_rows = len(all_structural) + all_cells.height
        t0 = time.perf_counter()
        logger.info("writing detail CSV (%d rows): %s", n_change_rows, detail_csv_path)
        _write_change_log_detail_csv(detail_csv_path, all_structural, all_cells)
        logger.info("wrote detail CSV in %.2fs", time.perf_counter() - t0)

        t0 = time.perf_counter()
        review_index = _write_folder_review_workbooks(
            control.output_path, run_id, table_reviews, conflicts, logger
        )
        logger.info("wrote review workbooks in %.2fs", time.perf_counter() - t0)

        t0 = time.perf_counter()
        logger.info(
            "writing Change Log index (Summary/Conflicts/ReviewFiles): %s", out_path
        )
        _write_change_log_excel(out_path, summary_rows, conflicts, review_index)
        logger.info("wrote Change Log index in %.2fs", time.perf_counter() - t0)

        tables_affected: set[str] = {r.table_name for r in all_structural}
        if all_cells.height:
            tables_affected.update(
                str(t) for t in all_cells.get_column("table_name").unique().to_list()
            )
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
        close_audit_logger(logger)


def _write_change_log_detail_csv(
    path: Path,
    structural: list[ChangeRow],
    cell_changes: pl.DataFrame,
) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    frames: list[pl.DataFrame] = []
    if structural:
        frames.append(rows_to_cell_df(structural).select(DETAIL_HEADERS))
    if cell_changes.height:
        frames.append(cell_changes.select(DETAIL_HEADERS))
    if frames:
        df = frames[0] if len(frames) == 1 else pl.concat(frames, how="vertical")
        df.write_csv(path, null_value="")
        return
    empty_cell_changes().select(DETAIL_HEADERS).write_csv(path, null_value="")


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


def _write_folder_review_workbooks(
    output_path: Path,
    run_id: str,
    table_reviews: dict[str, TableReviewModel],
    conflicts: list[dict],
    logger: logging.Logger,
) -> list[dict]:
    """Write per-folder review workbooks; return ReviewFiles index rows."""
    if not table_reviews:
        return []

    reviews_dir = output_path / f"ChangeLog_{run_id}_reviews"
    grouped = group_table_reviews(table_reviews)
    index_rows: list[dict] = []
    used_filenames: set[str] = set()

    for group in sorted(grouped):
        filename = _unique_group_filename(sanitize_group_filename(group), used_filenames)
        used_filenames.add(filename)
        dest = reviews_dir / filename
        group_reviews = grouped[group]
        group_conflicts = [
            c
            for c in conflicts
            if review_group_for_table(str(c.get("table_name", ""))) == group
        ]
        t0 = time.perf_counter()
        logger.info(
            "writing review workbook %s (%d table(s))", dest, len(group_reviews)
        )
        write_review_workbook(
            dest,
            group_reviews,
            conflict_headers=CONFLICT_HEADERS,
            conflict_rows=[_conflict_row_values(c) for c in group_conflicts],
        )
        logger.info("wrote %s in %.2fs", dest.name, time.perf_counter() - t0)
        rel = f"ChangeLog_{run_id}_reviews/{filename}"
        for table_name in sorted(group_reviews):
            index_rows.append(
                {
                    "table_name": table_name,
                    "group": group,
                    "review_file": rel,
                }
            )
    return index_rows


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
