"""Stage 2 – Validate & Integrate Change Log onto production tables."""

from __future__ import annotations

import csv
import shutil
from collections import defaultdict
from pathlib import Path
from typing import Literal

import polars as pl
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

from .audit import close_audit_logger, setup_audit_logger
from .control import ControlConfig, read_control
from .diff import ChangeRow
from .prophet_csv import (
    ProphetTable,
    discover_csv_tables,
    key_str_expr,
    read_prophet_csv,
    write_prophet_csv,
)
from .utils import file_hash, generate_run_id, yn_to_bool


REPORT_HEADERS = [
    "section",
    "change_request_id",
    "table_name",
    "change_type",
    "key_tuple",
    "column_name",
    "message",
    "severity",
]


def integrate_changes(
    control_path: Path,
    change_log_path: Path,
    mode: Literal["validate_only", "apply"],
    *,
    control: ControlConfig | None = None,
) -> Path:
    """
    Stage 2: Validate (and optionally apply) a Change Log onto production tables.

    Touched tables are loaded one at a time. CRs run in Control ``order`` against
    that table; ``apply`` spills each finished table to a staging folder and
    publishes to ``New_Production_Tables/`` only if every table passed.

    Returns the path to IntegrationReport_*.xlsx.
    """
    if mode not in {"validate_only", "apply"}:
        raise ValueError(f"mode must be 'validate_only' or 'apply', got {mode!r}")

    control_path = Path(control_path)
    change_log_path = Path(change_log_path)
    control = control or read_control(control_path)
    run_id = control.run_id or generate_run_id()
    control.output_path.mkdir(parents=True, exist_ok=True)

    logger, _log_path = setup_audit_logger(control.output_path, run_id, mode)
    status = "FAILED"
    report_path = control.output_path / f"IntegrationReport_{run_id}.xlsx"
    validation_messages: list[dict] = []
    staging_dir = control.output_path / f".stage2_{run_id}"

    try:
        logger.info("control_path=%s", control_path)
        logger.info("control_hash=%s", file_hash(control_path))
        logger.info("change_log_path=%s", change_log_path)
        logger.info("change_log_hash=%s", file_hash(change_log_path))
        logger.info("production_tables_path=%s", control.production_tables_path)

        approved_crs = control.included_requests(require_approved=True)
        cr_order = {r.change_request_id: r.order for r in approved_crs}
        cr_ids = {r.change_request_id for r in approved_crs}
        logger.info("change_requests_processed=%s", [r.change_request_id for r in approved_crs])

        detail_rows, conflicts = _load_change_log(change_log_path)
        detail_rows = [r for r in detail_rows if r.change_request_id in cr_ids]
        detail_rows.sort(
            key=lambda r: (
                cr_order.get(r.change_request_id, 10**9),
                r.table_name,
                r.change_type,
                r.key_tuple,
                r.column_name,
            )
        )
        logger.info("loaded %d detail change row(s)", len(detail_rows))

        unresolved = [
            c
            for c in conflicts
            if not yn_to_bool(c.get("resolved", "N"))
            and any(cid in cr_ids for cid in c.get("change_request_ids", []))
        ]
        if unresolved:
            for c in unresolved:
                msg = (
                    f"Unresolved conflict ({c.get('conflict_type')}) on table "
                    f"{c.get('table_name')} key={c.get('key_tuple')!r} "
                    f"col={c.get('column_name')!r} CRs={c.get('change_request_ids')}"
                )
                validation_messages.append(
                    _vmsg("conflict", "", c.get("table_name", ""), "", "", "", msg, "FAIL")
                )
                logger.error(msg)
            if mode == "apply":
                status = "FAILED"
                _write_integration_report(report_path, validation_messages, status, mode)
                logger.info("final_status=%s", status)
                return report_path

        logger.info("discovering production tables in %s", control.production_tables_path)
        prod_paths = discover_csv_tables(control.production_tables_path)
        touched = {r.table_name for r in detail_rows}
        logger.info(
            "discovered %d production table(s); processing %d touched one at a time",
            len(prod_paths),
            len(touched),
        )

        by_table_cr: dict[str, dict[str, list[ChangeRow]]] = defaultdict(lambda: defaultdict(list))
        for row in detail_rows:
            by_table_cr[row.table_name][row.change_request_id].append(row)

        ok = True
        staged: list[tuple[str, str, int]] = []
        staged_names: set[str] = set()
        n_touched = len(touched)
        for i, table_name in enumerate(sorted(touched), start=1):
            logger.info(
                "loading production table %s (%d/%d)",
                table_name,
                i,
                n_touched,
            )
            tables: dict[str, ProphetTable] = {}
            path = prod_paths.get(table_name)
            if path is not None:
                tables[table_name] = read_prophet_csv(path)

            cr_rows = by_table_cr.get(table_name, {})
            for cr in approved_crs:
                rows = cr_rows.get(cr.change_request_id)
                if not rows:
                    continue
                logger.info(
                    "[%s] validating %s (%d change row(s))...",
                    cr.change_request_id,
                    table_name,
                    len(rows),
                )
                table_ok, msgs = _validate_table_changes(
                    control, cr.change_request_id, table_name, rows, tables
                )
                validation_messages.extend(msgs)
                if not table_ok:
                    ok = False
                    logger.warning(
                        "[%s] validation FAILED for %s; skipped in-memory apply",
                        cr.change_request_id,
                        table_name,
                    )
                    continue
                logger.info("[%s] validation OK for %s", cr.change_request_id, table_name)
                tables[table_name] = _apply_table_changes(
                    control, cr.change_request_id, table_name, rows, tables
                )
                staged.append((cr.change_request_id, table_name, len(rows)))
                logger.info(
                    "[%s] staged %s in-memory (%d change row(s))",
                    cr.change_request_id,
                    table_name,
                    len(rows),
                )

            if mode == "apply" and ok and table_name in tables:
                staging_dir.mkdir(parents=True, exist_ok=True)
                table = tables[table_name]
                staged_path = staging_dir / Path(table_name).with_suffix(table.source_suffix)
                staged_path.parent.mkdir(parents=True, exist_ok=True)
                logger.info("writing staging %s -> %s", table_name, staged_path)
                write_prophet_csv(table, staged_path)
                staged_names.add(table_name)
            del tables

        if not ok:
            status = "FAILED"
            _write_integration_report(report_path, validation_messages, status, mode)
            logger.info(
                "n_validation_failures=%d",
                sum(1 for m in validation_messages if m["severity"] == "FAIL"),
            )
            logger.info("final_status=%s", status)
            return report_path

        if mode == "validate_only":
            status = "DRY_RUN_SUCCESS"
            validation_messages.append(
                _vmsg(
                    "summary",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "Validation passed (dry-run); no tables written.",
                    "INFO",
                )
            )
            _write_integration_report(report_path, validation_messages, status, mode)
            logger.info("tables_affected=%d", len(touched))
            logger.info("final_status=%s", status)
            return report_path

        for cr_id, table_name, n_rows in staged:
            validation_messages.append(
                _vmsg(
                    "apply",
                    cr_id,
                    table_name,
                    "",
                    "",
                    "",
                    f"Applied {n_rows} change row(s).",
                    "INFO",
                )
            )
            logger.info("[%s] applied %s", cr_id, table_name)

        out_dir = control.output_path / "New_Production_Tables"
        out_dir.mkdir(parents=True, exist_ok=True)

        written = 0
        written_names: set[str] = set()
        if staging_dir.is_dir():
            for src in staging_dir.rglob("*"):
                if not src.is_file():
                    continue
                rel = src.relative_to(staging_dir)
                out_path = out_dir / rel
                out_path.parent.mkdir(parents=True, exist_ok=True)
                logger.info("publishing %s -> %s", rel.as_posix(), out_path)
                shutil.copy2(src, out_path)
                written += 1
                written_names.add(rel.with_suffix("").as_posix())

        for name in staged_names:
            written_names.add(name)

        for name, src in prod_paths.items():
            if name in written_names:
                continue
            out_path = out_dir / Path(name).with_suffix(src.suffix)
            out_path.parent.mkdir(parents=True, exist_ok=True)
            logger.info("copying unchanged %s -> %s", name, out_path)
            shutil.copy2(src, out_path)
            written += 1

        validation_messages.append(
            _vmsg("summary", "", "", "", "", "", f"Wrote {written} table(s) to {out_dir}", "INFO")
        )
        status = "SUCCESS"
        _write_integration_report(report_path, validation_messages, status, mode)
        logger.info("tables_affected=%d", written)
        logger.info("output_dir=%s", out_dir)
        logger.info("final_status=%s", status)
        return report_path

    except Exception as exc:
        logger.exception("integrate_changes failed: %s", exc)
        validation_messages.append(
            _vmsg("error", "", "", "", "", "", str(exc), "FAIL")
        )
        try:
            _write_integration_report(report_path, validation_messages, status, mode)
        except Exception:
            pass
        logger.info("final_status=%s", status)
        raise
    finally:
        _remove_stage2_dir(staging_dir)
        close_audit_logger(logger)


def _remove_stage2_dir(path: Path) -> None:
    """Delete a Stage 2 staging folder if it exists (apply spill / crash cleanup)."""
    if path.is_dir():
        shutil.rmtree(path, ignore_errors=True)


def _vmsg(
    section: str,
    cr: str,
    table: str,
    change_type: str,
    key_tuple: str,
    column: str,
    message: str,
    severity: str,
) -> dict:
    return {
        "section": section,
        "change_request_id": cr,
        "table_name": table,
        "change_type": change_type,
        "key_tuple": key_tuple,
        "column_name": column,
        "message": message,
        "severity": severity,
    }


def _detail_csv_path_for(change_log_path: Path) -> Path:
    """Sidecar path: ChangeLog_{run_id}.xlsx → ChangeLog_{run_id}_Detail.csv."""
    return change_log_path.with_name(f"{change_log_path.stem}_Detail.csv")


def _opt_int(value: object) -> int | None:
    if value is None:
        return None
    s = str(value).strip()
    if s == "":
        return None
    try:
        return int(float(s))
    except (TypeError, ValueError):
        return None


def _change_row_from_dict(d: dict) -> ChangeRow:
    return ChangeRow(
        change_request_id=str(d.get("change_request_id") or "").strip(),
        table_name=str(d.get("table_name") or "").strip(),
        change_type=str(d.get("change_type") or "").strip(),
        n_keys_before=_opt_int(d.get("n_keys_before")),
        n_keys_after=_opt_int(d.get("n_keys_after")),
        key_tuple=str(d.get("key_tuple") or ""),
        column_name=str(d.get("column_name") or ""),
        old_value="" if d.get("old_value") is None else str(d.get("old_value")),
        new_value="" if d.get("new_value") is None else str(d.get("new_value")),
        notes=str(d.get("notes") or ""),
    )


def _iter_sheet_dicts(ws) -> list[dict]:
    rows = ws.iter_rows(values_only=True)
    header = next(rows, None)
    if not header:
        return []
    keys = [str(h).strip() if h else f"col{i}" for i, h in enumerate(header)]
    out: list[dict] = []
    for row in rows:
        if not row or all(c is None for c in row):
            continue
        out.append({keys[i]: (row[i] if i < len(row) else None) for i in range(len(keys))})
    return out


def _load_change_log(path: Path) -> tuple[list[ChangeRow], list[dict]]:
    detail_csv = _detail_csv_path_for(path)
    wb = load_workbook(path, data_only=True, read_only=True)
    try:
        if detail_csv.is_file():
            detail = _load_detail_csv(detail_csv)
        else:
            detail = []
            if "ChangeLog_Detail" in wb.sheetnames:
                for d in _iter_sheet_dicts(wb["ChangeLog_Detail"]):
                    if not str(d.get("change_request_id") or "").strip():
                        continue
                    detail.append(_change_row_from_dict(d))

        conflicts: list[dict] = []
        if "Conflicts" in wb.sheetnames:
            for d in _iter_sheet_dicts(wb["Conflicts"]):
                cr_raw = str(d.get("change_request_ids") or "")
                conflicts.append(
                    {
                        "conflict_type": str(d.get("conflict_type") or ""),
                        "table_name": str(d.get("table_name") or ""),
                        "key_tuple": str(d.get("key_tuple") or ""),
                        "column_name": str(d.get("column_name") or ""),
                        "change_request_ids": [x.strip() for x in cr_raw.split(",") if x.strip()],
                        "resolved": d.get("resolved", "N"),
                        "notes": str(d.get("notes") or ""),
                    }
                )
    finally:
        wb.close()
    return detail, conflicts


def _load_detail_csv(path: Path) -> list[ChangeRow]:
    detail: list[ChangeRow] = []
    with path.open("r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        for d in reader:
            if not d or not (d.get("change_request_id") or "").strip():
                continue
            detail.append(_change_row_from_dict(d))
    return detail


def _validate_table_changes(
    control: ControlConfig,
    cr_id: str,
    table_name: str,
    rows: list[ChangeRow],
    tables: dict[str, ProphetTable],
) -> tuple[bool, list[dict]]:
    msgs: list[dict] = []
    ok = True
    types = {r.change_type for r in rows}
    is_table_add = "table_add" in types

    if not is_table_add and table_name not in tables:
        ok = False
        msgs.append(
            _vmsg(
                "validation",
                cr_id,
                table_name,
                "",
                "",
                "",
                f"Table '{table_name}' not found in production (and not a table_add).",
                "FAIL",
            )
        )
        return ok, msgs

    for r in rows:
        if r.change_type == "key_count_change":
            if not control.is_key_count_approved(cr_id, table_name):
                ok = False
                msgs.append(
                    _vmsg(
                        "validation",
                        cr_id,
                        table_name,
                        "key_count_change",
                        "",
                        "",
                        "Key-count change blocked: not approved in Control.",
                        "FAIL",
                    )
                )

    declared = control.renames_for(cr_id, table_name)
    for r in rows:
        if r.change_type != "column_rename":
            continue
        if not any(d.old_column_name == r.old_value and d.new_column_name == r.new_value for d in declared):
            ok = False
            msgs.append(
                _vmsg(
                    "validation",
                    cr_id,
                    table_name,
                    "column_rename",
                    "",
                    r.column_name,
                    f"Column rename {r.old_value!r}->{r.new_value!r} not declared in ColumnRenames.",
                    "FAIL",
                )
            )

    if is_table_add or table_name not in tables:
        return ok, msgs

    table = tables[table_name]
    existing = with_key_str(table).select("_key_str")

    need_exist = [r for r in rows if r.change_type in {"value_update", "row_delete"}]
    if need_exist:
        chk = pl.DataFrame({"_key_str": [r.key_tuple for r in need_exist]}).unique()
        missing = set(chk.join(existing, on="_key_str", how="anti").get_column("_key_str").to_list())
        if missing:
            ok = False
            seen: set[tuple[str, str]] = set()
            for r in need_exist:
                marker = (r.key_tuple, r.change_type)
                if r.key_tuple not in missing or marker in seen:
                    continue
                seen.add(marker)
                msgs.append(
                    _vmsg(
                        "validation",
                        cr_id,
                        table_name,
                        r.change_type,
                        r.key_tuple,
                        r.column_name,
                        f"Key not found in current table for {r.change_type}.",
                        "FAIL",
                    )
                )

    add_keys = list(dict.fromkeys(r.key_tuple for r in rows if r.change_type == "row_add" and r.key_tuple))
    if add_keys:
        found = set(
            pl.DataFrame({"_key_str": add_keys})
            .join(existing, on="_key_str", how="inner")
            .get_column("_key_str")
            .to_list()
        )
        if found:
            ok = False
            seen_add: set[str] = set()
            for r in rows:
                if r.change_type != "row_add" or r.key_tuple not in found or r.key_tuple in seen_add:
                    continue
                seen_add.add(r.key_tuple)
                msgs.append(
                    _vmsg(
                        "validation",
                        cr_id,
                        table_name,
                        "row_add",
                        r.key_tuple,
                        r.column_name,
                        "Key already exists in current table for row_add.",
                        "FAIL",
                    )
                )

    if ok:
        msgs.append(
            _vmsg(
                "validation",
                cr_id,
                table_name,
                "",
                "",
                "",
                "Validation passed.",
                "INFO",
            )
        )
    return ok, msgs


def with_key_str(table: ProphetTable) -> pl.DataFrame:
    return table.data.with_columns(key_str_expr(table.key_columns).alias("_key_str"))


def _group_by_type(rows: list[ChangeRow]) -> dict[str, list[ChangeRow]]:
    grouped: dict[str, list[ChangeRow]] = defaultdict(list)
    for r in rows:
        grouped[r.change_type].append(r)
    return grouped


def _pivot_cells(rows: list[ChangeRow], value_attr: str) -> pl.DataFrame:
    if not rows:
        return pl.DataFrame()
    df = pl.DataFrame(
        {
            "key_tuple": [r.key_tuple for r in rows],
            "column_name": [r.column_name for r in rows],
            "value": [getattr(r, value_attr) for r in rows],
        }
    ).filter(pl.col("column_name") != "")
    if df.is_empty():
        return df
    return df.pivot(
        on="column_name",
        index="key_tuple",
        values="value",
        aggregate_function="last",
    )


def _list_get(parts: pl.Expr, i: int) -> pl.Expr:
    return pl.when(parts.list.len() > i).then(parts.list.get(i)).otherwise(pl.lit(None))


def _fill_keys_from_tuple(
    wide: pl.DataFrame,
    match_key_cols: list[str],
    final_key_cols: list[str],
) -> pl.DataFrame:
    if "key_tuple" not in wide.columns:
        return wide
    wide = wide.with_columns(pl.col("key_tuple").str.split("|").alias("_parts"))
    use_final = pl.col("_parts").list.len() == len(final_key_cols)
    for i, kc in enumerate(match_key_cols):
        part = _list_get(pl.col("_parts"), i)
        expr = pl.when(~use_final).then(part).otherwise(pl.lit(None))
        if kc in wide.columns:
            wide = wide.with_columns(pl.coalesce([pl.col(kc), expr]).alias(kc))
        else:
            wide = wide.with_columns(expr.cast(pl.Utf8).fill_null("").alias(kc))
    for i, kc in enumerate(final_key_cols):
        part = _list_get(pl.col("_parts"), i)
        expr = pl.when(use_final).then(part).otherwise(pl.lit(None))
        if kc in wide.columns:
            wide = wide.with_columns(pl.coalesce([pl.col(kc), expr]).alias(kc))
        else:
            wide = wide.with_columns(expr.cast(pl.Utf8).fill_null("").alias(kc))
    return wide.drop("_parts")


def _apply_table_changes(
    control: ControlConfig,
    cr_id: str,
    table_name: str,
    rows: list[ChangeRow],
    tables: dict[str, ProphetTable],
) -> ProphetTable:
    """Apply all change rows for one CR+table; return the updated ProphetTable."""
    by_type = _group_by_type(rows)

    if "table_add" in by_type:
        n_keys = 1
        for r in rows:
            if r.n_keys_after is not None:
                n_keys = r.n_keys_after
                break
        add_rows = by_type.get("row_add", [])
        col_order = list(dict.fromkeys(r.column_name for r in add_rows if r.column_name))
        wide = _pivot_cells(add_rows, "new_value")
        if wide.height:
            for c in col_order:
                if c not in wide.columns:
                    wide = wide.with_columns(pl.lit("").cast(pl.Utf8).alias(c))
            if "key_tuple" in wide.columns:
                wide = wide.sort("key_tuple")
            data = wide.select(col_order).cast(pl.Utf8).fill_null("")
        else:
            data = pl.DataFrame({c: [] for c in col_order}).cast(pl.Utf8)
        return ProphetTable(n_keys=n_keys, columns=col_order, data=data)

    table = tables[table_name]
    original_n_keys = table.n_keys
    n_keys = table.n_keys
    columns = list(table.columns)
    data = table.data
    leading_dummy_lines = list(table.leading_dummy_lines)
    trailing_dummy_lines = list(table.trailing_dummy_lines)
    source_path = table.source_path

    for r in by_type.get("key_count_change", []):
        if r.n_keys_after is not None:
            n_keys = r.n_keys_after

    for r in by_type.get("column_rename", []):
        old_c, new_c = r.old_value, r.new_value
        if old_c in columns:
            columns = [new_c if c == old_c else c for c in columns]
            data = data.rename({old_c: new_c})

    for r in by_type.get("column_add", []):
        col = r.column_name
        if col and col not in columns:
            if n_keys > original_n_keys:
                insert_at = max(original_n_keys - 1, 0)
                insert_at = min(insert_at, len(columns))
                columns.insert(insert_at, col)
            else:
                columns.append(col)
            data = data.with_columns(pl.lit("").cast(pl.Utf8).alias(col))
            data = data.select(columns)

    for r in by_type.get("column_delete", []):
        col = r.column_name
        if col in columns:
            columns = [c for c in columns if c != col]
            data = data.drop(col)

    match_key_cols = columns[: max(original_n_keys - 1, 0)]
    final_key_cols = columns[: max(n_keys - 1, 0)]
    data = data.with_columns(key_str_expr(match_key_cols).alias("_key_str"))

    delete_keys = {r.key_tuple for r in by_type.get("row_delete", [])}
    if delete_keys:
        data = data.filter(~pl.col("_key_str").is_in(list(delete_keys)))

    updates = by_type.get("value_update", [])
    if updates:
        by_key: dict[str, dict[str, str]] = defaultdict(dict)
        for r in updates:
            if r.column_name:
                by_key[r.key_tuple][r.column_name] = r.new_value
        update_cols = sorted({c for cv in by_key.values() for c in cv})
        for col in update_cols:
            if col not in data.columns:
                data = data.with_columns(pl.lit("").cast(pl.Utf8).alias(col))
            if col not in columns:
                columns.append(col)
        records = [{"_key_str": k, **{c: cv.get(c) for c in update_cols}} for k, cv in by_key.items()]
        upd = pl.DataFrame(records).with_columns(
            [pl.col(c).cast(pl.Utf8) for c in update_cols]
        )
        joined = data.join(upd, on="_key_str", how="left", suffix="_new")
        coalesced = [
            pl.coalesce([pl.col(f"{c}_new"), pl.col(c)]).alias(c)
            if f"{c}_new" in joined.columns
            else pl.col(c)
            for c in update_cols
        ]
        drop_cols = [f"{c}_new" for c in update_cols if f"{c}_new" in joined.columns]
        data = joined.with_columns(coalesced).drop(drop_cols)

    add_rows = by_type.get("row_add", [])
    if add_rows:
        wide = _pivot_cells(add_rows, "new_value")
        if wide.height:
            wide = _fill_keys_from_tuple(wide, match_key_cols, final_key_cols)
            for c in columns:
                if c not in wide.columns:
                    wide = wide.with_columns(pl.lit("").cast(pl.Utf8).alias(c))
            extra = [c for c in wide.columns if c not in columns and c not in {"key_tuple", "_key_str"}]
            columns.extend(extra)
            for c in columns:
                if c not in data.columns and c != "_key_str":
                    data = data.with_columns(pl.lit("").cast(pl.Utf8).alias(c))
            new_df = wide.select(columns).cast(pl.Utf8).fill_null("")
            new_df = new_df.with_columns(key_str_expr(match_key_cols).alias("_key_str"))
            data = pl.concat(
                [
                    data.select(columns + ["_key_str"]),
                    new_df.select(columns + ["_key_str"]),
                ],
                how="vertical",
            )

    data = data.select(columns)
    return ProphetTable(
        n_keys=n_keys,
        columns=columns,
        data=data,
        source_path=source_path,
        leading_dummy_lines=leading_dummy_lines,
        trailing_dummy_lines=trailing_dummy_lines,
    )


def _write_integration_report(
    path: Path,
    messages: list[dict],
    status: str,
    mode: str,
) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Validation_Report"
    bold = Font(bold=True)
    headers = REPORT_HEADERS + ["status", "mode"]
    for j, h in enumerate(headers, start=1):
        cell = ws.cell(1, j, h)
        cell.font = bold

    for i, m in enumerate(messages, start=2):
        for j, key in enumerate(REPORT_HEADERS, start=1):
            ws.cell(i, j, m.get(key, ""))
        ws.cell(i, len(REPORT_HEADERS) + 1, status)
        ws.cell(i, len(REPORT_HEADERS) + 2, mode)

    ws2 = wb.create_sheet("Summary")
    ws2["A1"] = "status"
    ws2["B1"] = status
    ws2["A2"] = "mode"
    ws2["B2"] = mode
    ws2["A3"] = "n_messages"
    ws2["B3"] = len(messages)
    ws2["A4"] = "n_failures"
    ws2["B4"] = sum(1 for m in messages if m.get("severity") == "FAIL")

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
