"""Table-level diffing between before/ and after/ Prophet CSVs."""

from __future__ import annotations

from dataclasses import dataclass, field

import polars as pl

from .control import ControlConfig
from .prophet_csv import ProphetTable, with_normalized_keys
from .utils import format_normalized_key


CHANGE_TYPES = (
    "value_update",
    "row_add",
    "row_delete",
    "column_add",
    "column_delete",
    "column_rename",
    "key_count_change",
    "table_add",
)


@dataclass
class ChangeRow:
    change_request_id: str
    table_name: str
    change_type: str
    n_keys_before: int | None
    n_keys_after: int | None
    key_tuple: str  # pipe-joined key values for Excel readability
    column_name: str
    old_value: str
    new_value: str
    notes: str = ""


@dataclass
class DiffResult:
    change_rows: list[ChangeRow] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    tables_touched: set[str] = field(default_factory=set)


def format_key_tuple(values: list | tuple | None) -> str:
    """Pipe-join key parts with numeric normalization (same as Stage 1/2 keys)."""
    return format_normalized_key(values)


def _cells_differ_expr(old_col: str, new_col: str) -> pl.Expr:
    """True when old/new are not equal as strings and not equal as floats."""
    old_s = pl.col(old_col).cast(pl.Utf8).fill_null("")
    new_s = pl.col(new_col).cast(pl.Utf8).fill_null("")
    old_f = old_s.cast(pl.Float64, strict=False)
    new_f = new_s.cast(pl.Float64, strict=False)
    both_numeric = old_f.is_not_null() & new_f.is_not_null() & (old_s != "") & (new_s != "")
    numeric_eq = both_numeric & (old_f == new_f)
    string_eq = old_s == new_s
    return ~(string_eq | numeric_eq)


def diff_table(
    change_request_id: str,
    table_name: str,
    before: ProphetTable | None,
    after: ProphetTable | None,
    control: ControlConfig,
) -> DiffResult:
    """
    Compute detailed changes between before and after for one table.

    Value comparison is numeric-aware (``1.10`` == ``1.1``). Keys are joined
    after the same numeric normalization.
    """
    result = DiffResult()

    if before is None and after is None:
        return result

    if before is None and after is not None:
        result.tables_touched.add(table_name)
        result.change_rows.append(
            ChangeRow(
                change_request_id=change_request_id,
                table_name=table_name,
                change_type="table_add",
                n_keys_before=None,
                n_keys_after=after.n_keys,
                key_tuple="",
                column_name="",
                old_value="",
                new_value="",
                notes="Table only present in after/",
            )
        )
        _emit_row_cell_changes(
            result,
            change_request_id=change_request_id,
            table_name=table_name,
            change_type="row_add",
            n_keys_before=None,
            n_keys_after=after.n_keys,
            keyed=with_normalized_keys(after.data, after.key_columns),
            columns=after.columns,
            value_from="new",
            notes="Part of table_add",
        )
        return result

    if before is not None and after is None:
        msg = (
            f"[{change_request_id}] Table '{table_name}' exists only in before/ — "
            "skipped (no automatic deletes)."
        )
        result.warnings.append(msg)
        return result

    assert before is not None and after is not None
    result.tables_touched.add(table_name)

    # --- Key-count change ---
    if before.n_keys != after.n_keys:
        result.change_rows.append(
            ChangeRow(
                change_request_id=change_request_id,
                table_name=table_name,
                change_type="key_count_change",
                n_keys_before=before.n_keys,
                n_keys_after=after.n_keys,
                key_tuple="",
                column_name="",
                old_value=str(before.n_keys),
                new_value=str(after.n_keys),
                notes="Structural key-count change",
            )
        )

    # --- Column renames (explicit declaration only) ---
    declared = control.renames_for(change_request_id, table_name)
    rename_old_to_new = {r.old_column_name: r.new_column_name for r in declared}
    rename_new_to_old = {v: k for k, v in rename_old_to_new.items()}

    for old_c, new_c in rename_old_to_new.items():
        result.change_rows.append(
            ChangeRow(
                change_request_id=change_request_id,
                table_name=table_name,
                change_type="column_rename",
                n_keys_before=before.n_keys,
                n_keys_after=after.n_keys,
                key_tuple="",
                column_name=f"{old_c}->{new_c}",
                old_value=old_c,
                new_value=new_c,
                notes="Declared in ColumnRenames",
            )
        )

    before_cols_mapped = [rename_old_to_new.get(c, c) for c in before.columns]
    after_cols = list(after.columns)

    before_set = set(before_cols_mapped)
    after_set = set(after_cols)

    for col in sorted(after_set - before_set):
        if col in rename_new_to_old:
            continue
        result.change_rows.append(
            ChangeRow(
                change_request_id=change_request_id,
                table_name=table_name,
                change_type="column_add",
                n_keys_before=before.n_keys,
                n_keys_after=after.n_keys,
                key_tuple="",
                column_name=col,
                old_value="",
                new_value="",
                notes="",
            )
        )

    deleted_cols = [
        c
        for c in before.columns
        if c not in rename_old_to_new and rename_old_to_new.get(c, c) not in after_set
    ]
    for col in sorted(deleted_cols):
        result.change_rows.append(
            ChangeRow(
                change_request_id=change_request_id,
                table_name=table_name,
                change_type="column_delete",
                n_keys_before=before.n_keys,
                n_keys_after=after.n_keys,
                key_tuple="",
                column_name=col,
                old_value="",
                new_value="",
                notes="",
            )
        )

    # Align before schema to after names via renames for value comparison
    before_aligned = before.data
    if rename_old_to_new:
        before_aligned = before_aligned.rename(
            {old: new for old, new in rename_old_to_new.items() if old in before_aligned.columns}
        )

    if before.n_keys == after.n_keys:
        key_cols = after.key_columns
    else:
        key_cols = [c for c in after.key_columns if c in before_aligned.columns]
        if not key_cols:
            return result

    common_value_cols = [
        c for c in after.columns if c in before_aligned.columns and c not in key_cols
    ]

    before_df = before_aligned
    after_df = after.data

    before_keep = [c for c in key_cols + common_value_cols if c in before_df.columns]
    after_extra = [
        c
        for c in after.columns
        if c not in before_aligned.columns and c not in key_cols
    ]
    after_keep = [
        c
        for c in key_cols + common_value_cols + after_extra
        if c in after_df.columns
    ]
    after_keep = list(dict.fromkeys(after_keep))

    b = with_normalized_keys(before_df.select(before_keep), key_cols)
    a = with_normalized_keys(after_df.select(after_keep), key_cols)

    b_keys = set(b["_key_str"].to_list())
    a_keys = set(a["_key_str"].to_list())

    added_keys = a_keys - b_keys
    deleted_keys = b_keys - a_keys
    common_keys = b_keys & a_keys

    if added_keys:
        a_add = a.filter(pl.col("_key_str").is_in(list(added_keys)))
        _emit_row_cell_changes(
            result,
            change_request_id=change_request_id,
            table_name=table_name,
            change_type="row_add",
            n_keys_before=before.n_keys,
            n_keys_after=after.n_keys,
            keyed=a_add,
            columns=after.columns,
            value_from="new",
            notes="",
        )

    if deleted_keys:
        b_del = b.filter(pl.col("_key_str").is_in(list(deleted_keys)))
        _emit_row_cell_changes(
            result,
            change_request_id=change_request_id,
            table_name=table_name,
            change_type="row_delete",
            n_keys_before=before.n_keys,
            n_keys_after=after.n_keys,
            keyed=b_del,
            columns=list(before_aligned.columns),
            value_from="old",
            notes="",
        )

    # Value updates on common keys (vectorized numeric-aware compare)
    if common_keys and common_value_cols:
        b_common = b.filter(pl.col("_key_str").is_in(list(common_keys)))
        a_common = a.filter(pl.col("_key_str").is_in(list(common_keys)))

        joined = b_common.select(
            ["_key_str"] + [pl.col(c).alias(f"old_{c}") for c in common_value_cols]
        ).join(
            a_common.select(
                ["_key_str"] + [pl.col(c).alias(f"new_{c}") for c in common_value_cols]
            ),
            on="_key_str",
            how="inner",
        )

        for col in common_value_cols:
            old_alias = f"old_{col}"
            new_alias = f"new_{col}"
            diffs = joined.filter(_cells_differ_expr(old_alias, new_alias)).select(
                [
                    pl.col("_key_str"),
                    pl.col(old_alias).cast(pl.Utf8).fill_null("").alias("old_value"),
                    pl.col(new_alias).cast(pl.Utf8).fill_null("").alias("new_value"),
                ]
            )
            if diffs.is_empty():
                continue
            for key_str, old_s, new_s in diffs.iter_rows():
                result.change_rows.append(
                    ChangeRow(
                        change_request_id=change_request_id,
                        table_name=table_name,
                        change_type="value_update",
                        n_keys_before=before.n_keys,
                        n_keys_after=after.n_keys,
                        key_tuple=key_str,
                        column_name=col,
                        old_value=old_s,
                        new_value=new_s,
                        notes="",
                    )
                )

    genuinely_new_cols = [c for c in after.columns if c not in before_aligned.columns]
    if genuinely_new_cols and common_keys:
        a_common = a.filter(pl.col("_key_str").is_in(list(common_keys)))
        _emit_row_cell_changes(
            result,
            change_request_id=change_request_id,
            table_name=table_name,
            change_type="value_update",
            n_keys_before=before.n_keys,
            n_keys_after=after.n_keys,
            keyed=a_common,
            columns=genuinely_new_cols,
            value_from="new",
            notes="Value on newly added column",
            force_old_empty=True,
        )

    return result


def _emit_row_cell_changes(
    result: DiffResult,
    *,
    change_request_id: str,
    table_name: str,
    change_type: str,
    n_keys_before: int | None,
    n_keys_after: int | None,
    keyed: pl.DataFrame,
    columns: list[str],
    value_from: str,
    notes: str,
    force_old_empty: bool = False,
) -> None:
    """Emit one ChangeRow per cell for row_add / row_delete / new-column fills."""
    present_cols = [c for c in columns if c in keyed.columns]
    if not present_cols or keyed.is_empty():
        return

    for row in keyed.select(["_key_str"] + present_cols).iter_rows(named=True):
        key_str = row["_key_str"]
        for col in present_cols:
            cell = row[col]
            val = "" if cell is None else str(cell)
            if force_old_empty:
                old_v, new_v = "", val
            elif value_from == "new":
                old_v, new_v = "", val
            else:
                old_v, new_v = val, ""
            result.change_rows.append(
                ChangeRow(
                    change_request_id=change_request_id,
                    table_name=table_name,
                    change_type=change_type,
                    n_keys_before=n_keys_before,
                    n_keys_after=n_keys_after,
                    key_tuple=key_str,
                    column_name=col,
                    old_value=old_v,
                    new_value=new_v,
                    notes=notes,
                )
            )


def detect_conflicts(change_rows: list[ChangeRow]) -> list[dict]:
    """
    Detect conflicts across change requests:
    same table + same key + same column with differing changes,
    structural collisions (key_count / column changes) on the same table
    from different CRs, or missing fills where a row_add and column_add
    leave their intersection cell without a Change Log value.
    """
    conflicts: list[dict] = []

    cell_index: dict[tuple[str, str, str], list[ChangeRow]] = {}
    for row in change_rows:
        if row.change_type not in {"value_update", "row_add", "row_delete"}:
            continue
        if not row.column_name:
            continue
        key = (row.table_name, row.key_tuple, row.column_name)
        cell_index.setdefault(key, []).append(row)

    for (table, key_tuple, column), rows in cell_index.items():
        cr_ids = {r.change_request_id for r in rows}
        if len(cr_ids) < 2:
            continue
        new_vals = {r.new_value for r in rows}
        old_vals = {r.old_value for r in rows}
        types = {r.change_type for r in rows}
        # Numeric-aware: if all new_values are pairwise equal numerically, still flag
        # multi-CR overlap (process conflict) — keep existing behavior of flagging
        # any multi-CR touch.
        conflicts.append(
            {
                "conflict_type": "cell_overlap",
                "table_name": table,
                "key_tuple": key_tuple,
                "column_name": column,
                "change_request_ids": sorted(cr_ids),
                "change_types": sorted(types),
                "old_values": sorted(old_vals),
                "new_values": sorted(new_vals),
                "notes": "Multiple change requests modify the same cell",
                "resolved": "N",
            }
        )

    structural_types = {
        "key_count_change",
        "column_add",
        "column_delete",
        "column_rename",
        "table_add",
    }
    struct_index: dict[str, list[ChangeRow]] = {}
    for row in change_rows:
        if row.change_type in structural_types:
            struct_index.setdefault(row.table_name, []).append(row)

    for table, rows in struct_index.items():
        cr_ids = {r.change_request_id for r in rows}
        if len(cr_ids) < 2:
            continue
        by_cr: dict[str, list[ChangeRow]] = {}
        for r in rows:
            by_cr.setdefault(r.change_request_id, []).append(r)
        if len(by_cr) >= 2:
            conflicts.append(
                {
                    "conflict_type": "structural_collision",
                    "table_name": table,
                    "key_tuple": "",
                    "column_name": "",
                    "change_request_ids": sorted(cr_ids),
                    "change_types": sorted({r.change_type for r in rows}),
                    "old_values": [],
                    "new_values": [],
                    "notes": "Multiple change requests apply structural changes to the same table",
                    "resolved": "N",
                }
            )

    # row_add × column_add intersection with no covering value in the Change Log
    by_table: dict[str, list[ChangeRow]] = {}
    for row in change_rows:
        by_table.setdefault(row.table_name, []).append(row)

    for table, rows in by_table.items():
        added_cols = {
            r.column_name: r.change_request_id
            for r in rows
            if r.change_type == "column_add" and r.column_name
        }
        if not added_cols:
            continue

        added_key_crs: dict[str, set[str]] = {}
        for r in rows:
            if r.change_type == "row_add" and r.key_tuple:
                added_key_crs.setdefault(r.key_tuple, set()).add(r.change_request_id)
        if not added_key_crs:
            continue

        covered = {
            (r.key_tuple, r.column_name)
            for r in rows
            if r.change_type in {"row_add", "value_update"}
            and r.key_tuple
            and r.column_name
        }

        for key_tuple, row_crs in sorted(added_key_crs.items()):
            for col, col_cr in sorted(added_cols.items()):
                if (key_tuple, col) in covered:
                    continue
                cr_ids = sorted(row_crs | {col_cr})
                conflicts.append(
                    {
                        "conflict_type": "missing_row_column_fill",
                        "table_name": table,
                        "key_tuple": key_tuple,
                        "column_name": col,
                        "change_request_ids": cr_ids,
                        "change_types": ["column_add", "row_add"],
                        "old_values": [],
                        "new_values": [],
                        "notes": (
                            "Row add + column add leave this cell without a value; "
                            "add a Detail value_update (or row_add cell) or set "
                            "resolved=Y if blank is intentional"
                        ),
                        "resolved": "N",
                    }
                )

    return conflicts
