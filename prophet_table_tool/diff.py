"""Table-level diffing between before/ and after/ Prophet CSVs."""

from __future__ import annotations

from dataclasses import dataclass, field

import polars as pl

from .control import ControlConfig
from .prophet_csv import ProphetTable


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
    if values is None:
        return ""
    return "|".join("" if v is None else str(v) for v in values)


def diff_table(
    change_request_id: str,
    table_name: str,
    before: ProphetTable | None,
    after: ProphetTable | None,
    control: ControlConfig,
) -> DiffResult:
    """
    Compute detailed changes between before and after for one table.

    - after only  → table_add
    - before only → warning + skip (caller should not call with after=None only;
      but we still handle it defensively)
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
        # Also emit row_add for each row so Stage 2 can rebuild the table
        after_keyed = after.with_key_tuple()
        for row in after_keyed.iter_rows(named=True):
            key_str = format_key_tuple(row["_key_tuple"])
            for col in after.columns:
                result.change_rows.append(
                    ChangeRow(
                        change_request_id=change_request_id,
                        table_name=table_name,
                        change_type="row_add",
                        n_keys_before=None,
                        n_keys_after=after.n_keys,
                        key_tuple=key_str,
                        column_name=col,
                        old_value="",
                        new_value="" if row[col] is None else str(row[col]),
                        notes="Part of table_add",
                    )
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

    # Normalize before columns through renames for comparison
    before_cols_mapped = [
        rename_old_to_new.get(c, c) for c in before.columns
    ]
    after_cols = list(after.columns)

    before_set = set(before_cols_mapped)
    after_set = set(after_cols)

    for col in sorted(after_set - before_set):
        # Skip if this is the new name of a declared rename
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

    # --- Row / value diffs (exact string match on key tuples) ---
    # Align before schema to after names via renames for value comparison
    before_aligned = before.data
    rename_exprs = [
        pl.col(old).alias(new)
        for old, new in rename_old_to_new.items()
        if old in before_aligned.columns
    ]
    if rename_exprs:
        before_aligned = before_aligned.rename(
            {old: new for old, new in rename_old_to_new.items() if old in before_aligned.columns}
        )

    # Use after key columns for matching when key count changed;
    # otherwise use the (possibly renamed) before key columns.
    if before.n_keys == after.n_keys:
        key_cols = after.key_columns
    else:
        # Structural change: still attempt match on overlapping key columns
        key_cols = [c for c in after.key_columns if c in before_aligned.columns]
        if not key_cols:
            # Cannot match rows — structural only; skip row-level diff
            return result

    common_value_cols = [
        c
        for c in after.columns
        if c in before_aligned.columns and c not in key_cols
    ]

    before_df = before_aligned
    after_df = after.data

    def with_keys(df: pl.DataFrame, cols: list[str]) -> pl.DataFrame:
        if not cols:
            return df.with_columns(
                pl.lit("").alias("_key_str"),
                pl.lit([]).cast(pl.List(pl.Utf8)).alias("_key_tuple"),
            )
        return df.with_columns(
            pl.concat_list([pl.col(c).cast(pl.Utf8) for c in cols]).alias("_key_tuple"),
            pl.concat_str([pl.col(c).cast(pl.Utf8) for c in cols], separator="|").alias(
                "_key_str"
            ),
        )

    # Only keep columns we need
    before_keep = [c for c in key_cols + common_value_cols if c in before_df.columns]
    after_keep = [c for c in key_cols + common_value_cols + [
        c for c in after.columns if c not in before_aligned.columns and c not in key_cols
    ] if c in after_df.columns]

    b = with_keys(before_df.select(before_keep), key_cols)
    a = with_keys(after_df.select(list(dict.fromkeys(after_keep))), key_cols)

    b_keys = set(b["_key_str"].to_list())
    a_keys = set(a["_key_str"].to_list())

    added_keys = a_keys - b_keys
    deleted_keys = b_keys - a_keys
    common_keys = b_keys & a_keys

    # Row adds
    if added_keys:
        a_add = a.filter(pl.col("_key_str").is_in(list(added_keys)))
        for row in a_add.iter_rows(named=True):
            key_str = row["_key_str"]
            for col in after.columns:
                if col not in row:
                    continue
                result.change_rows.append(
                    ChangeRow(
                        change_request_id=change_request_id,
                        table_name=table_name,
                        change_type="row_add",
                        n_keys_before=before.n_keys,
                        n_keys_after=after.n_keys,
                        key_tuple=key_str,
                        column_name=col,
                        old_value="",
                        new_value="" if row[col] is None else str(row[col]),
                        notes="",
                    )
                )

    # Row deletes
    if deleted_keys:
        b_del = b.filter(pl.col("_key_str").is_in(list(deleted_keys)))
        for row in b_del.iter_rows(named=True):
            key_str = row["_key_str"]
            for col in before_aligned.columns:
                if col not in row:
                    continue
                result.change_rows.append(
                    ChangeRow(
                        change_request_id=change_request_id,
                        table_name=table_name,
                        change_type="row_delete",
                        n_keys_before=before.n_keys,
                        n_keys_after=after.n_keys,
                        key_tuple=key_str,
                        column_name=col,
                        old_value="" if row[col] is None else str(row[col]),
                        new_value="",
                        notes="",
                    )
                )

    # Value updates on common keys
    if common_keys and common_value_cols:
        b_common = b.filter(pl.col("_key_str").is_in(list(common_keys))).sort("_key_str")
        a_common = a.filter(pl.col("_key_str").is_in(list(common_keys))).sort("_key_str")

        # Join on key for comparison
        joined = b_common.select(
            ["_key_str"] + [pl.col(c).alias(f"old_{c}") for c in common_value_cols]
        ).join(
            a_common.select(
                ["_key_str"] + [pl.col(c).alias(f"new_{c}") for c in common_value_cols]
            ),
            on="_key_str",
            how="inner",
        )

        for row in joined.iter_rows(named=True):
            key_str = row["_key_str"]
            for col in common_value_cols:
                old_v = row[f"old_{col}"]
                new_v = row[f"new_{col}"]
                old_s = "" if old_v is None else str(old_v)
                new_s = "" if new_v is None else str(new_v)
                if old_s != new_s:
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

    # Newly added columns: emit value fills for existing keys as value_update? 
    # Spec says column_add is enough structurally; values on new cols for existing
    # rows are captured if we treat them as value_update from empty.
    new_cols = [
        c for c in after.columns
        if c not in before_aligned.columns and c not in rename_new_to_old.values()
        or (c not in before_aligned.columns and c not in key_cols)
    ]
    # Cleaner:
    genuinely_new_cols = [c for c in after.columns if c not in before_aligned.columns]
    if genuinely_new_cols and common_keys:
        a_common = a.filter(pl.col("_key_str").is_in(list(common_keys)))
        for row in a_common.iter_rows(named=True):
            key_str = row["_key_str"]
            for col in genuinely_new_cols:
                if col not in row:
                    continue
                result.change_rows.append(
                    ChangeRow(
                        change_request_id=change_request_id,
                        table_name=table_name,
                        change_type="value_update",
                        n_keys_before=before.n_keys,
                        n_keys_after=after.n_keys,
                        key_tuple=key_str,
                        column_name=col,
                        old_value="",
                        new_value="" if row[col] is None else str(row[col]),
                        notes="Value on newly added column",
                    )
                )

    return result


def detect_conflicts(change_rows: list[ChangeRow]) -> list[dict]:
    """
    Detect conflicts across change requests:
    same table + same key + same column with differing changes,
    or structural collisions (key_count / column changes) on the same table
    from different CRs.
    """
    conflicts: list[dict] = []

    # Cell-level overlaps: (table, key_tuple, column) touched by >1 CR
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
        # Conflict if values disagree or multiple CRs touch same cell
        new_vals = {r.new_value for r in rows}
        old_vals = {r.old_value for r in rows}
        types = {r.change_type for r in rows}
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

    # Structural collisions on same table from different CRs
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
        # Flag if two CRs both change structure of same table
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

    return conflicts
