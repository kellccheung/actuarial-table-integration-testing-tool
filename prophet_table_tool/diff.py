"""Table-level diffing between before/ and after/ Prophet CSVs."""

from __future__ import annotations

from dataclasses import dataclass, field

import polars as pl

from .control import ControlConfig
from .prophet_csv import ProphetTable, with_normalized_keys
from .utils import values_equal


CELL_TYPES = frozenset({"value_update", "row_add", "row_delete"})
STRUCTURAL_TYPES = frozenset(
    {
        "key_count_change",
        "column_add",
        "column_delete",
        "column_rename",
        "table_add",
    }
)

CELL_SCHEMA: dict[str, pl.DataType] = {
    "change_request_id": pl.Utf8,
    "table_name": pl.Utf8,
    "change_type": pl.Utf8,
    "n_keys_before": pl.Int64,
    "n_keys_after": pl.Int64,
    "key_tuple": pl.Utf8,
    "column_name": pl.Utf8,
    "old_value": pl.Utf8,
    "new_value": pl.Utf8,
    "notes": pl.Utf8,
}


def empty_cell_changes() -> pl.DataFrame:
    return pl.DataFrame(schema=CELL_SCHEMA)


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
    structural_rows: list[ChangeRow] = field(default_factory=list)
    cell_changes: pl.DataFrame = field(default_factory=empty_cell_changes)
    warnings: list[str] = field(default_factory=list)
    tables_touched: set[str] = field(default_factory=set)

    @property
    def change_rows(self) -> list[ChangeRow]:
        """Structural rows plus materialized cell changes (tests / small tables)."""
        return self.structural_rows + cell_df_to_rows(self.cell_changes)


def _int_lit(value: int | None) -> pl.Expr:
    if value is None:
        return pl.lit(None, dtype=pl.Int64)
    return pl.lit(int(value), dtype=pl.Int64)


def _unpivot(
    df: pl.DataFrame,
    index: list[str],
    on: list[str],
    variable_name: str,
    value_name: str,
) -> pl.DataFrame:
    return df.unpivot(
        index=index, on=on, variable_name=variable_name, value_name=value_name
    )


def rows_to_cell_df(rows: list[ChangeRow]) -> pl.DataFrame:
    if not rows:
        return empty_cell_changes()
    return pl.DataFrame(
        {
            "change_request_id": [r.change_request_id for r in rows],
            "table_name": [r.table_name for r in rows],
            "change_type": [r.change_type for r in rows],
            "n_keys_before": [r.n_keys_before for r in rows],
            "n_keys_after": [r.n_keys_after for r in rows],
            "key_tuple": [r.key_tuple for r in rows],
            "column_name": [r.column_name for r in rows],
            "old_value": [r.old_value for r in rows],
            "new_value": [r.new_value for r in rows],
            "notes": [r.notes for r in rows],
        }
    ).with_columns(
        pl.col("n_keys_before").cast(pl.Int64),
        pl.col("n_keys_after").cast(pl.Int64),
        pl.col("old_value").cast(pl.Utf8).fill_null(""),
        pl.col("new_value").cast(pl.Utf8).fill_null(""),
        pl.col("notes").cast(pl.Utf8).fill_null(""),
    )


def cell_df_to_rows(df: pl.DataFrame) -> list[ChangeRow]:
    if df.is_empty():
        return []
    out: list[ChangeRow] = []
    for rec in df.iter_rows(named=True):
        n_before = rec["n_keys_before"]
        n_after = rec["n_keys_after"]
        out.append(
            ChangeRow(
                change_request_id=str(rec["change_request_id"] or ""),
                table_name=str(rec["table_name"] or ""),
                change_type=str(rec["change_type"] or ""),
                n_keys_before=None if n_before is None else int(n_before),
                n_keys_after=None if n_after is None else int(n_after),
                key_tuple=str(rec["key_tuple"] or ""),
                column_name=str(rec["column_name"] or ""),
                old_value="" if rec["old_value"] is None else str(rec["old_value"]),
                new_value="" if rec["new_value"] is None else str(rec["new_value"]),
                notes=str(rec["notes"] or ""),
            )
        )
    return out


def concat_cell_changes(frames: list[pl.DataFrame]) -> pl.DataFrame:
    nonempty = [f for f in frames if f.height]
    if not nonempty:
        return empty_cell_changes()
    if len(nonempty) == 1:
        return nonempty[0]
    return pl.concat(nonempty, how="vertical")


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
        result.structural_rows.append(
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
        result.cell_changes = _emit_row_cell_changes(
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

    result.tables_touched.add(table_name)

    if before.n_keys != after.n_keys:
        result.structural_rows.append(
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

    declared = control.renames_for(change_request_id, table_name)
    rename_old_to_new = {r.old_column_name: r.new_column_name for r in declared}
    rename_new_to_old = {v: k for k, v in rename_old_to_new.items()}

    for old_c, new_c in rename_old_to_new.items():
        result.structural_rows.append(
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
        result.structural_rows.append(
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
        result.structural_rows.append(
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

    b_keys = b.select("_key_str")
    a_keys = a.select("_key_str")

    cell_frames: list[pl.DataFrame] = []

    a_add = a.join(b_keys, on="_key_str", how="anti")
    if a_add.height:
        cell_frames.append(
            _emit_row_cell_changes(
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
        )

    b_del = b.join(a_keys, on="_key_str", how="anti")
    if b_del.height:
        cell_frames.append(
            _emit_row_cell_changes(
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
        )

    if common_value_cols:
        b_common = b.join(a_keys, on="_key_str", how="inner")
        a_common = a.join(b_keys, on="_key_str", how="inner")
        if b_common.height and a_common.height:
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
                cell_frames.append(
                    diffs.select(
                        pl.lit(change_request_id).alias("change_request_id"),
                        pl.lit(table_name).alias("table_name"),
                        pl.lit("value_update").alias("change_type"),
                        _int_lit(before.n_keys).alias("n_keys_before"),
                        _int_lit(after.n_keys).alias("n_keys_after"),
                        pl.col("_key_str").alias("key_tuple"),
                        pl.lit(col).alias("column_name"),
                        pl.col("old_value"),
                        pl.col("new_value"),
                        pl.lit("").alias("notes"),
                    )
                )

    genuinely_new_cols = [c for c in after.columns if c not in before_aligned.columns]
    if genuinely_new_cols:
        a_common_new = a.join(b_keys, on="_key_str", how="inner")
        if a_common_new.height:
            cell_frames.append(
                _emit_row_cell_changes(
                    change_request_id=change_request_id,
                    table_name=table_name,
                    change_type="value_update",
                    n_keys_before=before.n_keys,
                    n_keys_after=after.n_keys,
                    keyed=a_common_new,
                    columns=genuinely_new_cols,
                    value_from="new",
                    notes="Value on newly added column",
                )
            )

    result.cell_changes = concat_cell_changes(cell_frames)
    return result


def _emit_row_cell_changes(
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
) -> pl.DataFrame:
    """One cell-change row per present column for row_add / row_delete / new-col fills."""
    present_cols = [c for c in columns if c in keyed.columns]
    if not present_cols or keyed.is_empty():
        return empty_cell_changes()

    melted = _unpivot(
        keyed.select(["_key_str"] + present_cols),
        index=["_key_str"],
        on=present_cols,
        variable_name="column_name",
        value_name="cell",
    )
    val = pl.col("cell").cast(pl.Utf8).fill_null("")
    if value_from == "new":
        old_e, new_e = pl.lit(""), val
    else:
        old_e, new_e = val, pl.lit("")
    return melted.select(
        pl.lit(change_request_id).alias("change_request_id"),
        pl.lit(table_name).alias("table_name"),
        pl.lit(change_type).alias("change_type"),
        _int_lit(n_keys_before).alias("n_keys_before"),
        _int_lit(n_keys_after).alias("n_keys_after"),
        pl.col("_key_str").alias("key_tuple"),
        pl.col("column_name"),
        old_e.alias("old_value"),
        new_e.alias("new_value"),
        pl.lit(notes).alias("notes"),
    )


def detect_conflicts(
    change_rows: list[ChangeRow],
    cell_changes: pl.DataFrame | None = None,
) -> list[dict]:
    """
    Detect conflicts across change requests:
    same table + same key + same column with differing changes,
    structural collisions (key_count / column changes) on the same table
    from different CRs, or missing fills where a row_add and column_add
    leave their intersection cell without a Change Log value.
    """
    structural = [r for r in change_rows if r.change_type not in CELL_TYPES]
    cells = cell_changes if cell_changes is not None else empty_cell_changes()
    list_cells = [r for r in change_rows if r.change_type in CELL_TYPES]
    if list_cells:
        from_list = rows_to_cell_df(list_cells)
        cells = concat_cell_changes([cells, from_list])
    return _detect_conflicts(structural, cells)


def _detect_conflicts(
    structural: list[ChangeRow],
    cells: pl.DataFrame,
) -> list[dict]:
    conflicts: list[dict] = []

    column_add_index: dict[tuple[str, str], set[str]] = {}
    for row in structural:
        if row.change_type == "column_add" and row.column_name:
            column_add_index.setdefault(
                (row.table_name, row.column_name), set()
            ).add(row.change_request_id)

    if cells.height:
        cell_work = cells.filter(
            pl.col("change_type").is_in(["value_update", "row_add", "row_delete"])
            & (pl.col("column_name") != "")
        )
        if cell_work.height:
            hot = (
                cell_work.group_by(["table_name", "key_tuple", "column_name"])
                .agg(pl.col("change_request_id").n_unique().alias("n_cr"))
                .filter(pl.col("n_cr") >= 2)
                .drop("n_cr")
            )
            if hot.height:
                detail = cell_work.join(
                    hot, on=["table_name", "key_tuple", "column_name"], how="inner"
                )
                grouped = detail.group_by(["table_name", "key_tuple", "column_name"]).agg(
                    pl.col("change_request_id").alias("cr_ids"),
                    pl.col("new_value").alias("new_vals"),
                    pl.col("old_value").alias("old_vals"),
                    pl.col("change_type").alias("types"),
                )
                for rec in grouped.iter_rows(named=True):
                    cr_ids = {str(x) for x in rec["cr_ids"]}
                    if len(cr_ids) < 2:
                        continue
                    new_vals = ["" if v is None else str(v) for v in rec["new_vals"]]
                    old_vals = ["" if v is None else str(v) for v in rec["old_vals"]]
                    types = {str(t) for t in rec["types"]}
                    first_new = new_vals[0] if new_vals else ""
                    same_new_value = all(values_equal(v, first_new) for v in new_vals[1:])
                    table = rec["table_name"]
                    column = rec["column_name"]
                    col_add_crs = column_add_index.get((table, column), set())
                    sequenced_col = bool(col_add_crs & cr_ids) and "value_update" in types
                    sequenced_row = "row_add" in types and "value_update" in types
                    if same_new_value:
                        notes = (
                            "Multiple change requests modify the same cell with the same "
                            "new_value; set resolved=Y to accept both writes"
                        )
                    elif sequenced_row or sequenced_col:
                        kinds: list[str] = []
                        if sequenced_row:
                            kinds.append("row_add and value_update")
                        if sequenced_col:
                            kinds.append("column_add and value_update")
                        notes = (
                            "Multiple change requests modify the same cell ("
                            + " / ".join(kinds)
                            + "); set resolved=Y to apply both in Control order"
                        )
                    else:
                        notes = (
                            "Multiple change requests modify the same cell; set resolved=Y "
                            "to apply both in Control order (later CR wins)"
                        )
                    conflicts.append(
                        {
                            "conflict_type": "cell_overlap",
                            "table_name": table,
                            "key_tuple": rec["key_tuple"],
                            "column_name": column,
                            "change_request_ids": sorted(cr_ids),
                            "change_types": sorted(types),
                            "old_values": sorted(set(old_vals)),
                            "new_values": sorted(set(new_vals)),
                            "notes": notes,
                            "resolved": "N",
                        }
                    )

    struct_index: dict[str, list[ChangeRow]] = {}
    for row in structural:
        if row.change_type in STRUCTURAL_TYPES:
            struct_index.setdefault(row.table_name, []).append(row)

    for table, rows in struct_index.items():
        cr_ids = {r.change_request_id for r in rows}
        if len(cr_ids) < 2:
            continue
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

    added_col_rows = [
        (r.table_name, r.column_name, r.change_request_id)
        for r in structural
        if r.change_type == "column_add" and r.column_name
    ]
    if added_col_rows and cells.height:
        added_cols = pl.DataFrame(
            added_col_rows,
            schema=["table_name", "column_name", "col_cr"],
            orient="row",
        )
        added_keys = (
            cells.filter(
                (pl.col("change_type") == "row_add") & (pl.col("key_tuple") != "")
            )
            .select(["table_name", "key_tuple", "change_request_id"])
            .unique()
        )
        covered = (
            cells.filter(
                pl.col("change_type").is_in(["row_add", "value_update"])
                & (pl.col("key_tuple") != "")
                & (pl.col("column_name") != "")
            )
            .select(["table_name", "key_tuple", "column_name"])
            .unique()
        )
        if added_keys.height:
            pairs = added_keys.join(added_cols, on="table_name", how="inner")
            gaps = pairs.join(
                covered, on=["table_name", "key_tuple", "column_name"], how="anti"
            )
            if gaps.height:
                grouped_gaps = gaps.group_by(
                    ["table_name", "key_tuple", "column_name"]
                ).agg(
                    pl.col("change_request_id").unique().alias("row_crs"),
                    pl.col("col_cr").unique().alias("col_crs"),
                )
                for rec in sorted(
                    grouped_gaps.iter_rows(named=True),
                    key=lambda r: (r["table_name"], r["key_tuple"], r["column_name"]),
                ):
                    cr_ids = sorted(
                        {str(x) for x in rec["row_crs"]}
                        | {str(x) for x in rec["col_crs"]}
                    )
                    conflicts.append(
                        {
                            "conflict_type": "missing_row_column_fill",
                            "table_name": rec["table_name"],
                            "key_tuple": rec["key_tuple"],
                            "column_name": rec["column_name"],
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
