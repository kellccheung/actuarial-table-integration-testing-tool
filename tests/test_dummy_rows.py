"""Tests for leading/trailing dummy lines around the !N / * content block."""

from __future__ import annotations

from pathlib import Path

from prophet_table_tool.diff import diff_table
from prophet_table_tool.control import ControlConfig
from prophet_table_tool.prophet_csv import (
    discover_csv_tables,
    read_prophet_csv,
    write_prophet_csv,
)


SAMPLE_FAC = """\
Prophet table,,,,,
!4,Age,Duration,Product,Rate,Loading
*,20,1,PROD_A,0.0013,1.05
*,25,1,PROD_A,0.0015,1.05
*,30,1,PROD_B,0.002,1.1
Edited on 2026,,,,,
"""


def _minimal_control() -> ControlConfig:
    return ControlConfig(
        control_path=Path("Control.xlsx"),
        working_root=Path("."),
        mode="generate_changelog",
        production_tables_path=Path("."),
        output_path=Path("."),
        run_id="",
    )


def test_read_fac_strips_dummy_lines(tmp_path: Path):
    path = tmp_path / "MORT_TABLE.FAC"
    path.write_text(SAMPLE_FAC, encoding="utf-8")

    table = read_prophet_csv(path)

    assert table.n_keys == 4
    assert table.columns == ["Age", "Duration", "Product", "Rate", "Loading"]
    assert table.data.height == 3
    assert table.leading_dummy_lines == ["Prophet table,,,,,"]
    assert table.trailing_dummy_lines == ["Edited on 2026,,,,,"]
    assert table.data["Age"].to_list() == ["20", "25", "30"]


def test_write_preserves_production_dummies(tmp_path: Path):
    path = tmp_path / "MORT_TABLE.FAC"
    path.write_text(SAMPLE_FAC, encoding="utf-8")
    table = read_prophet_csv(path)

    # Mutate content only
    table.data = table.data.with_columns(
        table.data["Rate"].str.replace("0.0013", "0.0099")
    )

    out = tmp_path / "out" / "MORT_TABLE.FAC"
    write_prophet_csv(table, out)
    text = out.read_text(encoding="utf-8")

    assert text.startswith("Prophet table,,,,,\n!4,")
    assert "*,20,1,PROD_A,0.0099,1.05\n" in text
    assert text.rstrip("\n").endswith("Edited on 2026,,,,,")


def test_diff_ignores_differing_dummy_lines(tmp_path: Path):
    before_text = (
        "BEFORE DUMMY,,,,,\n"
        "!4,Age,Duration,Product,Rate,Loading\n"
        "*,20,1,PROD_A,0.0012,1.05\n"
        "BEFORE END,,,,,\n"
    )
    after_text = (
        "AFTER DUMMY,,,,,\n"
        "!4,Age,Duration,Product,Rate,Loading\n"
        "*,20,1,PROD_A,0.0012,1.05\n"
        "*,25,1,PROD_A,0.0015,1.05\n"
        "AFTER END,,,,,\n"
    )
    before_path = tmp_path / "before.FAC"
    after_path = tmp_path / "after.FAC"
    before_path.write_text(before_text, encoding="utf-8")
    after_path.write_text(after_text, encoding="utf-8")

    before = read_prophet_csv(before_path)
    after = read_prophet_csv(after_path)
    result = diff_table("CR1", "MORT_TABLE", before, after, _minimal_control())

    types = {r.change_type for r in result.change_rows}
    assert types == {"row_add"}
    assert all(r.key_tuple == "25|1|PROD_A" for r in result.change_rows)
    # Dummy-only differences must not create deletes/updates
    assert not any(r.change_type in {"row_delete", "value_update"} for r in result.change_rows)


def test_discover_prefers_fac_over_csv(tmp_path: Path):
    (tmp_path / "MORT_TABLE.csv").write_text(
        "!4,Age,Duration,Product,Rate,Loading\n*,20,1,PROD_A,0.0012,1.05\n",
        encoding="utf-8",
    )
    (tmp_path / "MORT_TABLE.FAC").write_text(SAMPLE_FAC, encoding="utf-8")
    (tmp_path / "EXPENSE_TABLE.csv").write_text(
        "!2,Product,Expense\n*,PROD_A,10\n",
        encoding="utf-8",
    )

    found = discover_csv_tables(tmp_path)
    assert set(found) == {"MORT_TABLE", "EXPENSE_TABLE"}
    assert found["MORT_TABLE"].suffix.lower() == ".fac"
    assert found["EXPENSE_TABLE"].suffix.lower() == ".csv"


def test_discover_recursive_prefers_fac(tmp_path: Path):
    sub = tmp_path / "Nested"
    sub.mkdir()
    (sub / "MORT_TABLE.csv").write_text(
        "!4,Age,Duration,Product,Rate,Loading\n*,20,1,PROD_A,0.0012,1.05\n",
        encoding="utf-8",
    )
    (sub / "MORT_TABLE.FAC").write_text(SAMPLE_FAC, encoding="utf-8")
    found = discover_csv_tables(tmp_path)
    assert set(found) == {"Nested/MORT_TABLE"}
    assert found["Nested/MORT_TABLE"].suffix.lower() == ".fac"


def test_read_plain_csv_without_dummies_still_works(tmp_path: Path):
    path = tmp_path / "MORT_TABLE.csv"
    path.write_text(
        "!4,Age,Duration,Product,Rate,Loading\n"
        "*,20,1,PROD_A,0.0012,1.05\n",
        encoding="utf-8",
    )
    table = read_prophet_csv(path)
    assert table.leading_dummy_lines == []
    assert table.trailing_dummy_lines == []
    assert table.data.height == 1
