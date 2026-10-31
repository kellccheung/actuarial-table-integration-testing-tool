"""CLI entry point for the Prophet Table Change Tool."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from openpyxl import load_workbook

from .changelog import StructureScanError, generate_change_log, run_structure_scan
from .control import ControlConfig, read_control
from .integrate import integrate_changes


class ChangeLogResolveError(Exception):
    """Raised when Stage 2 cannot resolve a Change Log workbook."""


def _newest_change_log(control: ControlConfig) -> Path:
    if not control.output_path.is_dir():
        raise ChangeLogResolveError(
            f"Output folder not found: {control.output_path}\n"
            "Run Stage 1 (generate_changelog) first, or check output_path in Control.xlsx."
        )
    logs = sorted(
        control.output_path.glob("ChangeLog_*.xlsx"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    if not logs:
        raise ChangeLogResolveError(
            f"No ChangeLog_*.xlsx found in {control.output_path}\n"
            "Run Stage 1 (generate_changelog) first, or place a Change Log in Output."
        )
    return logs[0]


def _resolve_change_log(
    control: ControlConfig, explicit: Path | None
) -> Path:
    """Resolve Stage 2 Change Log: explicit path, bare Output filename, or newest."""
    if explicit is None:
        return _newest_change_log(control)

    if explicit.is_file():
        return explicit

    if len(explicit.parts) == 1:
        candidate = control.output_path / explicit.name
        if candidate.is_file():
            return candidate
        raise ChangeLogResolveError(
            f"Change Log not found: {explicit.name}\n"
            f"Looked in {control.output_path}"
        )

    raise ChangeLogResolveError(f"Change Log not found: {explicit}")


def _integration_failed(path: Path) -> bool:
    """True when the Integration Report Summary status is FAILED."""
    wb = load_workbook(path, data_only=True, read_only=True)
    try:
        return str(wb["Summary"]["B1"].value or "") == "FAILED"
    finally:
        wb.close()


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Prophet Table Change Consolidation & Integration Tool",
    )
    parser.add_argument(
        "control",
        type=Path,
        help="Path to Control.xlsx",
    )
    parser.add_argument(
        "--change-log",
        type=Path,
        default=None,
        help=(
            "Path or filename of ChangeLog_*.xlsx (optional for validate_only / "
            "apply; default is newest in Output)"
        ),
    )
    parser.add_argument(
        "--mode",
        choices=["generate_changelog", "validate_only", "apply", "scan"],
        default=None,
        help="Override mode from Control.xlsx Config sheet",
    )
    args = parser.parse_args(argv)

    control = read_control(args.control)
    mode = args.mode or control.mode
    print(f"Mode: {mode}")

    if mode == "scan":
        scan_path, findings = run_structure_scan(control)
        n_fatal = sum(1 for row in findings if row["severity"] == "FATAL")
        print(f"Structure scan written to: {scan_path}")
        print(f"Findings: {len(findings)} ({n_fatal} fatal)")
        return 1 if n_fatal else 0

    if mode == "generate_changelog":
        try:
            out = generate_change_log(args.control, control=control)
        except StructureScanError as exc:
            print(str(exc), file=sys.stderr)
            return 1
        print(f"Change Log written to: {out}")
        return 0

    if mode == "validate_only":
        try:
            change_log = _resolve_change_log(control, args.change_log)
        except ChangeLogResolveError as exc:
            print(str(exc), file=sys.stderr)
            return 1
        print(f"Using Change Log: {change_log}")
        out = integrate_changes(args.control, change_log, "validate_only", control=control)
        print(f"Integration Report written to: {out}")
        return 1 if _integration_failed(out) else 0

    if mode == "apply":
        try:
            change_log = _resolve_change_log(control, args.change_log)
        except ChangeLogResolveError as exc:
            print(str(exc), file=sys.stderr)
            return 1
        print(f"Using Change Log: {change_log}")
        out = integrate_changes(args.control, change_log, "apply", control=control)
        print(f"Integration Report written to: {out}")
        return 1 if _integration_failed(out) else 0

    print(f"Unknown mode: {mode!r}", file=sys.stderr)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
