"""CLI entry point for the Prophet Table Change Tool."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .changelog import generate_change_log
from .control import read_control
from .integrate import integrate_changes


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
        help="Path to ChangeLog_*.xlsx (required for validate_only / apply)",
    )
    parser.add_argument(
        "--mode",
        choices=["generate_changelog", "validate_only", "apply"],
        default=None,
        help="Override mode from Control.xlsx Config sheet",
    )
    args = parser.parse_args(argv)

    control = read_control(args.control)
    mode = args.mode or control.mode

    if mode == "generate_changelog":
        out = generate_change_log(args.control)
        print(f"Change Log written to: {out}")
        return 0

    if mode in {"validate_only", "apply"}:
        if args.change_log is None:
            parser.error("--change-log is required for validate_only / apply modes")
        out = integrate_changes(args.control, args.change_log, mode)  # type: ignore[arg-type]
        print(f"Integration Report written to: {out}")
        return 0

    print(f"Unknown mode: {mode!r}", file=sys.stderr)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
