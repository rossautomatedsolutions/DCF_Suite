"""Command-line interface for the DCF factory."""

from __future__ import annotations

import argparse
from pathlib import Path

from dcf_factory.build_dcf import build_dcf


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="dcf_factory")
    subparsers = parser.add_subparsers(dest="command", required=True)

    build_parser = subparsers.add_parser("build", help="Generate the DCF workbook")
    build_parser.add_argument(
        "--out",
        required=True,
        type=Path,
        help="Output path for the generated workbook",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    if args.command == "build":
        build_dcf(args.out)
        return 0

    parser.error("Unknown command")
    return 1
