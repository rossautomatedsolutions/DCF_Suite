"""Deprecated entry point.

Use `python -m dcf_factory build --out <path>` to generate the workbook.
"""

from __future__ import annotations

import sys


def main() -> int:
    message = (
        "This entry point is deprecated. "
        "Use `python -m dcf_factory build --out <path>` instead."
    )
    print(message, file=sys.stderr)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
