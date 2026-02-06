"""Deprecated legacy generator.

The authoritative generator lives in `dcf_factory/`.
"""

from __future__ import annotations


def build_workbook():
    raise RuntimeError(
        "dcf_workbook.build_workbook is deprecated. "
        "Use `python -m dcf_factory build --out <path>` instead."
    )
