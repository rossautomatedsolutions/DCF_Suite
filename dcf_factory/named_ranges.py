"""Named range helpers."""

from __future__ import annotations

from openpyxl.workbook.defined_name import DefinedName


def add_named_range(workbook, name: str, sheet_title: str, cell: str) -> None:
    destination = f"'{sheet_title}'!${cell}"
    defined_name = DefinedName(name=name, attr_text=destination)
    workbook.defined_names.add(defined_name)
