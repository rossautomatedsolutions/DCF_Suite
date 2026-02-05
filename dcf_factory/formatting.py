"""Formatting helpers for the DCF workbook."""

from __future__ import annotations

from openpyxl.styles import Alignment, Font, NamedStyle, PatternFill


def build_styles() -> dict[str, NamedStyle]:
    header = NamedStyle(name="header")
    header.font = Font(bold=True, color="FFFFFF")
    header.fill = PatternFill("solid", fgColor="1F4E78")
    header.alignment = Alignment(horizontal="center", vertical="center")

    label = NamedStyle(name="label")
    label.font = Font(bold=True)
    label.alignment = Alignment(horizontal="left", vertical="center")

    number = NamedStyle(name="number")
    number.number_format = "#,##0.00"

    percent = NamedStyle(name="percent")
    percent.number_format = "0.00%"

    currency = NamedStyle(name="currency")
    currency.number_format = "$#,##0.00"

    return {
        "header": header,
        "label": label,
        "number": number,
        "percent": percent,
        "currency": currency,
    }


def register_styles(workbook, styles: dict[str, NamedStyle]) -> None:
    for style in styles.values():
        if style.name not in workbook.named_styles:
            workbook.add_named_style(style)
