"""Writes WEHAGO upload workbooks with shared OOXML strings.

WEHAGO's upload preview does not reliably decode the ``inlineStr`` cells
written by openpyxl. This module intentionally uses XlsxWriter only for the
final workbook: it emits ``xl/sharedStrings.xml`` and string cells reference
that table with ``t=\"s\"``. Do not reopen and save these files with openpyxl.
"""

from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass
from numbers import Real
from pathlib import Path

import xlsxwriter
from xlsxwriter.format import Format
from xlsxwriter.workbook import Workbook
from xlsxwriter.worksheet import Worksheet

from core.profile import ClientProfile
from core.target_schema import build_target_column_order, compute_grand_totals, drop_all_zero_columns
from core.wehago_schema import WehagoRecord

DATA_START_ROW = 2
INVALID_XML_CHARS = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")


def clean_excel_text(value: object) -> str:
    """Removes only XML-forbidden control characters from a cell value."""
    if value is None:
        return ""
    return INVALID_XML_CHARS.sub("", str(value))


@dataclass(frozen=True)
class _Formats:
    text: Format
    header: Format
    money: Format
    total_text: Format
    total_money: Format


def _build_formats(workbook: Workbook) -> _Formats:
    text = workbook.add_format(
        {
            "font_name": "Malgun Gothic",
            "font_size": 11,
        }
    )
    header = workbook.add_format(
        {
            "font_name": "Malgun Gothic",
            "font_size": 11,
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#E0E0E0",
            "border": 1,
            "border_color": "#999999",
        }
    )
    money = workbook.add_format(
        {
            "font_name": "Malgun Gothic",
            "font_size": 11,
            "num_format": "#,##0",
        }
    )
    total_text = workbook.add_format(
        {
            "font_name": "Malgun Gothic",
            "font_size": 11,
            "bold": True,
            "bg_color": "#F3F3F3",
            "top": 6,
            "top_color": "#666666",
        }
    )
    total_money = workbook.add_format(
        {
            "font_name": "Malgun Gothic",
            "font_size": 11,
            "bold": True,
            "bg_color": "#F3F3F3",
            "top": 6,
            "top_color": "#666666",
            "num_format": "#,##0",
        }
    )
    return _Formats(text, header, money, total_text, total_money)


def _write_value(
    worksheet: Worksheet,
    row: int,
    col: int,
    value: object,
    *,
    text_format: Format,
    money_format: Format,
) -> None:
    if isinstance(value, Real) and not isinstance(value, bool):
        worksheet.write_number(row, col, value, money_format)
        return
    worksheet.write_string(row, col, clean_excel_text(value), text_format)


def _write_sheet(
    worksheet: Worksheet,
    profile: ClientProfile,
    records: list[WehagoRecord],
    formats: _Formats,
) -> None:
    allowance_columns = drop_all_zero_columns(profile.allowance_columns, records, "allowance")
    deduction_columns = drop_all_zero_columns(profile.deduction_columns, records, "deduction")
    columns = build_target_column_order(allowance_columns, deduction_columns)

    for col_idx, col in enumerate(columns):
        worksheet.write_string(0, col_idx, clean_excel_text(col.label), formats.header)
        worksheet.set_column(col_idx, col_idx, 14 if col.group in ("allowance", "deduction", "total") else 10)
    worksheet.set_row(0, 20)

    for record_idx, record in enumerate(records, start=1):
        for col_idx, col in enumerate(columns):
            _write_value(
                worksheet,
                record_idx,
                col_idx,
                record.column_value(col.key),
                text_format=formats.text,
                money_format=formats.money,
            )

    totals_row_idx = len(records) + 1
    grand_totals = compute_grand_totals(columns, records)
    for col_idx, col in enumerate(columns):
        _write_value(
            worksheet,
            totals_row_idx,
            col_idx,
            grand_totals[col.key],
            text_format=formats.total_text,
            money_format=formats.total_money,
        )

    worksheet.freeze_panes(1, 0)


def validate_wehago_xlsx(path: Path) -> None:
    """Fails fast if a saved upload file is not in WEHAGO's required form."""
    with zipfile.ZipFile(path, "r") as archive:
        names = archive.namelist()
        if "xl/sharedStrings.xml" not in names:
            raise ValueError("WEHAGO compatibility error: sharedStrings.xml is missing.")

        sheet_files = [
            name
            for name in names
            if name.startswith("xl/worksheets/sheet") and name.endswith(".xml")
        ]
        for sheet_file in sheet_files:
            xml = archive.read(sheet_file).decode("utf-8")
            if 't="inlineStr"' in xml:
                raise ValueError(f"WEHAGO compatibility error: inlineStr exists in {sheet_file}.")
            if 't="s"' not in xml:
                raise ValueError(f"WEHAGO compatibility error: shared-string cells are missing from {sheet_file}.")

        shared_strings = archive.read("xl/sharedStrings.xml").decode("utf-8")
        if "사원명" not in shared_strings:
            raise ValueError("WEHAGO compatibility error: Korean headers are missing from shared strings.")


def _save_workbook(output_path: Path, sheets: list[tuple[str, ClientProfile, list[WehagoRecord]]]) -> None:
    workbook = xlsxwriter.Workbook(
        str(output_path),
        {
            "constant_memory": False,
            "strings_to_urls": False,
        },
    )
    formats = _build_formats(workbook)
    try:
        for sheet_name, profile, records in sheets:
            worksheet = workbook.add_worksheet(clean_excel_text(sheet_name))
            _write_sheet(worksheet, profile, records, formats)
    finally:
        workbook.close()

    validate_wehago_xlsx(output_path)


def export_wehago(profile: ClientProfile, records: list[WehagoRecord], output_path: Path) -> None:
    _save_workbook(output_path, [("Sheet1", profile, records)])


def export_wehago_multi_sheet(
    profile: ClientProfile, sheets: dict[str, list[WehagoRecord]], output_path: Path
) -> None:
    """Writes one workbook with one sheet per payroll service-line."""
    _save_workbook(output_path, [(sheet_name, profile, records) for sheet_name, records in sheets.items()])
