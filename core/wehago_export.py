"""Read WEHAGO's own payroll ledger export (급여자료입력 → Ctrl+G 내보내기).

This is the persisted per-employee evidence used before closing. The layout
was observed in 74 stored reference files and matches the export parsed by the
급여자동화 1.52 product (see REFERENCE_152_WEHAGO_UI.md):

  row 1: 사원코드 | 사원명 | 부서 | 직급 | 직종 | 수당 (merged) | 공제 (merged) | 차인지급액
  row 2: item labels under 수당 ... 지급액계, under 공제 ... 공제액계
  row 3+: one employee per row, code kept as the exact text ("1" != "01")
  last:   "합계" row with comma-formatted text totals

Every rule failure raises ``ExportFormatError``; nothing is guessed, rounded or
defaulted to zero. Names are returned for matching but never logged here.
"""
from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
import re
import unicodedata

from core.payroll_workflow import Amounts, PersonAmount


class ExportFormatError(ValueError):
    pass


_FILENAME = re.compile(
    r"(?P<company>.+?)\s*-\s*(?P<period>20\d{2}(?:0[1-9]|1[0-2]))(?:\s*\(\d+\))?\.xlsx$")


@dataclass(frozen=True)
class ExportIdentity:
    company: str
    year: int
    month: int


@dataclass(frozen=True)
class WehagoExport:
    people: tuple[PersonAmount, ...]
    totals: Amounts


def parse_export_filename(name: str) -> ExportIdentity:
    match = _FILENAME.fullmatch(unicodedata.normalize("NFC", Path(name).name))
    if not match:
        raise ExportFormatError("위하고 급여대장 내보내기 파일명(업체명 - YYYYMM.xlsx)이 아닙니다.")
    period = match["period"]
    return ExportIdentity(match["company"].strip(), int(period[:4]), int(period[4:]))


def _text(value) -> str:
    return "" if value is None else "".join(str(value).split())


def _won(value, where: str) -> Decimal:
    if value is None or isinstance(value, bool):
        raise ExportFormatError(f"{where}: 금액이 비어 있습니다.")
    if isinstance(value, float):
        if not value.is_integer():
            raise ExportFormatError(f"{where}: 원 단위 정수가 아닙니다.")
        value = int(value)
    text = str(value).strip().replace(",", "")
    try:
        amount = Decimal(text)
    except InvalidOperation:
        raise ExportFormatError(f"{where}: 금액을 읽지 못했습니다.") from None
    if not amount.is_finite() or amount != amount.to_integral_value():
        raise ExportFormatError(f"{where}: 원 단위 정수가 아닙니다.")
    return amount


def _layout(header1: list, header2: list) -> tuple[int, int, int]:
    h1 = [_text(v) for v in header1]
    h2 = [_text(v) for v in header2]
    if h1[:2] != ["사원코드", "사원명"]:
        raise ExportFormatError("위하고 급여자료입력의 엑셀 내려받기 파일이 아닙니다.")
    try:
        earn_start, ded_start = h1.index("수당"), h1.index("공제")
        net_col = h1.index("차인지급액")
        gross_col, ded_col = h2.index("지급액계"), h2.index("공제액계")
    except ValueError:
        raise ExportFormatError("위하고 내려받기 제목 구조가 변경되었습니다.") from None
    if not earn_start < gross_col < ded_start < ded_col < net_col:
        raise ExportFormatError("위하고 내려받기 제목 순서가 예상과 다릅니다.")
    if h2.count("지급액계") != 1 or h2.count("공제액계") != 1 or h1.count("차인지급액") != 1:
        raise ExportFormatError("위하고 내려받기 합계 열이 하나로 확인되지 않습니다.")
    # Item columns are deliberately not summed: real exports (company_68/98)
    # show WEHAGO's 지급액계 excluding some item columns, and duplicate item
    # titles (식대 in 수당 and 공제) are common. Closing compares WEHAGO's own
    # per-person totals, so only those three columns are authoritative.
    return gross_col, ded_col, net_col


def read_wehago_export(path: Path) -> WehagoExport:
    import openpyxl

    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        rows = [list(r) for r in workbook.worksheets[0].iter_rows(values_only=True)]
    finally:
        workbook.close()
    if len(rows) < 3:
        raise ExportFormatError("위하고 내보내기에 사원 행이 없습니다.")
    width = max(len(r) for r in rows)
    rows = [r + [None] * (width - len(r)) for r in rows]
    gross_col, ded_col, net_col = _layout(rows[0], rows[1])

    people: list[PersonAmount] = []
    seen: set[str] = set()
    total_row = None
    for index, row in enumerate(rows[2:], start=3):
        first = _text(row[0])
        if first == "합계":
            total_row = row
            if any(any(v not in (None, "") for v in r) for r in rows[index:]):
                raise ExportFormatError("합계 행 아래에 다른 자료가 있습니다.")
            break
        if not first and all(v in (None, "") for v in row):
            continue
        if not isinstance(row[0], str) or not re.fullmatch(r"\d+", first):
            raise ExportFormatError(f"{index}행: 사원코드 형식을 확인하세요(문자 그대로의 숫자여야 함).")
        if first in seen:
            raise ExportFormatError(f"위하고 내보내기에 동일한 사원코드가 반복됩니다: {first}")
        seen.add(first)
        name = str(row[1] or "").strip()
        if not name:
            raise ExportFormatError(f"{index}행: 이름 없는 사원이 있습니다.")
        where = f"{index}행 코드 {first}"
        gross = _won(row[gross_col], where)
        ded = _won(row[ded_col], where)
        net = _won(row[net_col], where)
        if gross - ded != net:
            raise ExportFormatError(f"{where}: 위하고 차인지급액이 맞지 않습니다.")
        people.append(PersonAmount(name, Amounts(gross, ded, net), first))
    if not people:
        raise ExportFormatError("선택한 지급일에 급여 사원이 없습니다.")
    sums = Amounts(*(sum((getattr(p.amounts, k) for p in people), Decimal(0))
                     for k in ("gross", "deductions", "net")))
    if total_row is None:
        raise ExportFormatError("위하고 내보내기의 합계 행을 찾지 못했습니다.")
    stated = Amounts(_won(total_row[gross_col], "합계행"), _won(total_row[ded_col], "합계행"),
                     _won(total_row[net_col], "합계행"))
    if stated != sums:
        raise ExportFormatError("위하고 내보내기 합계 행과 사원별 합계가 다릅니다.")
    return WehagoExport(tuple(people), stated)
