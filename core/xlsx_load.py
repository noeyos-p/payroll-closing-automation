"""Workbook loading that survives the real-world quirks found in the
2026-09 급여대장 batch (2026-10-07 전체 대조):

  - 사용자 지정 문서 속성(docProps/custom.xml)에 이름 없는 속성이 있으면
    openpyxl이 "StringProperty.name should be str" 오류로 파일 자체를 못
    연다(업체067/업체072/업체073 2개 파일). 급여 데이터와 무관한 부분이라
    그 속성만 건너뛰고 다시 연다.
  - 차트 전용 시트(Chartsheet)는 셀이 없어서 시트를 훑다가
    `'Chartsheet' object has no attribute 'max_column'`으로 죽었다
    (업체074365 급여) — `worksheets_only()`로 일반 시트만 돌려준다.
  - 암호가 걸린 엑셀(OLE 안에 EncryptionInfo)이나 확장자만 .xlsx인 옛 .xls는
    "File is not a zip file"이라는 알 수 없는 오류 대신 무엇을 하면 되는지
    알려주는 메시지로 바꾼다(암호는 우리가 알 수 없으므로 열 수는 없다).
"""

from __future__ import annotations

from pathlib import Path

import openpyxl

_OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


class WorkbookOpenError(ValueError):
    pass


def _describe_non_zip(path: Path) -> str | None:
    try:
        with open(path, "rb") as fh:
            head = fh.read(8)
    except OSError:
        return None
    if head != _OLE_MAGIC:
        return None
    try:
        with open(path, "rb") as fh:
            data = fh.read()
    except OSError:
        data = b""
    if "EncryptionInfo".encode("utf-16-le") in data or b"EncryptedPackage" in data.replace(b"\x00", b""):
        return (
            f"'{path.name}'은(는) 암호가 설정된 엑셀 파일이라 열 수 없습니다. "
            "엑셀에서 열어 [파일 → 정보 → 통합 문서 보호 → 암호 설정]에서 암호를 지우고 "
            "저장한 뒤 다시 올려주세요."
        )
    if "HwpSummaryInformation".encode("utf-16-le") in data:
        return f"'{path.name}'은(는) 확장자만 .xlsx이고 실제로는 한글(HWP) 문서입니다."
    return (
        f"'{path.name}'은(는) 확장자만 .xlsx이고 실제로는 예전 형식(.xls)입니다. "
        "엑셀에서 열어 'Excel 통합 문서(*.xlsx)'로 다른 이름으로 저장한 뒤 다시 올려주세요."
    )


def load_workbook(path: Path, **kwargs):
    """`openpyxl.load_workbook` + the fallbacks described in the module
    docstring. Raises `WorkbookOpenError` (a ValueError) with a readable
    Korean message when the file genuinely can't be opened."""
    try:
        return openpyxl.load_workbook(path, **kwargs)
    except Exception as exc:
        message = str(exc)
        if "StringProperty" in message or "custom" in message.lower():
            from openpyxl.reader import excel as _excel

            original = _excel.ExcelReader.read_custom
            _excel.ExcelReader.read_custom = lambda self: None  # type: ignore[assignment]
            try:
                return openpyxl.load_workbook(path, **kwargs)
            finally:
                _excel.ExcelReader.read_custom = original  # type: ignore[assignment]
        if "zip" in message.lower():
            described = _describe_non_zip(Path(path))
            if described:
                raise WorkbookOpenError(described) from exc
        raise


def worksheets_only(workbook) -> list:
    """Every sheet that actually has cells (skips Chartsheet)."""
    return [ws for ws in workbook.worksheets if hasattr(ws, "max_column")]
