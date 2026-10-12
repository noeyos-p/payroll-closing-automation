"""Local desktop company-list integration; no network or registry side effects.

Read a snapshot with ``read_company_list(path)``, then pass business-number keys
and ClosureResult sequences to ``update_company_list(snapshot, results)``.
Each sequence MUST describe every payroll kind present for that company/month;
an absent kind is not inferred. Repeated results for a kind (split input files)
are complete only when every result is closed. Only H and N are edited.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from hashlib import sha256
from io import BytesIO
import math
import os
from pathlib import Path
import posixpath
import re
import tempfile
from typing import Iterable, Mapping
from uuid import uuid4
from xml.dom import minidom
from xml.etree import ElementTree
from zipfile import ZipFile

from openpyxl import load_workbook


PAYROLL_KINDS = ("급여", "급+상", "추급")
MIDWAY = "중도"  # closed, but a 중도퇴사자 needs handling; counts as done on reruns
_COMPLETE = frozenset(("O", "○", "◯", "⭕", "〇", "⭘", MIDWAY))


class CompanyListError(ValueError):
    """Validation error; messages contain row numbers, never company data."""


class DuplicateCompanyError(CompanyListError):
    pass


class ConcurrentModificationError(CompanyListError):
    pass


def is_completed(value: object) -> bool:
    """Only a standalone O/circle means all completed; O,X is partial."""
    return isinstance(value, str) and value.strip().upper() in _COMPLETE


def normalize_business_number(value: object) -> str | None:
    """Restore zeroes lost by numeric Excel storage, without changing D cells."""
    if value is None or (isinstance(value, str) and not value.strip()):
        return None
    if isinstance(value, bool):
        raise CompanyListError("사업자번호 형식 오류")
    if isinstance(value, (int, float)):
        if not math.isfinite(value) or value != int(value) or value < 0:
            raise CompanyListError("사업자번호 형식 오류")
        digits = str(int(value))
    elif isinstance(value, str):
        digits = re.sub(r"[-\s]", "", value)
    else:
        raise CompanyListError("사업자번호 형식 오류")
    if not re.fullmatch(r"[0-9]{1,10}", digits):
        raise CompanyListError("사업자번호 형식 오류")
    return digits.zfill(10)


@dataclass(frozen=True, repr=False)
class CompanyRow:
    row_number: int
    number: object
    company_name: str
    manager: object
    business_number: str | None
    payment_basis: object
    business_income: object
    withholding_report: object
    payroll_input: object
    report_saved: object
    report_kind: object
    confirmation: object
    notes: object

    @property
    def completed(self) -> bool:
        return is_completed(self.payroll_input)


@dataclass(frozen=True, repr=False)
class CompanyList:
    path: Path
    rows: tuple[CompanyRow, ...]
    duplicate_business_numbers: Mapping[str, tuple[int, ...]]
    duplicate_names: Mapping[str, tuple[int, ...]]
    _source: bytes = field(repr=False)

    def find_company(self, *, business_number: object = None,
                     company_name: str | None = None) -> CompanyRow:
        """Exact identity only; duplicates must be resolved outside this module."""
        if business_number is not None:
            key = normalize_business_number(business_number)
            matches = [r for r in self.rows if key is not None and r.business_number == key]
            if company_name is not None:
                # Do not use a name to silently resolve a duplicate number.
                if len(matches) == 1 and matches[0].company_name.strip() != company_name.strip():
                    matches = []
        elif company_name is not None and company_name.strip():
            matches = [r for r in self.rows if r.company_name.strip() == company_name.strip()]
        else:
            raise CompanyListError("업체 식별 정보가 필요합니다")
        if len(matches) > 1:
            raise DuplicateCompanyError("중복 업체 행: " + ", ".join(str(r.row_number) for r in matches))
        if not matches:
            raise CompanyListError("일치하는 업체 행이 없습니다")
        return matches[0]


@dataclass(frozen=True)
class ClosureResult:
    kind: str
    closed: bool
    reason: str = ""
    # 중도퇴사자 "이름(YYYY-MM-DD)" — a fully closed company with any of these is
    # marked 중도 instead of O (user request 2026-10-12).
    midway: tuple[str, ...] = ()


@dataclass(frozen=True)
class UpdateResult:
    backup_path: Path | None
    changed_rows: tuple[int, ...]
    skipped_rows: tuple[int, ...]


def _load(data: bytes):
    return load_workbook(BytesIO(data), data_only=False, keep_links=True)


def _schema(wb):
    if "업체" not in wb.sheetnames:
        raise CompanyListError("업체 시트가 없습니다")
    ws = wb["업체"]
    for col, label in {1: "번호", 2: "업체명", 4: "사업자번호", 8: "급여입력여부", 14: "비고"}.items():
        if re.sub(r"\s", "", str(ws.cell(1, col).value or "")) != label:
            raise CompanyListError(f"업체 시트 헤더 오류: {col}열")
    return ws


def _duplicates(rows, key):
    groups: dict[str, list[int]] = {}
    for row in rows:
        value = key(row)
        if value:
            groups.setdefault(value, []).append(row.row_number)
    return {value: tuple(numbers) for value, numbers in groups.items() if len(numbers) > 1}


def read_company_list(path: str | Path) -> CompanyList:
    path = Path(path).resolve()
    if path.suffix.lower() != ".xlsx":
        raise CompanyListError("업체리스트는 .xlsx만 지원합니다")
    data = path.read_bytes()
    wb = _load(data)
    try:
        ws = _schema(wb)
        rows = []
        for cells in ws.iter_rows(min_row=2, max_col=14):
            v = [c.value for c in cells]
            if all(v[c] is None or str(v[c]).strip() == "" for c in (1, 3)):
                continue
            try:
                business_number = normalize_business_number(v[3])
            except CompanyListError:
                raise CompanyListError(f"사업자번호 형식 오류: {cells[0].row}행") from None
            if cells[1].data_type == "f":
                raise CompanyListError(f"업체명 수식은 지원하지 않습니다: {cells[0].row}행")
            rows.append(CompanyRow(cells[0].row, v[0], str(v[1] or ""), v[2],
                                   business_number, *v[4:10], v[12], v[13]))
        return CompanyList(path, tuple(rows), _duplicates(rows, lambda r: r.business_number),
                           _duplicates(rows, lambda r: r.company_name.strip()), data)
    finally:
        wb.close()


def _closure_values(results: Iterable[ClosureResult]) -> tuple[str | None, str]:
    groups: dict[str, list[ClosureResult]] = {}
    for result in results:
        if result.kind not in PAYROLL_KINDS or type(result.closed) is not bool:
            raise CompanyListError("급여 종류 또는 마감 상태 오류")
        if not isinstance(result.reason, str):
            raise CompanyListError("마감 사유는 문자열이어야 합니다")
        groups.setdefault(result.kind, []).append(result)
    if not groups:
        raise CompanyListError("존재하는 급여 종류의 결과가 필요합니다")
    flags, labels = [], []
    for kind in PAYROLL_KINDS:
        if kind not in groups:
            continue
        closed = all(r.closed for r in groups[kind])
        flag = "O" if closed else "X"
        flags.append(flag)
        reasons = sorted({" ".join(r.reason.split()) or "미마감" for r in groups[kind] if not r.closed})
        labels.append(f"{kind}: {flag}" + (f" ({' / '.join(reasons)})" if reasons else ""))
    status = "O" if all(f == "O" for f in flags) else (",".join(flags) if "O" in flags else None)
    midway = sorted({m for group in groups.values() for r in group for m in r.midway})
    if status == "O" and midway:
        status = MIDWAY
    note = "[급여마감] " + "; ".join(labels)
    if midway:
        note += "; 중도퇴사: " + ", ".join(midway)
    return status, note


def _check_current(snapshot: CompanyList) -> None:
    try:
        same = snapshot.path.read_bytes() == snapshot._source
    except FileNotFoundError:
        same = False
    if not same:
        raise ConcurrentModificationError("업체리스트가 변경되었습니다. 다시 읽은 후 처리하세요")


def _signature(wb):
    """Check every value/formula and style, including untouched worksheets."""
    return tuple((ws.title, ws.sheet_state, str(ws.freeze_panes),
                  tuple(str(r) for r in ws.merged_cells.ranges),
                  # An all-default style equals "no style": writing a value into a
                  # previously absent cell (unstyled appended row) gets zeros.
                  tuple((c.coordinate, c.value, c.data_type,
                         tuple(c._style) if c._style and any(c._style) else (),
                         c.number_format, c.hyperlink.target if c.hyperlink else None,
                         (c.comment.text, c.comment.author) if c.comment else None)
                        for row in ws for c in row if c.value is not None or c.has_style or c.comment or c.hyperlink),
                  tuple((k, d.width, d.hidden, d.min, d.max) for k, d in ws.column_dimensions.items()),
                  tuple((k, d.height, d.hidden) for k, d in ws.row_dimensions.items()))
                 for ws in wb.worksheets)


def _verify(data: bytes, expected) -> None:
    check = _load(data)
    try:
        if _signature(check) != expected:
            raise CompanyListError("저장 후 값/수식/서식 검증에 실패했습니다")
    finally:
        check.close()


def _patch_cells(source: bytes, rendered: bytes, row_numbers: list[int]) -> bytes:
    """Use openpyxl's H/N cell encoding, keeping every other ZIP part intact.

    A full openpyxl save can discard unsupported drawings/extensions and formula
    caches. Import only the edited cells into the original worksheet XML instead.
    DOM retains the original namespaces, extension elements and untouched cells.
    """
    main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    def sheet_path(archive):
        workbook = ElementTree.fromstring(archive.read("xl/workbook.xml"))
        sheet = next(s for s in workbook.find(f"{{{main}}}sheets") if s.get("name") == "업체")
        relations = ElementTree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        target = next(r.get("Target") for r in relations if r.get("Id") == sheet.get(f"{{{rel}}}id"))
        return posixpath.normpath(target.lstrip("/") if target.startswith("/") else "xl/" + target)

    out = BytesIO()
    with ZipFile(BytesIO(source)) as original, ZipFile(BytesIO(rendered)) as generated:
        target = sheet_path(original)
        dom = minidom.parseString(original.read(target))
        updated = minidom.parseString(generated.read(sheet_path(generated)))
        try:
            old_cells = {c.getAttribute("r"): c for c in dom.getElementsByTagNameNS(main, "c")}
            new_cells = {c.getAttribute("r"): c for c in updated.getElementsByTagNameNS(main, "c")}
            rows = {r.getAttribute("r"): r for r in dom.getElementsByTagNameNS(main, "row")}
            for row_number in row_numbers:
                for column in ("H", "N"):
                    address = f"{column}{row_number}"
                    old_cell = old_cells.get(address)
                    new_cell = new_cells.get(address)
                    if new_cell is None:
                        if old_cell is not None:
                            old_cell.parentNode.removeChild(old_cell)
                        continue
                    cell = dom.importNode(new_cell, deep=True)
                    # The generated nodes use the default spreadsheet namespace;
                    # declare it locally even if the original uses a prefix.
                    cell.setAttribute("xmlns", main)
                    if old_cell is not None:
                        old_cell.parentNode.replaceChild(cell, old_cell)
                    else:
                        row = rows[str(row_number)]
                        following = next((c for c in row.childNodes
                                          if c.nodeType == c.ELEMENT_NODE and c.localName == "c"
                                          and len(c.getAttribute("r").rstrip("0123456789")) == 1
                                          and c.getAttribute("r").rstrip("0123456789") > column), None)
                        # All columns beyond Z follow H/N too.
                        if following is None:
                            following = next((c for c in row.childNodes
                                              if c.nodeType == c.ELEMENT_NODE and c.localName == "c"
                                              and len(c.getAttribute("r").rstrip("0123456789")) > 1), None)
                        row.insertBefore(cell, following) if following else row.appendChild(cell)
            xml = dom.toxml(encoding="utf-8")
            with ZipFile(out, "w") as archive:
                archive.comment = original.comment
                for info in original.infolist():
                    archive.writestr(info, xml if info.filename == target else original.read(info.filename))
        finally:
            dom.unlink()
            updated.unlink()
    return out.getvalue()


def update_company_list(snapshot: CompanyList,
                        results: Mapping[str, Iterable[ClosureResult]]) -> UpdateResult:
    """Atomically replace the snapshot file after validation and a byte backup.

    Backup/staging files are in the same directory. Concurrent changes detected
    before replacement are refused. Like ordinary filesystem compare-and-replace,
    an uncooperative editor can race the final check; Excel locks fail safely.
    Only edited H/N XML cells are transplanted from openpyxl's serialization;
    other ZIP parts (including other sheets, formulas and styles) remain intact.
    """
    _check_current(snapshot)
    wb = _load(snapshot._source)
    staging: Path | None = None
    backup: Path | None = None
    changed, skipped = [], []
    try:
        ws = _schema(wb)
        seen = set()
        for key, closures in results.items():
            row = snapshot.find_company(business_number=key)
            if row.row_number in seen:
                raise DuplicateCompanyError("같은 업체에 여러 갱신 요청이 있습니다")
            seen.add(row.row_number)
            if row.completed:
                skipped.append(row.row_number)
                continue
            status, note = _closure_values(closures)
            h, n = ws.cell(row.row_number, 8), ws.cell(row.row_number, 14)
            if h.data_type == "f" or n.data_type == "f":
                raise CompanyListError(f"갱신 대상 H/N에 수식이 있습니다: {row.row_number}행")
            old_note = str(n.value) if n.value is not None else ""
            new_note = old_note if note in old_note.splitlines() else old_note + ("\n" if old_note else "") + note
            if len(new_note) > 32767:
                raise CompanyListError(f"비고 셀 길이 제한 초과: {row.row_number}행")
            if h.value != status or n.value != new_note:
                h.value, n.value = status, new_note
                n.data_type = "s"  # Existing literal notes may begin with '='.
                changed.append(row.row_number)
        if not changed:
            _check_current(snapshot)
            return UpdateResult(None, (), tuple(skipped))
        expected = _signature(wb)
        fd, name = tempfile.mkstemp(prefix=".payroll-company-", suffix=".xlsx", dir=snapshot.path.parent)
        staging = Path(name)
        with os.fdopen(fd, "wb") as stream:
            rendered = BytesIO()
            wb.save(rendered)
            stream.write(_patch_cells(snapshot._source, rendered.getvalue(), changed))
            stream.flush()
            os.fsync(stream.fileno())
        staged_data = staging.read_bytes()
        _verify(staged_data, expected)
        _check_current(snapshot)
        backup = snapshot.path.with_name(snapshot.path.name + f".{uuid4().hex}.bak")
        with backup.open("xb") as stream:
            stream.write(snapshot._source)
            stream.flush()
            os.fsync(stream.fileno())
        if backup.read_bytes() != snapshot._source:
            raise CompanyListError("백업 검증 실패")
        _check_current(snapshot)
        os.replace(staging, snapshot.path)
        staging = None
        committed = snapshot.path.read_bytes()
        if sha256(committed).digest() != sha256(staged_data).digest():
            raise ConcurrentModificationError("저장 직후 업체리스트가 변경되었습니다. 백업을 확인하세요")
        _verify(committed, expected)
        return UpdateResult(backup, tuple(changed), tuple(skipped))
    finally:
        wb.close()
        if staging is not None:
            staging.unlink(missing_ok=True)
