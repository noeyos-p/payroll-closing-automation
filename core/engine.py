"""Top-level one-shot pipeline: this month's 이지케어 원본 (.xlsx/.xlsm) +
this client's existing 위하고 Sheet1 (.xlsx) -> reconciled 위하고 upload
records + verification/mapping reports (core/reconcile.py).

Unlike the old value-cross-reference engine this replaced, there is no
"try until something reconciles" fallback here — core/sheet_layout.py's
detectors either recognize a sheet's structure from its header text or they
don't. When a workbook has more than one structurally-valid candidate sheet
(e.g. a workbook that bundles a signoff sheet alongside the real ledger, or
several category-split payroll tables — confirmed real case: 주식회사
업체040(본점) has separate "주간보호"/"기타근로소득자"/"가사사업소득자"
sheets, and picking automatically by row count doesn't always land on the
one the accountant actually means), `process()` defaults to the best-guess
candidate (most plausible employee rows) but the caller can pass an explicit
`source_sheet`/`reference_sheet` name — `ui/main_window.py` uses
`list_source_sheet_candidates`/`list_dest_sheet_candidates` to let the user
pick when there's more than one. Excel only for now (.hwp is not wired back
in yet — see CLAUDE.md)."""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path


from core.classify import canon_of_src, group_of_src
from core.gridutil import cell_text
from core.reconcile import ReconcileResult, reconcile
from core.xlsx_load import load_workbook, worksheets_only
from core.sheet_layout import (
    DestLayout,
    SourceLayout,
    detect_dest_layout,
    detect_source_layout,
    find_total_row,
    flatten_source_layout,
    last_nonblank_row,
)

_EXCEL_EXTENSIONS = {".xlsx", ".xlsm"}


def supported_extensions() -> set[str]:
    return set(_EXCEL_EXTENSIONS)


def _approx_data_row_count(worksheet, layout: SourceLayout) -> int:
    last_row = last_nonblank_row(worksheet)
    total_row = find_total_row(worksheet, layout.data_start_row, last_row) or (last_row + 1)
    return max(total_row - layout.data_start_row, 0)


_TOTAL_CANONS = {"지급액계", "공제액계", "차인지급액"}


def _has_any_amount_column(worksheet, layout: SourceLayout) -> bool:
    """A sheet can satisfy `detect_source_layout` (a 순번-like column + a
    성명/이름 column) without being a payroll table at all — confirmed real
    cases:
      - 업체014's workbook has a "계좌 리스트" (bank
        account list) sheet with 순번/은행명/계좌번호/이름/비고 columns, and
        its 순번 column happens to run far longer (blank rows still counted)
        than the real "종사자급여대장" sheet's single employee, so row-count
        ranking alone picked the account list as the "better" candidate.
      - The same workbook also has a "급여지급내역(확인용)" bank-transfer
        summary sheet — 연번/성명/금액/계좌번호/은행명/비고, a single lump
        "금액" per employee, no itemized breakdown at all. A stray "지급합계"
        subtotal row sitting between the header and the first employee row
        (an unusual placement — most files put their total row after all
        employees, not before) got swept into the header block, and its text
        landed on the "성명" column, making it misclassify as a "지급액계"
        (총계) column — which would have passed a group-only check.
    Real payroll tables always have at least one 수당/공제 column carrying
    an actual line item, not just a grand total — requiring the canon to be
    a genuine item (not 지급액계/공제액계/차인지급액) rules out both of the
    above without needing to special-case either shape."""
    return any(
        group_of_src(worksheet, layout, c) in ("수당", "공제")
        and canon_of_src(worksheet, layout, c) not in _TOTAL_CANONS
        for c in range(1, min(worksheet.max_column, 60) + 1)
    )


def _distinct_item_canon_count(worksheet, layout: SourceLayout) -> int:
    """How many DIFFERENT allowance/deduction canon names a sheet covers —
    a real payroll ledger spans many distinct item types (기본급/연장수당/
    야간수당/휴일수당/연차수당/국민연금/...); a narrow single-purpose report
    covers just one or two, repeated across a handful of related columns.
    Confirmed real case: 업체006's workbook also has a
    "선임수당-공단" sheet — a subsidy calculation for ONE allowance (선임
    수당) with its own 건강보험/장기요양/고용보험 employer-cost columns —
    that has MORE data rows than the real "급여대장" sheet (94 vs 67, most
    of them blank template rows) and would otherwise outrank it by row count
    alone, even though it's missing 연장수당/야간수당/휴일수당/연차수당
    entirely."""
    canons = {
        canon_of_src(worksheet, layout, c)
        for c in range(1, min(worksheet.max_column, 60) + 1)
        if group_of_src(worksheet, layout, c) in ("수당", "공제")
        and canon_of_src(worksheet, layout, c) not in _TOTAL_CANONS
    }
    canons.discard("")
    return len(canons)


class _ColumnLimitedSheet:
    """A worksheet view that ends at `max_column` — everything downstream
    (classification, mapping, flatten) only looks at `ws.max_column`, so
    capping it here keeps a separate helper table parked to the right of
    the real ledger out of the run entirely."""

    def __init__(self, worksheet, max_column: int):
        self._ws = worksheet
        self.max_column = max_column
        self.max_row = worksheet.max_row
        self.title = worksheet.title
        self.merged_cells = worksheet.merged_cells

    def cell(self, row: int, column: int):
        return self._ws.cell(row=row, column=column)


def _table_last_col(worksheet, layout: SourceLayout) -> int | None:
    """Last column of the main table: the column before the first run of
    2+ columns that are blank in every header row AND the first data rows.
    2026-10-07 (업체041): the ledger is 58 columns wide but
    the sheet carries 4대보험/세금 calculation tables out to column 280 —
    their 국민/건강/장기/고용 labels were being mapped too, adding billions."""
    rows = list(layout.header_rows) + list(
        range(layout.data_start_row, min(layout.data_start_row + 10 * layout.block_height, worksheet.max_row) + 1)
    )
    start = max(layout.name_col, layout.marker_col or 0) + 1
    repeated_cut = _repeated_block_cut(worksheet, layout)
    if repeated_cut is not None:
        return repeated_cut
    blank_run = 0
    for c in range(start, worksheet.max_column + 1):
        blank = all(worksheet.cell(row=r, column=c).value in (None, "") for r in rows) and all(
            cell_text(worksheet, r, c) in (None, "") for r in layout.header_rows
        )
        if blank:
            blank_run += 1
            if blank_run >= 2:
                return c - 2
        else:
            blank_run = 0
    return None


def _repeated_block_cut(worksheet, layout: SourceLayout) -> int | None:
    """If, after a 공제총액/실수령액-type total column, the sheet starts
    repeating 수당/공제 items it already had (국민/건강/장기/고용, 기본급여,
    식대비 …), that's a helper/calculation table glued on to the right —
    the main table ends at that total column."""
    max_col = min(worksheet.max_column, 400)
    seen: set[str] = set()
    seen_deduction = False
    total_cols: list[int] = []
    canons: dict[int, str] = {}
    raw: dict[int, str] = {}
    for c in range(1, max_col + 1):
        group = group_of_src(worksheet, layout, c)
        ca = canon_of_src(worksheet, layout, c)
        raw[c] = ca
        canons[c] = ca if group in ("수당", "공제") else ""
        if ca in ("공제액계", "차인지급액") and seen_deduction:
            total_cols.append(c)
        if group == "공제" and ca not in _TOTAL_CANONS:
            seen_deduction = True
    for t in total_cols:
        before = {canons[c] for c in range(1, t) if canons[c] and canons[c] not in _TOTAL_CANONS}
        repeats = sum(1 for c in range(t + 1, max_col + 1) if canons[c] in before)
        if repeats >= 3:
            # Keep total columns standing right after the cut: 업체042
            # (2026-10-12) has 공제총액 then 실지급액 before the 사업주부담금 block,
            # and cutting at 공제총액 dropped the 실지급액 column.
            while t + 1 <= max_col and raw.get(t + 1) in _TOTAL_CANONS:
                t += 1
            return t
    return None


def _source_candidates(path: Path) -> list[tuple[object, SourceLayout]]:
    workbook = load_workbook(path, data_only=True)
    candidates = []
    hidden: set[str] = set()
    for worksheet in worksheets_only(workbook):
        layout = detect_source_layout(worksheet)
        if layout is None:
            continue
        if getattr(worksheet, "sheet_state", "visible") != "visible":
            hidden.add(worksheet.title)
        last_col = _table_last_col(worksheet, layout)
        if last_col is not None and last_col < worksheet.max_column:
            worksheet = _ColumnLimitedSheet(worksheet, last_col)
        if layout.needs_flatten:
            # Grid has no sheet name of its own (it isn't backed by an
            # openpyxl worksheet) — the picker UI and `source_sheet=`
            # re-selection both key off `.title`, so it's copied over from
            # the real worksheet it was flattened from.
            grid, layout = flatten_source_layout(worksheet, layout)
            grid.title = worksheet.title
            worksheet = grid
        if _has_any_amount_column(worksheet, layout):
            candidates.append((worksheet, layout))
    # Hidden sheets are helpers, not the ledger: 업체042 (2026-10-12)
    # hides '은행' (same people plus 사원005/사원006 rows), 사원006,
    # 교통비 copies. Drop them whenever a visible candidate exists.
    if any(ws.title not in hidden for ws, _ in candidates):
        candidates = [(ws, layout) for ws, layout in candidates if ws.title not in hidden]
    # Prefer a real "No"/순번 sequence anchor over the name-only fallback
    # (see SourceLayout.has_marker), then how many distinct pay/deduction
    # item TYPES the sheet covers (see `_distinct_item_canon_count`), then
    # the sheet with the most plausible employee rows as a final tiebreaker
    # — a workbook can bundle a real 급여대장 table alongside something else
    # that also happens to satisfy detection (confirmed real case:
    # 업체043's workbook also has a "급여명세서" per-employee
    # pay-stub sheet that repeats "성명" once per slip; picking by header
    # block height alone let that outrank the real table).
    #
    # 2026-10-07 추가 기준 두 개(9월지급 전체 대조에서 확인):
    #  - 이름 칸이 대부분 "#REF!" 같은 오류값인 시트는 맨 뒤로(업체044노인케어
    #    복지센터: 깨진 "급여대장" 시트가 진짜 "급여계산" 시트를 이기고 있었음).
    #  - 파일명이 "8월귀속"인데 시트가 달마다 나뉘어 있으면("4월방문"/"8월방문"
    #    …, "8월 (9)"/"7월 (8)" …) 그 달 이름이 들어간 시트를 먼저.
    month = _accrual_month_from_filename(path)
    # 달마다 시트를 나눈 통합문서(월 이름이 붙은 후보 시트가 2개 이상)에서만
    # 쓴다 — "8월 교통비" 같은 부속 시트 하나 때문에 진짜 급여대장 시트가
    # 밀리면 안 된다(업체042).
    month_named = [ws for ws, _ in candidates if re.search(r"\d{1,2}\s*월", ws.title or "")]
    if len(month_named) < 2:
        month = None
    candidates.sort(
        key=lambda pair: (
            not _names_mostly_broken(pair[0], pair[1]),
            pair[1].has_marker,
            _distinct_item_canon_count(pair[0], pair[1]),
            _sheet_name_has_month(pair[0].title, month),
            # 항목 구성이 같은 시트가 여럿이면(은행/퇴직금/수당/급여대장최종 —
            # 업체042) 같은 이름이 두 번 나오는 시트를 뒤로.
            -_duplicate_name_count(pair[0], pair[1]),
            _approx_data_row_count(pair[0], pair[1]),
        ),
        reverse=True,
    )
    return candidates


def _accrual_month_from_filename(path: Path) -> int | None:
    match = re.search(r"(\d{1,2})\s*월\s*귀속", path.stem)
    return int(match.group(1)) if match else None


def _sheet_name_has_month(title: str, month: int | None) -> bool:
    if month is None:
        return False
    return re.search(rf"(?<!\d)(?:\d{{2}})?0?{month}\s*월", title or "") is not None


def _duplicate_name_count(worksheet, layout: SourceLayout) -> int:
    last_row = last_nonblank_row(worksheet)
    total_row = find_total_row(worksheet, layout.data_start_row, last_row) or (last_row + 1)
    seen: set[str] = set()
    dups = 0
    for r in range(layout.data_start_row, total_row, layout.block_height):
        v = worksheet.cell(row=r, column=layout.name_col).value
        key = str(v).replace(" ", "").strip() if v is not None else ""
        if not key:
            continue
        if key in seen:
            dups += 1
        seen.add(key)
    return dups


def _names_mostly_broken(worksheet, layout: SourceLayout) -> bool:
    values = []
    for i in range(12):
        v = worksheet.cell(row=layout.data_start_row + i * layout.block_height, column=layout.name_col).value
        if v not in (None, ""):
            values.append(str(v))
    if not values:
        return False
    broken = sum(1 for v in values if v.startswith("#") or v.endswith("!"))
    return broken * 2 > len(values)


def _dest_candidates(path: Path) -> list[tuple[object, DestLayout]]:
    workbook = load_workbook(path, data_only=True)
    # "Sheet1" first — that's Wehago's own convention for the upload sheet
    # (confirmed across every real sample in 파일/) — then any other sheet.
    sheets = worksheets_only(workbook)
    ordered = [w for w in sheets if w.title == "Sheet1"] + [w for w in sheets if w.title != "Sheet1"]
    candidates = []
    for worksheet in ordered:
        layout = detect_dest_layout(worksheet)
        if layout is not None:
            candidates.append((worksheet, layout))
    return candidates


@dataclass
class SheetCandidate:
    name: str
    employee_count: int  # best-effort estimate, for the picker's display label


def list_source_sheet_candidates(path: Path) -> list[SheetCandidate]:
    """For `ui/main_window.py`'s sheet picker — every sheet in `path` that
    looks like a 이지케어 원본 table, best guess first (see
    `_source_candidates`'s sort), with an approximate employee count so the
    user can tell a real payroll table from an accidental false positive at
    a glance."""
    return [
        SheetCandidate(worksheet.title, _approx_data_row_count(worksheet, layout))
        for worksheet, layout in _source_candidates(path)
    ]


def list_dest_sheet_candidates(path: Path) -> list[SheetCandidate]:
    result = []
    for worksheet, layout in _dest_candidates(path):
        last_row = last_nonblank_row(worksheet)
        total_row = find_total_row(worksheet, layout.data_start_row, last_row) or (last_row + 1)
        result.append(SheetCandidate(worksheet.title, max(total_row - layout.data_start_row, 0)))
    return result


def list_source_employee_names(path: Path, sheet_name: str | None = None) -> list[str]:
    """이 급여대장(원본)에 있는 사원 이름만 뽑는다 — 사업장 자동 인식(회사명 판별)
    용도. 최선 후보 시트(또는 `sheet_name` 지정 시 그 시트)에서 순번/이름 앵커
    아래 데이터 행을 걸어 이름 칸만 읽는다(core/reconcile.py의 소스 행 순회와
    같은 방식 — `find_total_row`로 합계 행 앞까지, `block_height`만큼 건너뛰며)."""
    candidates = _source_candidates(path)
    if not candidates:
        return []
    worksheet, layout = _pick_candidate(candidates, sheet_name)
    last_row = last_nonblank_row(worksheet)
    total_row = find_total_row(worksheet, layout.data_start_row, last_row) or (last_row + 1)
    names = []
    for r in range(layout.data_start_row, total_row, layout.block_height):
        name = worksheet.cell(row=r, column=layout.name_col).value
        if name:
            names.append(str(name))
    return names


def list_dest_employee_names(path: Path, sheet_name: str | None = None) -> list[str]:
    """위하고 참고파일에 이미 등록된 사원 이름 목록 — 사업장 등록(DB 저장) 시
    이 사업장의 "이름 지문"으로 쓴다(core/registry.py)."""
    candidates = _dest_candidates(path)
    if not candidates:
        return []
    worksheet, layout = _pick_candidate(candidates, sheet_name)
    last_row = last_nonblank_row(worksheet)
    total_row = find_total_row(worksheet, layout.data_start_row, last_row) or (last_row + 1)
    names = []
    for r in range(layout.data_start_row, total_row):
        name = worksheet.cell(row=r, column=layout.name_col).value
        if name:
            names.append(str(name))
    return names


def process(
    client_name: str,
    source_path: Path,
    reference_path: Path,
    *,
    source_sheet: str | None = None,
    reference_sheet: str | None = None,
    manual_overrides: dict[int, str | None] | None = None,
    learned_mappings: dict[str, str] | None = None,
    source_template: tuple[object, Path] | None = None,
) -> ReconcileResult | None:
    """`source_sheet`/`reference_sheet` pin the run to a specific sheet name
    (from `list_source_sheet_candidates`/`list_dest_sheet_candidates`) —
    used when a workbook has more than one structurally-valid candidate and
    the user picked one explicitly. Left as None, the best-guess candidate
    is used, same as before. `manual_overrides`/`learned_mappings` are
    forwarded to `core/reconcile.py`'s `reconcile()` — see its docstring."""
    if source_path.suffix.lower() not in _EXCEL_EXTENSIONS:
        raise ValueError(f"지원하지 않는 파일 형식입니다: {source_path.suffix} (현재는 엑셀 .xlsx/.xlsm만 지원합니다)")
    if reference_path.suffix.lower() not in _EXCEL_EXTENSIONS:
        raise ValueError("위하고 참고파일은 .xlsx만 지원합니다.")

    dest_candidates = _dest_candidates(reference_path)
    if not dest_candidates:
        return None
    dest_ws, dest_layout = _pick_candidate(dest_candidates, reference_sheet)

    # 이 사업장에 "원본 -> 이지케어 표준양식" 변환 규칙이 있으면(사람이 손으로
    # 다시 짠 수정본에서 학습, core/source_template.py) 먼저 그 양식으로 바꿔서
    # 처리한다. 사용자가 시트를 직접 고른 경우나, 규칙대로 바꿀 수 없으면
    # (머리글/시트를 못 찾음) 원래 방식으로 처리한다.
    if source_template is not None and source_sheet is None:
        from core.source_template import apply_recipe

        recipe, template_path = source_template
        try:
            converted = apply_recipe(recipe, source_path, template_path)
        except Exception as exc:
            import core.source_template as _st

            _st.LAST_APPLY_FAILURE = f"변환 중 오류: {exc}"
            converted = None
        if converted is not None:
            ws, layout, notes = converted
            result = reconcile(
                client_name, ws, layout, dest_ws, dest_layout,
                manual_overrides, learned_mappings=learned_mappings,
            )
            result.notes.extend(notes)
            return result

    source_candidates = _source_candidates(source_path)
    if not source_candidates:
        return None
    source_ws, source_layout = _pick_candidate(source_candidates, source_sheet)
    return reconcile(
        client_name, source_ws, source_layout, dest_ws, dest_layout,
        manual_overrides, learned_mappings=learned_mappings,
    )


def _pick_candidate(candidates: list[tuple[object, object]], sheet_name: str | None) -> tuple[object, object]:
    if sheet_name is None:
        return candidates[0]
    for worksheet, layout in candidates:
        if worksheet.title == sheet_name:
            return worksheet, layout
    raise ValueError(f"'{sheet_name}' 시트에서 급여 구조를 인식하지 못했습니다.")
