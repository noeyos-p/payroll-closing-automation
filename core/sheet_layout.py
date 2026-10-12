"""Header/data-row auto-detection for both the 이지케어-style source sheet
and the existing Wehago (Sheet1) destination sheet, replacing the VBA
macro's fixed row constants (SRC_H1=3/SRC_H2=5/SRC_DATA=7, DST_H1=1/
DST_H2=2/DST_DATA=3 in 파일/이지케어_Sheet1_현재시트_수정본_v3.bas).

Those constants only worked because the macro was written against one
specific client's export; real files in 파일/ vary (2-row header ending at
row 4 with data at row 5, or a 3-row header with a blank buffer row and
data at row 6 — see CLAUDE.md for the confirmed real examples). Detection
here finds the "No/이름" or "사원코드/사원명" anchor row, then the actual
first data row by scanning for where the marker/name column's own values
start looking like real data — so the header block can be any height.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from core.canon import norm
from core.grid import Grid
from core.gridutil import cell_text

_CODE_LABELS = {"사원코드", "사원번호", "직원코드", "직원번호", "사번"}
_DST_NAME_LABELS = {"사원명", "직원명", "성명", "이름", "종사자명"}
# "연번" (실사 확인: 업체056) — "No"/"순번"과 뜻은 같지만 다른
# 단어를 쓰는 원본이 있음. "번호" (실사 확인: 업체057) — 마찬가지로
# 뜻은 같지만 다른 단어.
_SRC_MARKER_LABELS = {"no", "순번", "연번", "번호"}
_SRC_NAME_LABELS = {"이름", "성명", "직원명", "사원명", "종사자명"}
_BIRTH_LABELS = {"생년월일", "생년월", "주민등록번호", "주민번호", "주민등록", "주민", "생년"}
_TABLE_COMPANION_LABELS = _BIRTH_LABELS | {
    "근로일수", "근로시간", "직급", "직책", "직종", "호봉", "입사일", "입사일자", "부서", "사번", "급여구분",
}
# Extra header-row words that only `_detect_offset_identity_source_layout`
# and its `flatten_source_layout` counterpart need to recognize as "still
# header territory, not a person's name" — confirmed real case:
# 업체058's 원본 reuses column E across offsets for 급여유형/
# 성명/인력구분, and "인력구분"/"지급일자" are short enough Korean text to
# otherwise pass `_looks_like_person_name`'s loose shape check. Kept OUT of
# `_TABLE_COMPANION_LABELS` deliberately: that set also feeds the unrelated
# "no marker, name+companion" fallback's own anchor-row confirmation
# earlier in `detect_source_layout`, where adding these same words would
# let that fallback claim this exact anchor row too (with a much worse,
# non-flattened layout) before `_detect_offset_identity_source_layout` — a
# strictly later fallback — ever gets a turn.
_OFFSET_IDENTITY_NAME_EXCLUSIONS = {"인력구분", "지급일자", "급여유형", "직접비", "간접비"}
_NAME_BLACKLIST = {
    "합계", "사원명", "직원명", "성명", "이름", "사원코드", "사원번호", "직원코드", "직원번호", "사번",
    "부서", "직급", "직종", "no", "순번",
}
_MAX_ANCHOR_SCAN_ROWS = 10
_MAX_HEADER_SCAN_ROWS = 15


def _looks_like_person_name(value: object) -> bool:
    s = str(value if value is not None else "").strip()
    if not s or norm(s) in _NAME_BLACKLIST:
        return False
    core_chars = re.sub(r"[()0-9./\-\s]", "", s)
    if not (1 <= len(core_chars) <= 8):
        return False
    return any("가" <= ch <= "힣" for ch in core_chars)


def _is_plausible_employee_name(value: object) -> bool:
    """`_looks_like_person_name` plus one more exclusion: a header LABEL
    word that happens to be short Korean text (e.g. "생년월일") also passes
    that loose shape check — harmless everywhere `_looks_like_person_name`
    is normally called (right below a confirmed anchor row, values start
    immediately), but not when scanning several rows down looking for where
    real data actually begins, where a lower header row's own label can be
    mistaken for a person's name and short-circuit the scan too early
    (confirmed real case: 급여명세대장(2026년7월).xlsx has "이름" on row 1 but
    "생년월일" — not a name — directly in the same column two rows down, on
    what's still a header row)."""
    if norm(str(value if value is not None else "")) in _TABLE_COMPANION_LABELS:
        return False
    return _looks_like_person_name(value)


def _is_plausible_employee_name_offset(value: object) -> bool:
    """`_is_plausible_employee_name` plus `_OFFSET_IDENTITY_NAME_EXCLUSIONS`
    — used only by `_detect_offset_identity_source_layout` and its
    `flatten_source_layout` counterpart, never by the ordinary detection
    paths (see that set's docstring for why it's kept separate from
    `_TABLE_COMPANION_LABELS`)."""
    if norm(str(value if value is not None else "")) in _OFFSET_IDENTITY_NAME_EXCLUSIONS:
        return False
    return _is_plausible_employee_name(value)


def _as_positive_int(value: object) -> int | None:
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        n = value
    elif isinstance(value, str):
        try:
            n = float(value.replace(",", "").strip())
        except ValueError:
            return None
    else:
        return None
    return int(n) if n > 0 and float(n).is_integer() else None


@dataclass
class DestLayout:
    header_row1: int  # anchor row: 사원코드/사원명/... group labels (e.g. "수당")
    header_row2: int | None  # leaf item labels; None if the header is only 1 row
    data_start_row: int
    name_col: int
    code_col: int | None
    birth_col: int | None


@dataclass
class SourceLayout:
    header_rows: list[int]  # anchor row .. row right before data, in order
    data_start_row: int
    name_col: int
    marker_col: int
    birth_col: int | None
    mode: str = "row"  # "row" (1 employee = 1 row) — see SourceLayoutColumnar for the other mode
    # How many physical rows make up one employee's block — 1 normally, or
    # more when the ledger spreads one employee's identity/detail fields
    # across a couple of rows (confirmed real case: 업체043's
    # 급여대장 sheet — row 1 of a block has 순번/성명/생년월일/기본급 etc.,
    # row 2 has 담당업무/입사일/부양가족 with the marker column blank).
    # core/reconcile.py sums every block row's numeric cells together and
    # only reads the employee's name off the block's first row.
    block_height: int = 1
    # True when a real "No"/"순번" sequence column anchored detection (the
    # strict, higher-confidence path); False when there was no such column
    # and detection fell back to just the name column's own values looking
    # like real people (confirmed real case: 업체043's 급여대장
    # sheet has no 순번 column at all). core/engine.py prefers strict
    # candidates when a workbook has several structurally-plausible sheets
    # (e.g. an individual pay-stub sheet that repeats "성명" once per
    # employee can otherwise out-rank the real table).
    has_marker: bool = True
    # True when block_height > 1 because the SAME physical columns are
    # reused across the block's rows for DIFFERENT items — one header row
    # per offset (confirmed real case: 업체016, see
    # `flatten_source_layout`) — rather than the safe-to-sum continuation-
    # row shape. `core/engine.py` must call `flatten_source_layout(ws,
    # layout)` before using this layout any further; the row/col numbers
    # here are meaningless on their own until that rewrite happens.
    needs_flatten: bool = False
    # Physical row offset (within a block) where marker_col/name_col/
    # birth_col's own values actually live -- 0 for every previously
    # confirmed needs_flatten case (identity fields sit on the block's
    # first physical row), but confirmed real case 업체058
    # has identity fields (직원번호/성명) one row further down than item
    # columns' own offset 0, because the identity columns (A/E) are
    # independently relabeled on EVERY header row (직종/급여유형 -> offset
    # 0, 직원번호/성명 -> offset 1, 지급일자/인력구분 -> offset 2) while the
    # item columns only start their own offset scheme one row later. See
    # `flatten_source_layout`'s identity rescan.
    identity_offset: int = 0
    # How many of `header_rows`' LAST rows are the item columns' own
    # per-offset labels -- normally equal to `block_height` (every
    # physical block row has its own header row), but confirmed real case
    # 업체058 has one leading header row (직종/급여유형 group
    # row) that only labels the identity columns at offset 0, not any
    # item column, so `header_rows` has one MORE row than block_height.
    # None (the default) means "same as block_height", the previous
    # implicit assumption.
    item_offset_rows: int | None = None


@dataclass
class SourceLayoutColumnar:
    """Best-effort structure for a 열 기반 (employee-as-column) source sheet
    — no confirmed real Excel sample exists yet (see CLAUDE.md); this exists
    so such a file gets a shot at automatic processing instead of an
    outright "지원하지 않는 형식" failure, but is not value-verified."""

    label_col: int  # column holding item labels (성명/기본급/...) on each row
    name_row: int
    item_rows: list[int]
    block_start_col: int
    block_width: int
    block_count: int
    mode: str = "column"


def last_nonblank_row(ws, max_col: int = 30) -> int:
    """Port of VBA `LastRow` — last row with any non-empty cell in the first
    `max_col` columns (avoids openpyxl's `ws.max_row`, which can overcount
    due to formatting applied to otherwise-empty trailing rows)."""
    scan_cols = min(max_col, ws.max_column)
    for r in range(ws.max_row, 0, -1):
        if any(ws.cell(row=r, column=c).value not in (None, "") for c in range(1, scan_cols + 1)):
            return r
    return 1


def find_total_row(ws, start_row: int, end_row: int, max_col: int = 5) -> int | None:
    """Port of VBA `FindTotal` — scans the first few columns of each row for
    a literal "합계" label marking the totals row.

    Also accepts "합 계 (총 1명)"/"총계"/"총합계" — confirmed real case
    (2026-10-07): 업체059's sheet ends its first table with
    "합   계 (총  1명)" and then repeats the same employees in a second
    table further down; missing the totals row made every employee look
    like a 동명이인 and nobody got matched."""
    for r in range(start_row, end_row + 1):
        for c in range(1, min(max_col, ws.max_column) + 1):
            key = norm(cell_text(ws, r, c))
            if key in ("합계", "총계", "총합계") or (key.startswith("합계") and key[2:3] in ("총", "")):
                return r
    return None


def detect_dest_layout(ws) -> DestLayout | None:
    max_col = ws.max_column
    max_row = ws.max_row
    anchor_row = code_col = name_col = None
    for r in range(1, min(_MAX_ANCHOR_SCAN_ROWS, max_row) + 1):
        row_code_col = row_name_col = None
        for c in range(1, max_col + 1):
            label = norm(cell_text(ws, r, c))
            if label in _CODE_LABELS and row_code_col is None:
                row_code_col = c
            if label in _DST_NAME_LABELS and row_name_col is None:
                row_name_col = c
        if row_code_col is not None and row_name_col is not None:
            anchor_row, code_col, name_col = r, row_code_col, row_name_col
            break
    if anchor_row is None:
        return None

    # Header is 1 or 2 rows: if the row right after the anchor already looks
    # like real employee data, there's no separate leaf-label row.
    probe = cell_text(ws, anchor_row + 1, name_col)
    if _looks_like_person_name(probe):
        header_row2 = None
        data_start_row = anchor_row + 1
    else:
        header_row2 = anchor_row + 1
        data_start_row = anchor_row + 2
        # Scan forward a little in case there's more than one leaf-label row.
        for r in range(data_start_row, min(data_start_row + _MAX_HEADER_SCAN_ROWS, max_row) + 1):
            if _looks_like_person_name(cell_text(ws, r, name_col)):
                data_start_row = r
                break
            header_row2 = r

    birth_col = None
    for c in range(1, max_col + 1):
        for r in (anchor_row, header_row2) if header_row2 else (anchor_row,):
            if norm(cell_text(ws, r, c)) in _BIRTH_LABELS:
                birth_col = c
                break
        if birth_col:
            break

    return DestLayout(
        header_row1=anchor_row,
        header_row2=header_row2,
        data_start_row=data_start_row,
        name_col=name_col,
        code_col=code_col,
        birth_col=birth_col,
    )


def detect_source_layout(ws) -> SourceLayout | None:
    """Tries the ordinary detection paths first, and only reaches for
    `_detect_offset_identity_source_layout`'s narrower, more specific shape
    (identity columns relabeled per-offset like item columns, one row
    deeper — see that function's docstring) once every other path has come
    up empty, so it can never pre-empt a file that already matches a more
    common, better-established pattern."""
    layout = _detect_source_layout_primary(ws)
    if layout is not None and not getattr(layout, "relaxed_anchor", False):
        return layout
    # 동반 단어 없이 급여 항목명만으로 잡은 앵커(2026-10-07 완화)는 더 구체적인
    # 업체058형 양식이 맞으면 그쪽을 우선한다.
    specific = _detect_offset_identity_source_layout(ws)
    return specific if specific is not None else layout


def _detect_source_layout_primary(ws) -> SourceLayout | None:
    max_col = ws.max_column
    max_row = ws.max_row
    anchor_row = marker_col = name_col = None
    for r in range(1, min(_MAX_ANCHOR_SCAN_ROWS, max_row) + 1):
        row_marker_col = row_name_col = None
        for c in range(1, min(max_col, 10) + 1):
            label = norm(cell_text(ws, r, c))
            if label in _SRC_MARKER_LABELS and row_marker_col is None:
                row_marker_col = c
        for c in range(1, min(max_col, 10) + 1):
            label = norm(cell_text(ws, r, c))
            if label in _SRC_NAME_LABELS and row_name_col is None:
                row_name_col = c
        if row_marker_col is not None and row_name_col is not None:
            anchor_row, marker_col, name_col = r, row_marker_col, row_name_col
            break

    if anchor_row is not None:
        data_start_row = None
        search_end = min(anchor_row + _MAX_HEADER_SCAN_ROWS, max_row)
        for r in range(anchor_row + 1, search_end + 1):
            if _as_positive_int(ws.cell(row=r, column=marker_col).value) == 1:
                data_start_row = r
                break
        if data_start_row is not None:
            # A labeled "No"/순번 column doesn't guarantee 1 employee = 1
            # physical row. Confirmed real case: 업체016's 원본
            # merges 연번/성명 (and vertically stacks values) across 3 rows
            # per employee, with the marker's own raw value only ever
            # present on the block's first row — reading it as
            # block_height=1 anyway silently split each employee into 3
            # fake ones with scrambled amounts.
            #
            # The gap to the next marker value (2) tells us the block
            # height. Unlike the one validated safe-to-sum multi-row case
            # (업체043, found via the unlabeled fallback below —
            # its continuation row's amount columns are blank/non-numeric,
            # so summing harmlessly adds zero), 업체016 reuses the
            # SAME column for block_height DIFFERENT, real, numeric
            # quantities depending on which row-offset you're on (e.g.
            # column M is 기본급 on offset 0, 주휴수당 on offset 1, and a
            # 사대보험 환급 on offset 2) — summing is a different kind of
            # wrong number that slips past the accountant unnoticed,
            # because the block's OWN stated totals also only live on the
            # first row, so 지급액계 등 still happens to "match" while every
            # individual line item is garbage.
            #
            # The only way this app can process such a shape safely is if
            # it can pin down, per offset, which header row actually labels
            # that offset's values — that's only true when there's exactly
            # one header row per offset (`flatten_source_layout` below
            # handles the rewrite). Any other shape (fewer/more header rows
            # than offsets) has no reliable way to tell "safe to sum" from
            # "column reused for different items" apart, so it still
            # refuses outright rather than guess.
            block_height = 1
            block_search_end = min(data_start_row + 20, max_row)
            for r in range(data_start_row + 1, block_search_end + 1):
                if _as_positive_int(ws.cell(row=r, column=marker_col).value) == 2:
                    block_height = r - data_start_row
                    break
            layout = _finish_source_layout(
                ws, anchor_row, data_start_row, name_col, marker_col, has_marker=True, block_height=block_height
            )
            if block_height == 1:
                return layout
            # The offset-specific sub-header rows are always the LAST
            # `block_height` header rows, directly above the data — anything
            # earlier is a pure super-group label not tied to any one
            # offset (confirmed real case: 업체016's row 4, "인적
            # 사항"/"기본사항"/"지급사항"/... merged across many columns,
            # got pulled into header_rows by the upward-extension logic in
            # `_finish_source_layout` because a merge happened to make it
            # "cover" one blank spacer column rows 5-7 don't otherwise
            # label — but rows 5/6/7 themselves already carry specific item
            # labels on their own, so that extra row isn't needed here).
            if len(layout.header_rows) >= block_height:
                layout.needs_flatten = True
                return layout
            return None

    # No "No"/"순번" column at all (confirmed real case: 업체043's
    # 급여대장 sheet lists "성  명" with no sequence column) — fall back to
    # a name-only anchor, and find the first data row by looking for where
    # the name column's own values start looking like real people instead
    # of counting a numeric marker. Requires the anchor row to also carry at
    # least one other table-header word nearby (생년월일/근로일수/직급/...) —
    # a genuine table header row lists several column headers side by side,
    # whereas a per-employee pay-STUB sheet's "성명" cell is a single label
    # immediately followed by that one person's name as a VALUE (confirmed
    # real case: the same workbook's "급여명세서" sheet — row 5 is `성명 |
    # (blank) | 김민송 | (blank) | (blank) | 급여지급일 | ...`, no second
    # table-header word next to "성명", which otherwise passed as a false
    # table anchor and misread ~30 payslips as ~30 fabricated "employees").
    anchor_row = name_col = None
    for r in range(1, min(_MAX_ANCHOR_SCAN_ROWS, max_row) + 1):
        row_name_col = None
        for c in range(1, min(max_col, 10) + 1):
            if norm(cell_text(ws, r, c)) in _SRC_NAME_LABELS:
                row_name_col = c
                break
        if row_name_col is None:
            continue
        # A row carrying "사원코드"/"사원번호" alongside the name AND a
        # "수당"/"공제" GROUP-spanning label is a Wehago DESTINATION-shaped
        # header, not a 이지케어 source ledger — confirmed real case:
        # 업체060's "원본" file is itself already Wehago-shaped
        # (사원코드/사원명/부서/직급/직종/수당 at row 2, "수당" being the
        # group header spanning the amount columns), which otherwise passed
        # this fallback too (부서/직급/직종 all count as "table companion"
        # words) and got misread as a source ledger. The code-label alone
        # isn't a safe enough signal on its own — "사번" in particular is
        # just the generic Korean word for "employee number" and shows up
        # on plenty of genuine 이지케어 exports too (confirmed real case:
        # 업체022's 원본 has a single flat header row — 사번/성명/
        # 직위/부서/입사일/퇴사일/기본급/연장수당/.../차인지급액 — with each
        # amount column already its own specific leaf item, no "수당"/"공제"
        # group row above them anywhere; excluding it on "사번" alone left
        # it with zero recognized source candidates).
        has_dest_code_label = any(
            norm(cell_text(ws, r, c)) in _CODE_LABELS for c in range(1, min(max_col, 15) + 1)
        )
        has_group_label_nearby = any(
            norm(cell_text(ws, rr, c)) in ("수당", "공제")
            for rr in range(max(1, r - 2), r + 1)
            for c in range(1, max_col + 1)
        )
        if has_dest_code_label and has_group_label_nearby:
            continue
        # The companion word doesn't have to sit on the exact same row as
        # "이름"/"성명" — a genuine multi-row header can spread its column
        # labels across a couple of rows the same way `_finish_source_layout`
        # already extends header_rows upward for a group-label row (confirmed
        # real case: 급여명세대장(2026년7월).xlsx has "이름" alone on row 1,
        # with "직종"/"생년월일" one and two rows below it — the name row
        # itself carries no companion word, so the same-row-only check found
        # nothing and detection failed outright). Capped at 2 rows below,
        # not a wide scan: the 급여명세서 false-positive this guard exists for
        # is a person's actual VALUE sitting right next to "성명" on that same
        # row, immediately followed by dozens of unrelated payslip detail rows
        # for that one person — not a compact 2-3 row header block — so a
        # short window stays safe without reopening that false match.
        has_companion_header = any(
            norm(cell_text(ws, rr, c)) in _TABLE_COMPANION_LABELS
            for rr in range(r, min(r + 2, max_row) + 1)
            for c in range(1, min(max_col, 15) + 1)
        )
        relaxed_anchor = not has_companion_header
        if has_companion_header or _pay_item_label_count(ws, r, min(r + 2, max_row), max_col) >= 3:
            # 2026-10-07 (업체054): "성명 | 기본급 | 주휴수당 | 연차수당 |
            # … | 국민연금 | 건강보험"처럼 동반 단어(생년월일 등)는 없어도 급여
            # 항목이 여러 개 나란히 있으면 그 자체가 진짜 표 머리글이다 —
            # 명세서 시트의 "성명 | 김민송" 한 줄과는 확실히 구분된다.
            anchor_row, name_col = r, row_name_col
            anchor_relaxed = relaxed_anchor
            break
    if anchor_row is None:
        return None

    # Before trusting text-heuristic name detection (which false-positives
    # on short Korean phrases like job titles — see below), look for an
    # unlabeled but genuine 순번 sequence near the name column; finding one
    # is strong, unambiguous evidence of both the true data start row and
    # the employee block height (confirmed real case: 업체043's
    # col 1 has no "No"/"순번" header text, but does hold a real 1..N
    # sequence every 2 rows — each employee's row is followed by a
    # 담당업무/입사일 continuation row with a blank marker cell).
    wide_search_end = min(anchor_row + 500, max_row)
    sequence = _detect_value_sequence_marker(
        ws, anchor_row + 1, wide_search_end, range(1, min(max_col, 10) + 1)
    )
    if sequence is not None:
        marker_col, data_start_row, block_height = sequence
        layout = _finish_source_layout(
            ws, anchor_row, data_start_row, name_col, marker_col, has_marker=True, block_height=block_height
        )
        # 2026-10-07: 업체034/업체035 양식은 직원 1명이 2줄인데 2번째 줄이
        # 고용보험/소득세/연차수당처럼 1번째 줄과 "다른 항목"이다(헤더도
        # 7행/8행에 줄마다 따로 적혀 있음). 이걸 합산하면 국민연금+고용보험이
        # 한 칸에 섞인다 — 헤더가 줄마다 다른 항목을 적고 있으면 펼친다.
        if (
            block_height > 1
            and len(layout.header_rows) >= block_height
            and _looks_like_flattenable_header(ws, layout.header_rows[-block_height:], max_col)
        ):
            layout.needs_flatten = True
        layout.relaxed_anchor = anchor_relaxed  # type: ignore[attr-defined]
        return layout

    data_start_row = None
    search_end = min(anchor_row + _MAX_HEADER_SCAN_ROWS, max_row)
    for r in range(anchor_row + 1, search_end + 1):
        if _is_plausible_employee_name(ws.cell(row=r, column=name_col).value):
            data_start_row = r
            break
    if data_start_row is None:
        return None

    # No marker column exists here at all (this branch only runs once both
    # the labeled-marker and the unlabeled-numeric-sequence paths above have
    # already failed), so block height can't be read off any counter the way
    # `flatten_source_layout`'s labeled-marker caller does. Instead, infer it
    # straight from the header text itself: if the header block's rows
    # broadly relabel the SAME physical columns with different item names —
    # rather than a normal group>leaf hierarchy, where the group row's text
    # repeats across the columns it spans — that's the same "one header row
    # per block offset" shape `flatten_source_layout` already knows how to
    # unpack (confirmed real case: 급여명세대장(2026년7월).xlsx has no
    # 순번/No column anywhere, but its 3 header rows relabel columns like F
    # 기본급/주휴수당/연차수당 — genuinely 3 different pay items reusing one
    # column across a 3-physical-row employee block).
    header_rows_prelim = list(range(anchor_row, data_start_row))
    block_height = len(header_rows_prelim) if _looks_like_flattenable_header(ws, header_rows_prelim, max_col) else 1
    name_col_relabeled = sum(
        1 for r in header_rows_prelim if isinstance(ws.cell(row=r, column=name_col).value, str)
        and ws.cell(row=r, column=name_col).value.strip()
    ) >= block_height
    if (
        block_height > 1
        and not name_col_relabeled
        and not _names_follow_block_height(ws, name_col, data_start_row, block_height)
    ):
        # 2026-10-07 (업체054): 헤더가 두 줄이라는 것만으로 블록으로
        # 보면 안 된다 — 이름이 매 줄마다 있으면(사원011, 사원009, …) 1명=1줄이다.
        block_height = 1

    layout = _finish_source_layout(
        ws, anchor_row, data_start_row, name_col, name_col, has_marker=False, block_height=block_height
    )
    if block_height == 1:
        layout.relaxed_anchor = anchor_relaxed  # type: ignore[attr-defined]
        return layout
    if len(layout.header_rows) >= block_height:
        layout.needs_flatten = True
        layout.relaxed_anchor = anchor_relaxed  # type: ignore[attr-defined]
        return layout
    return None


def _detect_offset_identity_source_layout(ws) -> SourceLayout | None:
    """Last-resort detector for a shape none of the paths above handle: no
    "No"/순번 column anywhere (직원번호 holds real, non-sequential employee
    codes), AND the identity columns themselves are relabeled on every
    header row exactly like item columns are, one row further UP than
    where item columns' own labeling starts — confirmed real case:
    업체058's 원본. There, column A reads 직종(row4, offset 0)
    / 직원번호(row5, offset 1) / 지급일자(row6, offset 2), column E reads
    급여유형(row4) / 성명(row5) / 인력구분(row6), while the item columns
    (기본급 등) only start their own per-offset labels at row5 (offset 0)
    — row4 is pure section-group text for them (기본급여 및 제수당/제공제,
    each a single wide merge spanning many item columns at once). So the
    identity fields' real values sit ONE row deeper into each employee's
    block than the item fields' offset 0 does — `identity_offset` records
    that gap, and `flatten_source_layout`'s identity rescan uses it instead
    of assuming identity always lives on the block's very first row.

    Requires BOTH a name label (성명/이름/...) and an employee-code label
    (직원번호/사원번호/...) on the SAME row, with a genuinely different
    label sitting directly above BOTH of those exact columns (not just
    anywhere on the row above) — that's what tells apart a real multi-row
    identity header from a stray "성명" cell elsewhere in the sheet.
    """
    max_col = ws.max_column
    max_row = ws.max_row
    for r in range(2, min(_MAX_ANCHOR_SCAN_ROWS, max_row) + 1):
        row_name_col = row_code_col = None
        for c in range(1, min(max_col, 20) + 1):
            label = norm(cell_text(ws, r, c))
            if label in _SRC_NAME_LABELS and row_name_col is None:
                row_name_col = c
            if label in _CODE_LABELS and row_code_col is None:
                row_code_col = c
        if row_name_col is None or row_code_col is None:
            continue

        header_top = r - 1
        above_name = cell_text(ws, header_top, row_name_col)
        above_code = cell_text(ws, header_top, row_code_col)
        if not (isinstance(above_name, str) and above_name.strip()):
            continue
        if not (isinstance(above_code, str) and above_code.strip()):
            continue
        if norm(above_name) in _SRC_NAME_LABELS or norm(above_code) in _CODE_LABELS:
            continue  # just a repeated/merged label, not a distinct sibling offset row

        identity_offset = r - header_top  # always 1 given header_top = r - 1

        search_end = min(r + _MAX_HEADER_SCAN_ROWS, max_row)
        first_name_row = None
        for rr in range(r, search_end + 1):
            if _is_plausible_employee_name_offset(ws.cell(row=rr, column=row_name_col).value):
                first_name_row = rr
                break
        if first_name_row is None:
            continue
        data_start_row = first_name_row - identity_offset
        if data_start_row <= header_top:
            continue

        # Block height: the identity column is reused for other offsets'
        # unrelated text within the same block (지급일자/인력구분 등), which
        # can itself pass the loose "looks like a name" shape check — so
        # instead of trusting the very next hit, require the SAME step
        # size to keep landing on plausible names for several employees
        # running, the same confidence bar `_detect_value_sequence_marker`
        # applies to a numeric 순번 sequence.
        block_height = None
        for period in range(2, 21):
            consecutive = 0
            for k in range(1, 7):
                probe = first_name_row + k * period
                if probe > max_row:
                    break
                if _is_plausible_employee_name_offset(ws.cell(row=probe, column=row_name_col).value):
                    consecutive += 1
                else:
                    break
            if consecutive >= 3:
                block_height = period
                break
        if block_height is None or block_height <= identity_offset:
            continue

        layout = SourceLayout(
            header_rows=[header_top] + [r + k for k in range(block_height)],
            data_start_row=data_start_row,
            name_col=row_name_col,
            marker_col=row_code_col,
            birth_col=None,
            has_marker=False,
            block_height=block_height,
            needs_flatten=True,
            identity_offset=identity_offset,
            item_offset_rows=block_height,
        )
        return layout
    return None


def _names_follow_block_height(ws, name_col: int, data_start_row: int, block_height: int) -> bool:
    """True when the name column really looks like one name per
    `block_height` rows (name on each block's first row, none on the rows
    in between) for the first few blocks."""
    checked = 0
    for b in range(4):
        start = data_start_row + b * block_height
        if not _is_plausible_employee_name(ws.cell(row=start, column=name_col).value):
            break
        for k in range(1, block_height):
            if _is_plausible_employee_name(ws.cell(row=start + k, column=name_col).value):
                return False
        checked += 1
    return checked >= 2


_PAY_ITEM_CANONS = {
    "기본급", "주휴수당", "연차수당", "연장근로수당", "야간수당", "휴일근로수당", "식대", "차량유지비",
    "국민연금", "건강보험", "고용보험", "장기요양보험료", "소득세", "지방소득세",
}


def _pay_item_label_count(ws, first_row: int, last_row: int, max_col: int) -> int:
    from core.canon import canon

    found = set()
    for r in range(first_row, last_row + 1):
        for c in range(1, min(max_col, 60) + 1):
            text = cell_text(ws, r, c)
            if isinstance(text, str) and text.strip():
                ca = canon(text)
                if ca in _PAY_ITEM_CANONS:
                    found.add(ca)
    return len(found)


def _looks_like_flattenable_header(ws, header_rows: list[int], max_col: int) -> bool:
    """True if a broad majority of physical columns carry a genuinely
    different raw label on EVERY one of `header_rows` — see the call site's
    comment for the real case this exists for. Reads raw cell values, not
    merge-aware `cell_text`: a column that's really just ONE item spanning
    the whole header block (e.g. 이름/연번) has its header cells merged, so
    only the anchor cell has a value and the others read blank here — which
    correctly keeps it out of the "reused" count, whereas genuinely distinct
    per-offset item columns (기본급/주휴수당/연차수당, unmerged) all count.
    A normal group>leaf hierarchy fails this instead: the group row's text
    repeats identically across every column it spans rather than being
    unique per column, so those columns don't count as "reused" here either."""
    height = len(header_rows)
    if height < 2:
        return False
    reused_cols = 0
    for c in range(1, max_col + 1):
        labels = []
        for r in header_rows:
            raw = ws.cell(row=r, column=c).value
            if isinstance(raw, str) and raw.strip():
                labels.append(raw.strip())
        if len(labels) == height and len(set(labels)) == height:
            reused_cols += 1
    return reused_cols >= max(3, max_col // 4)


def _text_cols(ws, row: int, max_col: int) -> set[int]:
    return {c for c in range(1, max_col + 1) if isinstance(cell_text(ws, row, c), str) and cell_text(ws, row, c).strip()}


def _finish_source_layout(
    ws,
    anchor_row: int,
    data_start_row: int,
    name_col: int,
    marker_col: int,
    *,
    has_marker: bool,
    block_height: int = 1,
) -> SourceLayout:
    max_col = ws.max_column

    # A row directly above the marker/name anchor row can still be a real
    # per-column group header, not just document metadata — confirmed real
    # case: 업체056's 원본 splits "연번/종사자명/.../급여/급여총계
    # /개인 공제액/..." across TWO rows, with the anchor row (연번+종사자명
    # together) sitting BELOW the row that actually labels the money columns
    # — that row is blank in the 연번/종사자명/직급 columns entirely, so
    # anchoring there alone loses "급여"'s label completely (its column's
    # item text comes out blank and the column is never recognized as
    # anything). Only extend upward while the candidate row labels columns
    # the already-known header rows leave BLANK — a document-metadata line
    # ("사업자번호 / 서비스종류" in 2 unrelated cells — confirmed real case:
    # the same workbook's "급여지급내역(확인용)" sheet) sits in columns the
    # header already covers, so it correctly doesn't qualify.
    header_start = anchor_row
    covered_cols: set[int] = set()
    for r in range(anchor_row, data_start_row):
        covered_cols |= _text_cols(ws, r, max_col)
    while header_start > 1:
        candidate_cols = _text_cols(ws, header_start - 1, max_col)
        # 한 셀을 가로로 길게 병합한 제목/기관정보 줄("장기요양기관기호 : … 기관명 :
        # …")은 병합 덕분에 여러 열에 글자가 있는 것처럼 보일 뿐 머리글이 아니다
        # (업체054, 2026-10-07) — 서로 다른 글자가 2개 이상이어야 한다.
        distinct_texts = {str(cell_text(ws, header_start - 1, c)).strip() for c in candidate_cols}
        if len(candidate_cols) >= 2 and len(distinct_texts) >= 2 and candidate_cols - covered_cols:
            header_start -= 1
            covered_cols |= candidate_cols
        else:
            break

    birth_col = None
    for r in range(anchor_row, data_start_row):
        for c in range(1, max_col + 1):
            if norm(cell_text(ws, r, c)) in _BIRTH_LABELS:
                birth_col = c
                break
        if birth_col:
            break

    return SourceLayout(
        header_rows=list(range(header_start, data_start_row)),
        data_start_row=data_start_row,
        name_col=name_col,
        marker_col=marker_col,
        birth_col=birth_col,
        has_marker=has_marker,
        block_height=block_height,
    )


def _vertical_merge_label(ws, row: int, col: int) -> str:
    """Text of a merge that covers (row, col) coming straight DOWN from a
    cell above in the same column — not a horizontal merge from the left."""
    merged_cells = getattr(ws, "merged_cells", None)
    if merged_cells is None or ws.cell(row=row, column=col).value is not None:
        return ""
    for rng in merged_cells.ranges:
        if rng.min_row < row <= rng.max_row and rng.min_col == col == rng.max_col:
            text = ws.cell(row=rng.min_row, column=rng.min_col).value
            return text.strip() if isinstance(text, str) else ""
    return ""


def flatten_source_layout(ws, layout: SourceLayout) -> tuple[object, SourceLayout]:
    """Rewrites a `needs_flatten` layout (see `SourceLayout.needs_flatten`)
    into a plain single-row-per-employee grid. `layout.header_rows` has
    exactly `block_height` rows (checked by the caller before setting
    `needs_flatten`), one per physical row-offset within a block, in the
    same top-to-bottom order as the offsets themselves — confirmed real
    case: 업체016's `header_rows=[5,6,7]` label offsets 0/1/2 of
    each 3-row block in that order (row 5 = 기본급/생년월일/..., row 6 =
    주휴수당/입사일/..., row 7 = 사대보험 환급/나이(만)/...). A physical
    column therefore becomes up to `block_height` separate logical columns
    — one per offset that actually has a label — each reading its value off
    that offset's row in every employee's block, instead of summing across
    the block (which would silently add unrelated quantities together).

    Returns `(grid, new_layout)`; callers must use the returned grid, not
    the original worksheet, together with the returned layout."""
    max_col = ws.max_column
    block_height = layout.block_height
    # Only the last `item_offset_rows` header rows are offset-specific for
    # ITEM columns — any earlier ones are super-group labels not tied to a
    # single offset (see the `needs_flatten` assignment above), and
    # skipping them is safe because the offset rows already carry
    # specific-enough item labels on their own. Normally this is the same
    # count as `block_height` (every physical block row has its own header
    # row) — `item_offset_rows` only differs when the identity columns
    # need one MORE leading header row than the item columns do (confirmed
    # real case: 업체058, see `SourceLayout.item_offset_rows`).
    item_offset_rows = layout.item_offset_rows if layout.item_offset_rows is not None else block_height
    header_rows = layout.header_rows[-item_offset_rows:]

    # (orig_col, offset, label), in reading order: left-to-right through the
    # columns, then top-to-bottom through that column's offsets. Deliberately
    # reads the RAW cell value here, not the merge-aware `cell_text` used
    # everywhere else in this module — a column that's genuinely ONE item
    # for the whole block (연번/성명 here) has its 3 header cells merged
    # into one, so only the merge's anchor cell actually holds a value and
    # the other 2 offsets correctly read blank and get skipped; reading
    # `cell_text` here would have "inherited" that same label onto all 3
    # offsets and produced 2 phantom duplicate columns (confirmed real
    # case: 업체016's 연번/성명 columns are merged this way, while
    # its genuinely-3-different-items columns like 기본급/주휴수당/사대보험
    # 환급 are 3 separate, unmerged cells and come through untouched).
    # When `identity_offset != 0`, the identity columns are NOT one merge
    # spanning the whole header block (confirmed real case:
    # 업체058 has 3 plain, unmerged cells per identity column —
    # 직종/직원번호/지급일자 in column A, 급여유형/성명/인력구분 in column E),
    # so unlike every previously confirmed case, a raw per-offset scan WOULD
    # pick them up here too — each as its own, wrongly-offset, wrongly-
    # labeled logical column (e.g. column A's row-5 text "직원번호" paired
    # with row 9's job-title value). Left in, that also blocks the dedicated
    # identity rescan below via its "already covered" check. Skip these
    # columns here entirely; the rescan below is solely responsible for them
    # in that case.
    identity_cols = (
        {c for c in (layout.marker_col, layout.name_col, layout.birth_col) if c is not None}
        if layout.identity_offset != 0
        else set()
    )
    logical_cols: list[tuple[int, int, str]] = []
    for c in range(1, max_col + 1):
        if c in identity_cols:
            continue
        found = False
        for o in range(len(header_rows)):
            raw = ws.cell(row=header_rows[o], column=c).value
            label = raw.strip() if isinstance(raw, str) else ""
            if label:
                logical_cols.append((c, o, label))
                found = True
        if not found and c not in (layout.marker_col, layout.name_col, layout.birth_col):
            # 2026-10-07 (업체034/업체035 양식): "기 본 급 ①"/"실수령액"처럼 헤더
            # 블록 전체를 세로로 병합한 열은 오프셋 행에 원시 값이 없어서
            # 통째로 빠지고 있었다 — 병합이 오프셋 첫 행까지 내려오면 그
            # 블록의 첫 줄 값으로 본다.
            merged = _vertical_merge_label(ws, header_rows[0], c)
            if merged:
                logical_cols.append((c, 0, merged))

    # 연번/이름(/생년월일) describe the whole block, not one offset within
    # it, so their header cell is typically merged across the FULL header
    # block rather than repeated per offset — and the per-offset scan above
    # only samples the last `block_height` header rows, which can exclude
    # the row that merge is anchored on (confirmed real case:
    # 업체061: header_rows=[5,6,7] but block_height=2, so
    # only rows 6-7 are sampled above; 연번/직원명 are merged over 5-7 with
    # their text anchored on row 5, so the scan above never saw them and
    # `new_marker_col`/`new_name_col` below silently fell back to whatever
    # physical column 1/2 happened to be — an unrelated amount column — so
    # every employee's name read back as a number and nobody matched).
    # Re-check these two (three) specific columns across the FULL original
    # header block with merge-aware `cell_text`, and anchor their data read
    # at block offset 0 — the data-area merge for these columns (e.g.
    # A8:A9) is likewise anchored at the block's first physical row. This
    # is the ordinary case (`layout.identity_offset == 0`, true for every
    # previously confirmed needs_flatten file); see the `identity_offset`
    # branch below for the one confirmed exception.
    identity_logical: dict[int, int] = {}
    if layout.identity_offset == 0:
        for orig_col in (layout.marker_col, layout.name_col, layout.birth_col):
            if orig_col is None:
                continue
            if orig_col == layout.name_col:
                expected_labels = _SRC_NAME_LABELS
            elif orig_col == layout.birth_col:
                expected_labels = _BIRTH_LABELS
            else:
                expected_labels = _SRC_MARKER_LABELS | _CODE_LABELS
            existing = [
                i for i, (c, o, lab) in enumerate(logical_cols)
                if c == orig_col and o == 0 and norm(lab) in expected_labels
            ]
            if existing:
                identity_logical[orig_col] = existing[0] + 1
                continue
            if any(c == orig_col for c, _, _ in logical_cols) and orig_col == layout.marker_col and orig_col != layout.name_col:
                continue
            label = None
            for r in layout.header_rows:
                text = cell_text(ws, r, orig_col)
                if isinstance(text, str) and text.strip() and norm(text) in expected_labels:
                    label = text.strip()
                    break
            if label is None:
                for r in layout.header_rows:
                    text = cell_text(ws, r, orig_col)
                    if isinstance(text, str) and text.strip():
                        label = text.strip()
                        break
            logical_cols.append((orig_col, 0, label or ""))
            identity_logical[orig_col] = len(logical_cols)
    else:
        # Confirmed real case: 업체058's 성명/직원번호 sit one
        # header row deeper than item columns' own offset 0 (see
        # `SourceLayout.identity_offset`), so "first non-blank row wins"
        # would grab the WRONG sibling label (직종/급여유형, one row above)
        # instead. Require the row's text to actually name the field we're
        # looking for (성명/이름/... for name_col, 직원코드류 for marker_col,
        # 생년월일류 for birth_col) and read that row's own position in
        # `layout.header_rows` as the data offset — a genuinely different
        # sibling label elsewhere in the block can't be mistaken for the
        # identity field itself this way.
        for orig_col in (layout.marker_col, layout.name_col, layout.birth_col):
            if orig_col is None or any(c == orig_col for c, _, _ in logical_cols):
                continue
            if orig_col == layout.marker_col:
                expected_labels = _SRC_MARKER_LABELS | _CODE_LABELS
            elif orig_col == layout.name_col:
                expected_labels = _SRC_NAME_LABELS
            else:
                expected_labels = _BIRTH_LABELS
            label = offset = None
            for idx, r in enumerate(layout.header_rows):
                text = cell_text(ws, r, orig_col)
                if isinstance(text, str) and text.strip() and norm(text.strip()) in expected_labels:
                    label, offset = text.strip(), idx
                    break
            if label:
                logical_cols.append((orig_col, offset, label))

    # Find each employee block's starting row — more reliable than a "합계"
    # text scan or last-non-blank-row, since trailing summary/footer rows
    # (column totals, a grand total) sit right after the last real block
    # with no marker value of their own and would otherwise count as a
    # fake extra employee (confirmed real case: 업체016's sheet
    # has 3 such footer rows immediately after employee 40's block).
    #
    # A multi-page ledger can also interrupt the sequence mid-way with a
    # per-page subtotal ("합계") row followed by a *repeated* header block
    # before the numbering resumes (confirmed real case: 업체057
    # 급여대장 — employee 11's block ends, then 3 rows of "합계", a blank
    # row, and a fresh 4-row header block, before employee 12 continues).
    # Stopping at the first non-matching row (as this used to) silently
    # dropped every employee past the first page. `_MAX_GAP_LOOKAHEAD`
    # tolerates that kind of page-break furniture by scanning a bounded
    # distance ahead for the next expected marker value instead of giving
    # up immediately — but only jumps forward when the number itself is
    # found again, so it can't accidentally invent an employee out of
    # unrelated cells.
    _MAX_GAP_LOOKAHEAD = 20
    block_start_rows: list[int] = []
    if layout.has_marker:
        # Walk the marker column's own 1,2,3,... sequence.
        expected = 1
        r = layout.data_start_row
        while r <= ws.max_row:
            if _as_positive_int(ws.cell(row=r, column=layout.marker_col).value) == expected:
                block_start_rows.append(r)
                expected += 1
                r += block_height
                continue
            resumed_at = None
            for lookahead in range(1, _MAX_GAP_LOOKAHEAD + 1):
                probe_row = r + lookahead
                if probe_row > ws.max_row:
                    break
                if _as_positive_int(ws.cell(row=probe_row, column=layout.marker_col).value) == expected:
                    resumed_at = probe_row
                    break
            if resumed_at is None:
                break
            r = resumed_at
    elif layout.identity_offset == 0:
        # No marker column exists at all (confirmed real case:
        # 급여명세대장(2026년7월).xlsx) — walk the name column's own values
        # instead, one plausible employee name per block start.
        block_count = 0
        while True:
            r = layout.data_start_row + block_count * block_height
            if r > ws.max_row:
                break
            if not _is_plausible_employee_name(ws.cell(row=r, column=layout.name_col).value):
                break
            block_count += 1
        block_start_rows = [layout.data_start_row + i * block_height for i in range(block_count)]
    else:
        # Same idea, but the name column's real value sits `identity_offset`
        # rows into the block rather than on its very first row (confirmed
        # real case: 업체058 — see `SourceLayout.identity_offset`).
        # This file is also multi-page, with page-number/signature furniture
        # between pages (confirmed real case: the same file's 1/3, 2/3, 3/3
        # footers) — plain sequential stepping would stop dead at the first
        # page break, so this reuses the same bounded-lookahead tolerance the
        # `has_marker` branch above already relies on for that.
        r = layout.data_start_row
        while r <= ws.max_row:
            if _is_plausible_employee_name_offset(
                ws.cell(row=r + layout.identity_offset, column=layout.name_col).value
            ):
                block_start_rows.append(r)
                r += block_height
                continue
            resumed_at = None
            for lookahead in range(1, _MAX_GAP_LOOKAHEAD + 1):
                probe_row = r + lookahead
                if probe_row + layout.identity_offset > ws.max_row:
                    break
                if _is_plausible_employee_name_offset(
                    ws.cell(row=probe_row + layout.identity_offset, column=layout.name_col).value
                ):
                    resumed_at = probe_row
                    break
            if resumed_at is None:
                break
            r = resumed_at

    # 오프셋 행 위의 상위 그룹 헤더 행(공제내역/사업장 부담금 공제내역 등)도
    # 펼친 표에 그대로 남긴다 — 이게 없으면 "사업장 부담금 > 국민연금"처럼
    # 회사 부담분 열이 직원 공제로 한 번 더 잡힌다(2026-10-07 업체034/업체035 양식).
    group_rows = layout.header_rows[: max(0, len(layout.header_rows) - item_offset_rows)]
    rows: list[list[object]] = []
    for g in group_rows:
        row_vals = []
        for c, _o, _label in logical_cols:
            text = cell_text(ws, g, c)
            row_vals.append(text.strip() if isinstance(text, str) and text.strip() else None)
        rows.append(row_vals)
    rows.append([label for _, _, label in logical_cols])
    for block_start in block_start_rows:
        rows.append([ws.cell(row=block_start + o, column=c).value for c, o, _ in logical_cols])
    grid = Grid(rows)
    header_height = len(group_rows) + 1

    new_marker_col = identity_logical.get(layout.marker_col)
    new_name_col = identity_logical.get(layout.name_col)
    for idx, (c, _, label) in enumerate(logical_cols, start=1):
        if c == layout.marker_col and norm(label) in (_SRC_MARKER_LABELS | _CODE_LABELS) and new_marker_col is None:
            new_marker_col = idx
        if c == layout.name_col and norm(label) in _SRC_NAME_LABELS and new_name_col is None:
            new_name_col = idx
    if new_marker_col is None:
        new_marker_col = next((i for i, (c, _o, _l) in enumerate(logical_cols, start=1) if c == layout.marker_col), None)

    # Prefer the logical column actually labeled 생년월일: when name and birth
    # share one physical column on different block rows (업체009 전주:
    # 이름/직종/생년월일 stacked in B), identity_logical maps birth_col to the
    # NAME column and same-name employees (사원012/68) could not be told apart.
    new_birth_col = None
    if layout.birth_col is not None:
        new_birth_col = next((idx for idx, (c, _, label) in enumerate(logical_cols, start=1)
                              if c == layout.birth_col and norm(label) in _BIRTH_LABELS),
                             identity_logical.get(layout.birth_col))

    new_layout = SourceLayout(
        header_rows=list(range(1, header_height + 1)),
        data_start_row=header_height + 1,
        name_col=new_name_col if new_name_col is not None else layout.name_col,
        marker_col=new_marker_col if new_marker_col is not None else 1,
        birth_col=new_birth_col,
        mode="row",
        block_height=1,
        has_marker=True,
        needs_flatten=False,
    )
    return grid, new_layout


def _detect_value_sequence_marker(
    ws, start_row: int, end_row: int, candidate_cols
) -> tuple[int, int, int] | None:
    """Finds an unlabeled 순번 column by its VALUES alone — a strict
    1,2,3,... sequence at a perfectly consistent row gap (the gap is the
    employee block height; >1 means continuation rows, see
    SourceLayout.block_height). Stricter than
    core/heuristic_match.py's old `_looks_like_sequence` (that tolerated a
    ragged 1-5 row gap) since this has no header-text confirmation to lean
    on — a false positive here silently mis-splits every employee block."""
    best: tuple[int, int, int, int] | None = None  # (row_count, col, first_row, gap)
    for col in candidate_cols:
        hits = []
        for r in range(start_row, end_row + 1):
            n = _as_positive_int(ws.cell(row=r, column=col).value)
            if n is not None:
                hits.append((r, n))
        if len(hits) < 3 or hits[0][1] > 3:
            continue
        # 2026-10-07 (업체035): 순번 1..N 아래쪽(명세서/참고표 영역)에 다른 숫자가
        # 더 있으면 전체가 "순번 아님"으로 버려졌다 — 처음부터 끊김 없이
        # 1씩 늘고 간격이 일정한 구간까지만 순번으로 본다.
        run = [hits[0]]
        for prev, cur in zip(hits, hits[1:]):
            if cur[1] != prev[1] + 1:
                break
            if len(run) >= 2 and cur[0] - prev[0] != run[1][0] - run[0][0]:
                break
            run.append(cur)
        if len(run) < 3:
            continue
        hits = run
        gaps = [hits[i + 1][0] - hits[i][0] for i in range(len(hits) - 1)]
        candidate = (len(hits), col, hits[0][0], gaps[0])
        if best is None or candidate[0] > best[0]:
            best = candidate
    if best is None:
        return None
    _, col, first_row, gap = best
    return col, first_row, gap


def detect_source_layout_columnar(ws) -> SourceLayoutColumnar | None:
    """Best-effort detector for a 열 기반(직원=열, 항목=행) source sheet — no
    confirmed real Excel sample exists in 파일/ yet (unlike the HWP case in
    the sister project's CLAUDE.md), so this is unvalidated against real
    data. Looks for a label column (usually column A) whose rows spell out
    item names (성명/기본급/...), with employee blocks running across the
    columns to its right."""
    max_col = ws.max_column
    max_row = ws.max_row
    for label_col in range(1, min(max_col, 3) + 1):
        name_row = None
        for r in range(1, min(max_row, 30) + 1):
            if norm(cell_text(ws, r, label_col)) in _SRC_NAME_LABELS:
                name_row = r
                break
        if name_row is None:
            continue

        # Collect item-label rows below the name row until labels stop
        # canonicalizing to anything payroll-shaped (blank run or a totals
        # marker ends the block).
        item_rows: list[int] = []
        blank_streak = 0
        for r in range(name_row + 1, min(max_row, name_row + 40) + 1):
            label = str(cell_text(ws, r, label_col) or "").strip()
            if not label:
                blank_streak += 1
                if blank_streak >= 3:
                    break
                continue
            blank_streak = 0
            if norm(label) == "합계":
                break
            item_rows.append(r)
        if not item_rows:
            continue

        # Block width: smallest gap between consecutive non-empty cells on
        # the name row, to the right of the label column.
        name_cols = [c for c in range(label_col + 1, max_col + 1) if cell_text(ws, name_row, c)]
        if len(name_cols) < 1:
            continue
        gaps = [name_cols[i + 1] - name_cols[i] for i in range(len(name_cols) - 1)]
        block_width = min(gaps) if gaps else 1
        block_start_col = name_cols[0]
        block_count = len(name_cols)

        return SourceLayoutColumnar(
            label_col=label_col,
            name_row=name_row,
            item_rows=item_rows,
            block_start_col=block_start_col,
            block_width=max(block_width, 1),
            block_count=block_count,
        )
    return None


# --- header text extraction, shared by core/classify.py -------------------


def source_full_header(ws, layout: SourceLayout, col: int) -> str:
    """Port of VBA `SrcFullHeader` — every header row's own label for this
    column, top-to-bottom, joined with " > "; generalized from a fixed
    2-row span to however many header rows were actually detected.

    A cell merged vertically across several header rows (a single-row item
    with no separate group label above it) reads back the same text for
    each of those rows, which without deduping produces a noisy, needlessly
    long "기본급 > 기본급 > 기본급" instead of just "기본급" — confirmed
    real case (업체013), and the padding was wide enough to crowd
    out column space in 항목매핑결과's table, truncating other rows'
    genuinely-different text."""
    parts: list[str] = []
    for r in layout.header_rows:
        raw = cell_text(ws, r, col)
        if isinstance(raw, str) and raw.strip():
            text = raw.strip()
            if not parts or parts[-1] != text:
                parts.append(text)
    return " > ".join(parts)


_SRC_ITEM_SKIP = {
    "급여", "지급", "수당", "공제", "월급여", "근무내용", "공제내역", "기관부담금", "서비스별직원급여",
}


_SRC_ITEM_QUALIFIERS = {"고정", "변동", "포함", "미포함", "수기입력", "자동", "자동계산"}


def _parent_label(ws, layout: SourceLayout, col: int, below_row: int) -> str:
    for r in reversed([h for h in layout.header_rows if h < below_row]):
        raw = cell_text(ws, r, col)
        if isinstance(raw, str) and raw.strip() and norm(raw) not in _SRC_ITEM_SKIP:
            return raw.strip()
    return ""


def source_item(ws, layout: SourceLayout, col: int) -> str:
    """Port of VBA `SrcItem` — the most specific (bottom-most non-blank)
    header label for this column, skipping pure group-header words. VBA
    scanned bottom-up over a fixed 2-row span; this walks the detected
    header rows bottom-up instead, which also correctly resolves columns
    where a later header row's label shadows an earlier group label sharing
    the same column (confirmed real case: 업체062's column 15
    has "기타수당" on one header row and "연차수당" directly below it in the
    next — the bottom one, 연차수당, is what the macro (and Wehago) actually
    means for that column). Header cells that hold a raw number (e.g. a
    4대보험 rate row like 0.0475 sitting between the leaf-label row and the
    first data row — confirmed real case: 업체063's 원본 sheet)
    are skipped instead of stringified — a number is never a real item
    label, and letting one through here would shadow the actual label
    above it."""
    fallback = ""
    for r in reversed(layout.header_rows):
        raw = cell_text(ws, r, col)
        if not isinstance(raw, str):
            continue
        text = raw.strip()
        key = norm(text)
        if not text:
            continue
        if "인건비신고" in key or "사업자부담" in key:
            continue
        if key in _SRC_ITEM_QUALIFIERS and not fallback:
            # "기본급 > 고정", "연차수당 > 포함" — 맨 아래 칸이 항목명이 아니라
            # 수식어다(2026-10-07 업체064/업체065/업체066).
            # 바로 위 칸의 실제 항목명을 쓴다.
            parent = _parent_label(ws, layout, col, r)
            if parent:
                return parent
        if key in _SRC_ITEM_SKIP:
            # A generic group word ("급여") only gets skipped in favor of a
            # more specific label lower in the header block — if it's the
            # ONLY label this column ever has (confirmed real case:
            # 업체056's "급여" column has nothing more specific
            # anywhere), a generic name beats losing the column entirely.
            if not fallback:
                fallback = text
            continue
        return text
    return fallback


def dst_full_header(ws, layout: DestLayout, col: int) -> str:
    rows = [layout.header_row1] + ([layout.header_row2] if layout.header_row2 else [])
    parts = []
    for r in rows:
        text = str(cell_text(ws, r, col) or "").strip()
        if text:
            parts.append(text)
    return " > ".join(parts)


def dst_group_label(ws, layout: DestLayout, col: int) -> str:
    return str(cell_text(ws, layout.header_row1, col) or "").strip()


def dst_item(ws, layout: DestLayout, col: int) -> str:
    if layout.header_row2:
        leaf = str(cell_text(ws, layout.header_row2, col) or "").strip()
        if leaf:
            return leaf
    return str(cell_text(ws, layout.header_row1, col) or "").strip()
