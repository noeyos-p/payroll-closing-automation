"""검증결과/항목매핑결과 — the two report tables the VBA macro
(파일/이지케어_Sheet1_현재시트_수정본_v3.bas) used to write into the workbook
itself, shown here as two tabs in the main window instead (this app always
exports just the 정리본 sheet — see CLAUDE.md). 검증결과 lists every
employee's 사업장 vs 위하고 지급/공제/차인지급액 and how they were resolved
(정상/차이/추가됨/삭제됨/확인필요); 항목매핑결과 lists which source column
auto-matched which Wehago column, and which ones didn't.

Originally a separate popup (`QDialog`) — moved to inline tabs in
`ui/main_window.py` (사용자 요청, 2026-08-12: 별도 창 대신 탭으로 전환)
so results stay visible alongside 검토표 without an extra window. This
module only builds the two tables; `ui/main_window.py` owns the tab
widget, the apply button, and the "선택한 매칭 적용" wiring.

항목매핑결과's "미매칭" rows are the one editable part: a header text
canon.py doesn't recognize (a new client's own wording for an existing
item, or something genuinely ambiguous like a 두루누리 지원금 whose sign
convention can't be guessed from the label alone — see core/reconcile.py's
`_NEVER_AUTO_MAP`) can be pointed at a destination column directly, instead
of needing a code change for every new synonym. Picking one and applying
re-runs the whole reconciliation with that mapping forced in
(core/reconcile.py's `manual_overrides`) — nothing is saved, same as
everything else in this app; it only applies to this run."""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtGui import QBrush, QColor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QHeaderView,
    QTableWidget,
    QTableWidgetItem,
)

from core.profile import ClientProfile
from core.reconcile import MappingRow, ReconcileResult, VerificationRow

_DIFF_STATUSES = {"차이", "추가됨_차이"}
_REMOVED_COLOR = QBrush(QColor(240, 240, 240))
_DIFF_COLOR = QBrush(QColor(255, 214, 214))
_ADDED_COLOR = QBrush(QColor(224, 240, 224))
# 자동매칭은 됐지만 canon.py상 "비슷하지만 다른 항목일 수 있는" 별칭으로
# 매칭된 경우(core/canon.py의 is_ambiguous_alias) — 틀렸다는 뜻은 아니고,
# 사람이 한 번 더 확인하는 게 안전하다는 신호(사용자 요청, 2026-08-07).
_AMBIGUOUS_COLOR = QBrush(QColor(255, 244, 179))
_NO_CHOICE = "(선택 안 함)"

_VERIFICATION_HEADERS = [
    "성명", "상태", "사업장 지급합계", "위하고 반영 지급합계", "지급 차이",
    "사업장 공제합계", "위하고 반영 공제합계", "공제 차이",
    "사업장 실수령액", "위하고 반영 차인지급액", "차인지급액 차이", "메모",
]
_MAPPING_HEADERS = ["구분", "사업장 급여대장 항목", "위하고 반영 열", "비고", "직접 매칭"]


def _num(value: float | None) -> str:
    return "" if value is None else f"{value:,.0f}"


def _read_only_table(headers: list[str], *, auto_fit_cols: range | None = None) -> QTableWidget:
    """`auto_fit_cols` columns get Qt's own continuous `ResizeToContents`
    mode instead of a one-time `resizeColumnsToContents()` snapshot — a
    snapshot taken before the tab is actually shown/polished can use the
    wrong (pre-stylesheet) font metrics and under-size CJK text, which is
    exactly what truncated "장기근속수당" down to "장기…" even after
    widening the column: the computed width itself was wrong, not just too
    small. `ResizeToContents` recomputes against the real, live font
    whenever the section needs it, so it can't go stale like that."""
    table = QTableWidget()
    table.setColumnCount(len(headers))
    table.setHorizontalHeaderLabels(headers)
    table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
    table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
    table.verticalHeader().setVisible(False)
    if auto_fit_cols is not None:
        header = table.horizontalHeader()
        for col in auto_fit_cols:
            header.setSectionResizeMode(col, QHeaderView.ResizeMode.ResizeToContents)
    return table


def _set_row(table: QTableWidget, row_idx: int, values: list[str], background: QBrush | None) -> None:
    for col_idx, text in enumerate(values):
        # Header text straight from the workbook often has a real newline
        # baked in (e.g. "국민연금\n(4.5%)", however that cell was wrapped in
        # Excel). A QTableWidgetItem never actually renders a second visual
        # line for it — the cell stays single-line — but Qt's own width
        # calculation for a "\n"-containing string sizes to the widest
        # *line* (as if it WERE wrapped), not the full run. Net effect: the
        # column got sized for the short second half ("(4.5%)") while
        # actually needing room for the whole thing on one line, so it
        # always looked truncated no matter how the column was resized.
        # Flattening to a space makes the measured width and the rendered
        # width agree.
        flat_text = text.replace("\n", " ") if text else text
        item = QTableWidgetItem(flat_text)
        item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
        if text:
            item.setToolTip(flat_text)  # full text on hover even if the column elides it
        if background is not None:
            item.setBackground(background)
        table.setItem(row_idx, col_idx, item)


def build_verification_table(rows: list[VerificationRow]) -> QTableWidget:
    table = _read_only_table(_VERIFICATION_HEADERS, auto_fit_cols=range(len(_VERIFICATION_HEADERS)))
    table.setRowCount(len(rows))
    for i, row in enumerate(rows):
        background = None
        if row.status == "삭제됨":
            background = _REMOVED_COLOR
        elif row.status in _DIFF_STATUSES:
            background = _DIFF_COLOR
        elif row.status.startswith("추가됨"):
            background = _ADDED_COLOR
        _set_row(
            table,
            i,
            [
                row.name,
                row.status,
                _num(row.src_pay),
                _num(row.dst_pay),
                _num(row.diff_pay),
                _num(row.src_ded),
                _num(row.dst_ded),
                _num(row.diff_ded),
                _num(row.src_net),
                _num(row.dst_net),
                _num(row.diff_net),
                row.memo,
            ],
            background,
        )
    table.resizeColumnsToContents()
    return table


def build_mapping_table(
    rows: list[MappingRow], profile: ClientProfile
) -> tuple[QTableWidget, dict[int, tuple[QComboBox, str | None]]]:
    """Every row (not just "미매칭") gets a picker, pre-selected to its
    current destination — including 자동매칭/수동매칭 rows, so a wrong
    auto-match can be redirected or explicitly forced to "매칭 안 함", not
    just a never-matched one filled in. The returned dict (source column ->
    (combo, original destination key)) is what the caller's apply button
    diffs against to build `manual_overrides`."""
    # Columns 0-3 are plain text (구분/사업장 급여대장 항목/위하고 반영
    # 열/비고) and always auto-fit to their own content; column 4 holds
    # the picker combobox and keeps a fixed, comfortably-usable width
    # instead (a combo's dropdown doesn't need to hug its current
    # selection's text length).
    table = _read_only_table(_MAPPING_HEADERS, auto_fit_cols=range(4))
    table.setRowCount(len(rows))
    combo_by_source_col: dict[int, tuple[QComboBox, str | None]] = {}
    for i, row in enumerate(rows):
        if row.kind == "미매칭":
            background = _DIFF_COLOR
        elif row.ambiguous:
            background = _AMBIGUOUS_COLOR
        else:
            background = None
        _set_row(table, i, [row.kind, row.source_header, row.dest_column_label, row.note, ""], background)

        combo = QComboBox()
        combo.addItem(_NO_CHOICE, None)
        for c in profile.allowance_columns:
            combo.addItem(f"[수당] {c.label}", c.key)
        for c in profile.deduction_columns:
            combo.addItem(f"[공제] {c.label}", c.key)
        if row.dest_key is not None:
            idx = combo.findData(row.dest_key)
            if idx >= 0:
                combo.setCurrentIndex(idx)
        table.setCellWidget(i, 4, combo)
        combo_by_source_col[row.source_col] = (combo, row.dest_key)
    table.setColumnWidth(4, 220)
    table.horizontalHeader().setStretchLastSection(False)
    return table, combo_by_source_col


def _grand_total_text(result: ReconcileResult) -> str:
    """Whole-run 지급액계/공제액계/차인지급액 vs the source ledger's own
    printed 합계 row (core/reconcile.py's `grand_total_*`) — unlike the
    per-employee checks in 검증결과, this catches money that went missing
    because an employee was dropped as 삭제됨 or left 확인필요, not just a
    wrong per-employee amount."""
    if not result.grand_total_source:
        return ""
    parts = []
    for field_name in ("지급액계", "공제액계", "차인지급액"):
        if field_name not in result.grand_total_source:
            continue
        stated = result.grand_total_source[field_name]
        computed = result.grand_total_computed[field_name]
        if field_name in result.grand_total_mismatches:
            diff = computed - stated
            parts.append(
                f'<span style="color:#c62828;font-weight:bold;">{field_name} 전체합계 불일치 '
                f"(원본 {stated:,.0f}원 → 결과 {computed:,.0f}원, 차이 {diff:+,.0f}원)</span>"
            )
        else:
            parts.append(f"{field_name} 전체합계 일치")
    return " · ".join(parts)


def summary_text(result: ReconcileResult) -> str:
    summary = result.summary
    base = (
        f"정상 매칭 {summary.matched}명 · 신규추가 {summary.added}명 · "
        f"삭제 {summary.removed}명 · 합계차이 {summary.diff}명 · 확인필요 {summary.need_check}명"
    )
    grand_total = _grand_total_text(result)
    return f"{base} · {grand_total}" if grand_total else base
