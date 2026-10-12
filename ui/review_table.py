"""Editable review grid for WehagoRecord rows. Plain QTableWidget (not a
custom model) — record counts here are small (dozens to low hundreds), so
the simplicity is worth more than virtualized-model performance.

Column set is built per-profile (core/target_schema.py) but with any
allowance/deduction column that is 0 for every employee this run dropped —
a client's Wehago profile registers a fixed item set, but not every item
has an amount every month. For any total field the profile tracks a source
value (core/reconcile.py's `_apply_source_totals_check`), a read-only "원본 X계" column
is inserted directly to the LEFT of our computed "X계" column so the two
sit side by side — the source ledger's own totals are not always correct,
so we always recompute from raw line items ourselves, but this makes it
obvious exactly where and by how much they disagree (only those two cells
turn red, not the whole row).

On-screen the header is 2 rows tall with merged group cells (수당/공제/...
spanning their columns) — this is purely a display choice and has nothing
to do with the exported file, which core/export.py always writes as a flat,
single-row, unmerged header (Wehago's bulk-upload parser requires that; the
two are independent)."""

from __future__ import annotations

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QBrush, QColor, QFont
from PySide6.QtWidgets import QAbstractItemView, QTableWidget, QTableWidgetItem

from core.profile import ClientProfile
from core.target_schema import IDENTITY_COLUMNS, drop_all_zero_columns
from core.totals import compute_taxable_split, compute_totals
from core.wehago_schema import WehagoRecord

_NON_NUMERIC_KEYS = {"사원코드", "사원명", "부서", "직급", "직종"}
_COMPUTED_KEYS = {"비과세합계", "과세합계", "지급액계", "공제액계", "차인지급액"}
_AMOUNT_GROUPS = {"수당", "공제", "합계"}
_SOURCE_PREFIX = "__source_"
_MISMATCH_COLOR = QBrush(QColor(255, 214, 214))
_TOTALS_ROW_COLOR = QBrush(QColor(230, 230, 230))
_GROUP_HEADER_COLOR = QBrush(QColor(224, 224, 224))
# Header highlight for a column whose auto-match relied on an
# is_ambiguous_alias label (core/canon.py) — same yellow as
# ui/verification_dialog.py's 항목매핑결과 flag, so a column worth double
# checking is visible here too, without switching to the 항목매핑결과 tab
# (사용자 요청, 2026-08-07).
_AMBIGUOUS_HEADER_COLOR = QBrush(QColor(255, 244, 179))
_GRAND_TOTAL_LABEL = "합계"

# Compact sizing so more columns fit on screen without horizontal scrolling.
_CELL_FONT_SIZE = 9
_ROW_HEIGHT = 22
_GROUP_ROW_HEIGHT = 20
_DEFAULT_COL_WIDTH = 64

# Row 0 is the merged group-label row (인적사항/수당/공제/합계), row 1 is the
# leaf column label (사원코드/기본급/국민연금/원본공제액계/...); actual
# employee data starts at row 2.
_GROUP_ROW = 0
_LABEL_ROW = 1
_DATA_START_ROW = 2


def _source_field(key: str) -> str | None:
    return key[len(_SOURCE_PREFIX):] if key.startswith(_SOURCE_PREFIX) else None


def _column_specs(profile: ClientProfile, records: list[WehagoRecord]) -> list[tuple[str, str, str, bool]]:
    """(key, label, group, editable) in physical column order.

    A read-only "원본 X" column is inserted directly before ANY column where
    at least one employee this run has an entry in `record.mismatched_fields`
    — currently that's only the three grand totals, populated by
    core/mapping.py's `_check_source_totals` when they disagree with what the
    source file itself independently states (normal rounding of individual
    line items is NOT flagged — 2266.875 -> 2267 is expected, correct
    behavior, not a discrepancy). Written generically so any future
    per-line-item check would show up here automatically too.
    """
    allowance_columns = drop_all_zero_columns(profile.allowance_columns, records, "allowance")
    deduction_columns = drop_all_zero_columns(profile.deduction_columns, records, "deduction")
    fields_with_source_diff = {f for r in records for f in r.mismatched_fields}

    def with_source(key: str, label: str, group: str, editable: bool) -> list[tuple[str, str, str, bool]]:
        cols = []
        if key in fields_with_source_diff:
            cols.append((f"{_SOURCE_PREFIX}{key}", f"원본{label}", group, False))
        cols.append((key, label, group, editable))
        return cols

    specs: list[tuple[str, str, str, bool]] = [(c.key, c.label, "인적사항", True) for c in IDENTITY_COLUMNS]
    for c in allowance_columns:
        specs += with_source(c.key, c.label, "수당", True)
    specs += with_source("비과세합계", "비과세합계", "합계", False)
    specs += with_source("과세합계", "과세합계", "합계", False)
    specs += with_source("지급액계", "지급액계", "합계", False)
    for c in deduction_columns:
        specs += with_source(c.key, c.label, "공제", True)
    specs += with_source("공제액계", "공제액계", "합계", False)
    specs += with_source("차인지급액", "차인지급액", "합계", False)
    return specs


def _format_cell(group: str, value) -> str:
    if value is None:
        return ""
    if group in _AMOUNT_GROUPS and isinstance(value, (int, float)):
        # "원본 X" cells can carry the unrounded source value (that's the
        # whole point of showing them next to our rounded value) — flooring
        # to 0 decimals here would hide exactly the fraction the accountant
        # needs to see, so only drop the decimal point when there isn't one.
        return f"{value:,.0f}" if float(value).is_integer() else f"{value:,.3f}".rstrip("0").rstrip(".")
    return str(value)


class ReviewTable(QTableWidget):
    # Emitted whenever the selection changes, e.g. "합계 1,234,567 · 평균
    # 12,345 · 셀 100개" — like Excel's status-bar selection summary. Empty
    # string when nothing numeric is selected. A parent widget (main_window)
    # displays this in a status label; the table itself has no room for one.
    selection_summary_changed = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._records: list[WehagoRecord] = []
        self._column_specs: list[tuple[str, str, str, bool]] = []
        self._non_taxable_keys: set[str] = set()
        self._ambiguous_keys: set[str] = set()
        self._totals_row_idx = -1
        self._loading = False
        self.itemChanged.connect(self._on_item_changed)
        self.itemSelectionChanged.connect(self._on_selection_changed)

        cell_font = QFont()
        cell_font.setPointSize(_CELL_FONT_SIZE)
        self.setFont(cell_font)
        self.horizontalHeader().setVisible(False)  # rows 0-1 ARE the header, visually
        self.horizontalHeader().setDefaultSectionSize(_DEFAULT_COL_WIDTH)
        self.horizontalHeader().setMinimumSectionSize(40)
        self.verticalHeader().setDefaultSectionSize(_ROW_HEIGHT)
        self.verticalHeader().setVisible(False)
        self.setEditTriggers(
            QAbstractItemView.EditTrigger.DoubleClicked | QAbstractItemView.EditTrigger.EditKeyPressed
        )
        self.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)

    def load_records(
        self, records: list[WehagoRecord], profile: ClientProfile, ambiguous_keys: set[str] | None = None
    ) -> None:
        self._loading = True
        self.clearSpans()
        self._column_specs = _column_specs(profile, records)
        self._non_taxable_keys = {c.key for c in profile.allowance_columns if not c.taxable}
        self._ambiguous_keys = ambiguous_keys or set()
        self._records = records

        self.setColumnCount(len(self._column_specs))
        # +3 rows: merged group-label row, leaf-label row, grand-total row.
        self.setRowCount(len(records) + 3)

        self._build_header_rows()
        self.setRowHeight(_GROUP_ROW, _GROUP_ROW_HEIGHT)
        self.setRowHeight(_LABEL_ROW, _GROUP_ROW_HEIGHT)

        bold = QFont()
        bold.setPointSize(_CELL_FONT_SIZE)
        bold.setBold(True)

        for i, record in enumerate(records):
            row_idx = _DATA_START_ROW + i
            for col_idx, (key, _, group, editable) in enumerate(self._column_specs):
                source_field = _source_field(key)
                raw_value = record.source_totals.get(source_field) if source_field else record.column_value(key)
                item = QTableWidgetItem(_format_cell(group, raw_value))
                if not editable:
                    item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                # Only the exact total cell(s) that disagree with the source
                # — and its adjacent "원본 X계" cell — turn red, not the
                # whole row, so it's obvious at a glance which total is off
                # and by how much (both cells are right next to each other).
                mismatch_field = source_field or key
                if mismatch_field in record.mismatched_fields:
                    item.setBackground(_MISMATCH_COLOR)
                    stated = record.source_totals.get(mismatch_field)
                    computed = record.column_value(mismatch_field)
                    if stated is not None:
                        diff = computed - stated
                        item.setToolTip(
                            f"{mismatch_field}: 원본 {stated:,.0f}원 → 파싱 {computed:,.0f}원 "
                            f"(차이 {diff:+,.0f}원)"
                        )
                self.setItem(row_idx, col_idx, item)

        self._totals_row_idx = _DATA_START_ROW + len(records)
        for col_idx, (key, _, group, _editable) in enumerate(self._column_specs):
            item = QTableWidgetItem()
            item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            item.setBackground(_TOTALS_ROW_COLOR)
            item.setFont(bold)
            self.setItem(self._totals_row_idx, col_idx, item)
        self._refresh_totals_row()

        # Size each column to its actual content (header labels + widest
        # value in that column) instead of a one-size-fits-all default —
        # short columns like 고용보험 no longer eat as much width as columns
        # with longer labels/values.
        self.resizeColumnsToContents()
        for col_idx in range(len(self._column_specs)):
            self.setColumnWidth(col_idx, min(self.columnWidth(col_idx) + 6, 220))

        self._loading = False

    def _build_header_rows(self) -> None:
        bold = QFont()
        bold.setPointSize(_CELL_FONT_SIZE)
        bold.setBold(True)

        # Row 0: group name, merged across every column that shares it
        # (e.g. "수당" spans 기본급..활동비) — same idea as the merged group
        # header the exported file used to have, before that was flattened
        # for Wehago's upload parser; this one is purely a screen display.
        col = 0
        n = len(self._column_specs)
        while col < n:
            group = self._column_specs[col][2]
            span_end = col
            while span_end + 1 < n and self._column_specs[span_end + 1][2] == group:
                span_end += 1
            span = span_end - col + 1

            item = QTableWidgetItem(group)
            item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            item.setBackground(_GROUP_HEADER_COLOR)
            item.setFont(bold)
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.setItem(_GROUP_ROW, col, item)
            if span > 1:
                self.setSpan(_GROUP_ROW, col, 1, span)

            col = span_end + 1

        # Row 1: the actual per-column name (사원코드, 기본급, 원본공제액계, ...).
        for col_idx, (key, label, _group, _editable) in enumerate(self._column_specs):
            item = QTableWidgetItem(label)
            item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            if key in self._ambiguous_keys:
                item.setBackground(_AMBIGUOUS_HEADER_COLOR)
                item.setToolTip(
                    f"\"{label}\"은(는) 사업장마다 다른 뜻일 수 있는 항목명으로 자동 매칭됐습니다 — "
                    "항목매핑결과 탭에서 실제 매칭된 원본 항목을 확인하세요."
                )
            else:
                item.setBackground(_GROUP_HEADER_COLOR)
            item.setFont(bold)
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.setItem(_LABEL_ROW, col_idx, item)

    def records(self) -> list[WehagoRecord]:
        return self._records

    def _on_item_changed(self, item: QTableWidgetItem) -> None:
        if self._loading:
            return
        row_idx = item.row()
        if row_idx in (_GROUP_ROW, _LABEL_ROW):
            return
        col_idx = item.column()
        key, _, _, _ = self._column_specs[col_idx]
        record = self._records[row_idx - _DATA_START_ROW]

        if _source_field(key) is not None:
            return  # "원본 X계" columns are informational only, never editable
        if key in _NON_NUMERIC_KEYS:
            setattr(record, key, item.text())
            return
        if key in _COMPUTED_KEYS:
            return  # not directly editable; recomputed below

        try:
            value = float(item.text().replace(",", ""))
        except ValueError:
            value = 0.0

        if key in record.allowances:
            record.allowances[key] = value
        elif key in record.deductions:
            record.deductions[key] = value

        # A manual edit is the accountant deliberately overriding the value,
        # so the "this differs from the source" flag on this specific cell
        # no longer applies (the adjacent "원본 X" column, if any, stays put
        # for other rows — only this row's own discrepancy is resolved).
        record.mismatched_fields.discard(key)

        self._loading = True
        item.setText(f"{value:,.0f}")
        item.setBackground(QBrush())
        self._loading = False

        self._recompute_row(row_idx)
        self._refresh_mismatch_highlight(row_idx)
        self._refresh_totals_row()

    def _recompute_row(self, row_idx: int) -> None:
        record = self._records[row_idx - _DATA_START_ROW]
        record.지급액계, record.공제액계, record.차인지급액 = compute_totals(
            record.allowances, record.deductions
        )
        record.비과세합계, record.과세합계 = compute_taxable_split(record.allowances, self._non_taxable_keys)
        self._loading = True
        for key in ("비과세합계", "과세합계", "지급액계", "공제액계", "차인지급액"):
            col_idx = next((i for i, spec in enumerate(self._column_specs) if spec[0] == key), None)
            if col_idx is not None:
                self.item(row_idx, col_idx).setText(f"{record.column_value(key):,.0f}")
        self._loading = False

    def _refresh_mismatch_highlight(self, row_idx: int) -> None:
        """A manual edit to a line item (수당/공제 cell) changes 지급액계/
        공제액계/차인지급액 too, but those totals' red "differs from source"
        highlight was only ever set once, at `load_records` time — without
        this, editing a value so the recomputed total now matches the
        source's own stated total (e.g. correcting 두루누리지원액 from
        +60,000 to -60,000 so 차인지급액 lines up with 원본 실지급액) left
        the cell red forever, even though the discrepancy is now resolved."""
        record = self._records[row_idx - _DATA_START_ROW]
        self._loading = True
        for field in ("지급액계", "공제액계", "차인지급액"):
            stated = record.source_totals.get(field)
            if stated is None:
                continue
            computed = record.column_value(field)
            is_mismatch = abs(computed - stated) > 1
            if is_mismatch:
                record.mismatched_fields.add(field)
            else:
                record.mismatched_fields.discard(field)
            for key in (f"{_SOURCE_PREFIX}{field}", field):
                col_idx = next((i for i, spec in enumerate(self._column_specs) if spec[0] == key), None)
                if col_idx is None:
                    continue
                cell = self.item(row_idx, col_idx)
                if is_mismatch:
                    diff = computed - stated
                    cell.setBackground(_MISMATCH_COLOR)
                    cell.setToolTip(
                        f"{field}: 원본 {stated:,.0f}원 → 파싱 {computed:,.0f}원 (차이 {diff:+,.0f}원)"
                    )
                else:
                    cell.setBackground(QBrush())
                    cell.setToolTip("")
        self._loading = False

    def _refresh_totals_row(self) -> None:
        if self._totals_row_idx < 0:
            return
        self._loading = True
        for col_idx, (key, _, group, _editable) in enumerate(self._column_specs):
            item = self.item(self._totals_row_idx, col_idx)
            source_field = _source_field(key)
            if key == "사원코드":
                item.setText(_GRAND_TOTAL_LABEL)
            elif key in _NON_NUMERIC_KEYS:
                item.setText("")
            elif source_field:
                total = sum(r.source_totals.get(source_field, 0) for r in self._records)
                item.setText(f"{total:,.0f}")
            elif group in _AMOUNT_GROUPS:
                total = sum(r.column_value(key) for r in self._records)
                item.setText(f"{total:,.0f}")
        self._loading = False

    def _on_selection_changed(self) -> None:
        """Like Excel's status-bar selection summary: sum/average/count of
        whatever numeric cells are currently selected, regardless of which
        columns or rows they span (shift-click or drag across a range)."""
        if self._loading:
            self.selection_summary_changed.emit("")
            return

        values: list[float] = []
        for item in self.selectedItems():
            text = item.text().strip().replace(",", "")
            if not text:
                continue
            try:
                values.append(float(text))
            except ValueError:
                continue

        if not values:
            self.selection_summary_changed.emit("")
            return

        total = sum(values)
        average = total / len(values)
        self.selection_summary_changed.emit(
            f"선택 {len(values)}개 셀 · 합계 {total:,.0f} · 평균 {average:,.1f}"
        )
