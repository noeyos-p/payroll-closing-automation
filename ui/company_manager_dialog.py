"""등록된 사업장 목록 — 이름/사원수/등록일 확인, 삭제, 위하고 참고파일
재등록(같은 이름으로 다시 등록 = upsert). 자동 저장(급여대장 처리 성공 시
조용히 등록/갱신)과는 별개로, 배치 처리를 시작하기 전에 미리 위하고 파일을
등록해두고 싶을 때를 위한 화면이다."""

from __future__ import annotations

from pathlib import Path

from PySide6.QtWidgets import (
    QDialog,
    QFileDialog,
    QHBoxLayout,
    QHeaderView,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
)

from core import registry


class CompanyManagerDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("사업장 관리")
        self.resize(640, 420)

        layout = QVBoxLayout(self)
        self._table = QTableWidget(0, 3)
        self._table.setHorizontalHeaderLabels(["사업장명", "사원수", "등록일"])
        self._table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self._table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self._table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self._table, stretch=1)

        btn_row = QHBoxLayout()
        register_btn = QPushButton("새 사업장 등록")
        register_btn.clicked.connect(self._on_register)
        btn_row.addWidget(register_btn)
        reregister_btn = QPushButton("선택 사업장 위하고 파일 재등록")
        reregister_btn.clicked.connect(self._on_reregister)
        btn_row.addWidget(reregister_btn)
        delete_btn = QPushButton("선택 사업장 삭제")
        delete_btn.clicked.connect(self._on_delete)
        btn_row.addWidget(delete_btn)
        btn_row.addStretch(1)
        close_btn = QPushButton("닫기")
        close_btn.clicked.connect(self.accept)
        btn_row.addWidget(close_btn)
        layout.addLayout(btn_row)

        self._reload()

    def _reload(self) -> None:
        companies = registry.list_companies()
        self._table.setRowCount(len(companies))
        for row, company in enumerate(companies):
            self._table.setItem(row, 0, QTableWidgetItem(company.name))
            count_item = QTableWidgetItem()
            count_item.setData(0, company.employee_count)
            self._table.setItem(row, 1, count_item)
            self._table.setItem(row, 2, QTableWidgetItem(company.registered_at))
            self._table.item(row, 0).setData(1000, company.id)

    def _selected_company_id(self) -> int | None:
        rows = self._table.selectionModel().selectedRows() if self._table.selectionModel() else []
        if not rows:
            return None
        item = self._table.item(rows[0].row(), 0)
        return item.data(1000) if item else None

    def _on_register(self) -> None:
        path_str, _ = QFileDialog.getOpenFileName(self, "위하고 참고파일 선택", "", "Excel Files (*.xlsx)")
        if not path_str:
            return
        from PySide6.QtWidgets import QInputDialog

        name, ok = QInputDialog.getText(self, "사업장 등록", "사업장 이름")
        if not ok or not name.strip():
            return
        try:
            registry.register_company(name.strip(), Path(path_str))
        except Exception as exc:
            QMessageBox.critical(self, "등록 오류", f"등록하지 못했습니다:\n{exc}")
            return
        self._reload()

    def _on_reregister(self) -> None:
        company_id = self._selected_company_id()
        if company_id is None:
            QMessageBox.information(self, "선택 필요", "재등록할 사업장을 먼저 선택하세요.")
            return
        company = registry.get_company(company_id)
        if company is None:
            return
        path_str, _ = QFileDialog.getOpenFileName(
            self, f"{company.name}의 새 위하고 참고파일 선택", "", "Excel Files (*.xlsx)"
        )
        if not path_str:
            return
        try:
            registry.register_company(company.name, Path(path_str))
        except Exception as exc:
            QMessageBox.critical(self, "재등록 오류", f"재등록하지 못했습니다:\n{exc}")
            return
        self._reload()

    def _on_delete(self) -> None:
        company_id = self._selected_company_id()
        if company_id is None:
            QMessageBox.information(self, "선택 필요", "삭제할 사업장을 먼저 선택하세요.")
            return
        company = registry.get_company(company_id)
        if company is None:
            return
        confirm = QMessageBox.question(
            self, "삭제 확인", f"'{company.name}' 등록을 삭제할까요?\n(위하고 참고파일 사본도 함께 지워지진 않습니다.)"
        )
        if confirm != QMessageBox.StandardButton.Yes:
            return
        registry.delete_company(company_id)
        self._reload()
