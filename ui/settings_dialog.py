"""환경설정 화면 — 위하고 양식 폴더 / 위하고 자동 업로드용 엑셀 저장 폴더
경로를 언제든 바꿀 수 있게 한다(둘 다 예전엔 코드에 하드코딩돼 있었음,
사용자 요청, 2026-09-21). 위하고 자동 로그인용 아이디/비밀번호도 여기서
저장한다 — 평문 저장(사용자 요청, 이 개인용 도구의 기존 방침과 동일,
`core/settings.py` 참고)."""

from __future__ import annotations

from PySide6.QtWidgets import (
    QDialog,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QVBoxLayout,
)

from core.settings import Settings, load_settings, save_settings


class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("환경설정")
        self.resize(560, 220)
        self._settings = load_settings()

        layout = QVBoxLayout(self)
        form = QFormLayout()

        self._wehago_form_edit = QLineEdit(self._settings.wehago_form_folder)
        form.addRow("위하고 양식 폴더", self._wrap_with_browse(self._wehago_form_edit))

        self._upload_export_edit = QLineEdit(self._settings.upload_export_folder)
        form.addRow("위하고 자동 업로드용 저장 폴더", self._wrap_with_browse(self._upload_export_edit))

        self._wehago_id_edit = QLineEdit(self._settings.wehago_id)
        form.addRow("위하고 아이디", self._wehago_id_edit)

        self._wehago_password_edit = QLineEdit(self._settings.wehago_password)
        self._wehago_password_edit.setEchoMode(QLineEdit.EchoMode.Password)
        form.addRow("위하고 비밀번호", self._wehago_password_edit)

        layout.addLayout(form)
        note = QLabel(
            "아이디/비밀번호를 입력해두면 위하고 자동 업로드 시 로그인 화면에서 "
            "자동으로 입력·로그인합니다. 비워두면 지금처럼 직접 로그인합니다."
        )
        note.setWordWrap(True)
        layout.addWidget(note)

        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        save_btn = QPushButton("저장")
        save_btn.clicked.connect(self._on_save)
        btn_row.addWidget(save_btn)
        cancel_btn = QPushButton("취소")
        cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(cancel_btn)
        layout.addLayout(btn_row)

    def _wrap_with_browse(self, edit: QLineEdit):
        from PySide6.QtWidgets import QWidget

        container = QWidget()
        row = QHBoxLayout(container)
        row.setContentsMargins(0, 0, 0, 0)
        row.addWidget(edit, 1)
        browse_btn = QPushButton("찾아보기")
        browse_btn.clicked.connect(lambda: self._on_browse(edit))
        row.addWidget(browse_btn)
        return container

    def _on_browse(self, edit: QLineEdit) -> None:
        folder = QFileDialog.getExistingDirectory(self, "폴더 선택", edit.text())
        if folder:
            edit.setText(folder)

    def _on_save(self) -> None:
        save_settings(
            Settings(
                wehago_form_folder=self._wehago_form_edit.text().strip(),
                upload_export_folder=self._upload_export_edit.text().strip(),
                wehago_id=self._wehago_id_edit.text().strip(),
                wehago_password=self._wehago_password_edit.text(),
            )
        )
        self.accept()
