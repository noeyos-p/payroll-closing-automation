"""위하고 자동 업로드 화면 — 이 앱이 만든 "위하고 업로드용" 엑셀 파일들을
골라서 `core/wehago_upload.py`의 브라우저 자동화로 한 번에 올린다.

`ui/main_window.py`의 다운로드 버튼들과 마찬가지로, 백그라운드 작업 결과는
`queue.Queue`에만 쌓고 이 다이얼로그의 `QTimer`가 메인 스레드에서 주기적으로
비운다 — 백그라운드 스레드에서 Qt 시그널을 직접 건드리면 나던 크래시를
피하려고 이 프로젝트가 이미 정한 규칙(`ui/batch_worker.py` 모듈 docstring
참고)을 그대로 따른다."""

from __future__ import annotations

import queue
from pathlib import Path

from PySide6.QtCore import QTimer
from PySide6.QtWidgets import (
    QAbstractItemView,
    QDialog,
    QFileDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
)

from core.wehago_filename import UploadFileInfo, parse_upload_filename

_COLUMNS = ["파일명", "업체", "귀속월", "지급구분", "상태"]


def pick_saved_upload_files(parent) -> list[Path]:
    """이미 저장해둔 위하고 업로드용 파일을 고른다 — 테스트하면서 같은 파일을
    매번 다시 저장하지 않고 바로 올리기 위함(사용자 요청, 2026-09-29).
    환경설정의 업로드용 저장 폴더에서 시작한다."""
    from core.settings import load_settings

    folder = load_settings().upload_export_folder
    start_dir = folder if folder and Path(folder).is_dir() else ""
    paths, _ = QFileDialog.getOpenFileNames(parent, "저장된 위하고 업로드 파일 선택", start_dir, "Excel Files (*.xlsx)")
    return [Path(path) for path in paths]


class WehagoUploadDialog(QDialog):
    def __init__(self, parent=None, initial_paths: list[Path] | None = None, auto_start: bool = False,
                 expected_totals: dict[Path, int] | None = None):
        super().__init__(parent)
        self._expected_totals = expected_totals or {}
        self.setWindowTitle("위하고 자동 업로드")
        self.resize(760, 480)

        self._items: dict[int, UploadFileInfo] = {}
        self._row_of_item: dict[int, int] = {}
        self._next_item_id = 0
        self._result_queue: "queue.Queue[tuple[str, object]]" = queue.Queue()
        self._pending = 0

        layout = QVBoxLayout(self)

        btn_row = QHBoxLayout()
        add_btn = QPushButton("파일 추가")
        add_btn.clicked.connect(self._on_add_files)
        btn_row.addWidget(add_btn)

        remove_btn = QPushButton("선택 제거")
        remove_btn.clicked.connect(self._on_remove_selected)
        btn_row.addWidget(remove_btn)

        btn_row.addStretch(1)
        btn_row.addWidget(QLabel("동시 처리 개수"))
        self._concurrency_spin = QSpinBox()
        self._concurrency_spin.setRange(1, 20)
        # 2026-09-22 실사: 새 배포본에서 기본값 5로 처음 시도했을 때 브라우저가
        # 막 뜨자마자 탭 여러 개를 동시에 여는 부하로 보이는 크래시("Target
        # page, context or browser has been closed")가 났다 — 낮은 값에서
        # 시작해서 안정적이면 사용자가 직접 늘리도록 기본값을 낮춰둔다.
        self._concurrency_spin.setValue(2)
        btn_row.addWidget(self._concurrency_spin)
        layout.addLayout(btn_row)

        self._table = QTableWidget(0, len(_COLUMNS))
        self._table.setHorizontalHeaderLabels(_COLUMNS)
        self._table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self._table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self._table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self._table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self._table, stretch=1)

        note = QLabel(
            "시작하면 브라우저 창이 뜹니다 — 환경설정에 위하고 아이디/비밀번호를 저장해뒀으면 "
            "자동으로 로그인하고, 비워뒀으면 직접 로그인해주세요(로그인 후 자동으로 이어서 진행). "
            "이미 열려 있는 위하고 메인 화면(검색창이 보이는 상태)이 있으면 로그인 없이 그 창을 그대로 씁니다."
        )
        note.setWordWrap(True)
        layout.addWidget(note)

        bottom_row = QHBoxLayout()
        bottom_row.addStretch(1)
        self._start_btn = QPushButton("자동 업로드 시작")
        self._start_btn.clicked.connect(self._on_start)
        bottom_row.addWidget(self._start_btn)
        layout.addLayout(bottom_row)

        self._poll_timer = QTimer(self)
        self._poll_timer.timeout.connect(self._drain_result_queue)
        self._poll_timer.start(150)

        for path in initial_paths or []:
            self._add_file(path)

        if auto_start and self._items:
            # 사용자 요청(2026-09-21): 확정한 항목을 저장한 뒤 바로 이어서
            # 업로드까지 자동으로 시작 — 이 창을 열자마자 다시 "시작" 버튼을
            # 누르게 하지 않는다. 생성자 안에서 바로 부르면 이 다이얼로그가
            # 아직 화면에 보이기 전이라 사용자가 창이 뜨는 것도 못 볼 수
            # 있으니 한 틱 늦춘다.
            QTimer.singleShot(0, self._on_start)

    # ---------- 목록 관리 ----------

    def _add_file(self, path: Path) -> None:
        if any(info.path == path for info in self._items.values()):
            return  # 같은 파일을 두 번 고르면 두 번 올라가지 않도록 무시
        info = parse_upload_filename(path)
        row = self._table.rowCount()
        self._table.insertRow(row)
        self._table.setItem(row, 0, QTableWidgetItem(path.name))
        if info is None:
            for col in range(1, len(_COLUMNS)):
                self._table.setItem(row, col, QTableWidgetItem("-" if col != 4 else "파일명 인식 불가"))
            return
        item_id = self._next_item_id
        self._next_item_id += 1
        self._items[item_id] = info
        self._row_of_item[item_id] = row
        self._table.setItem(row, 1, QTableWidgetItem(info.company))
        self._table.setItem(row, 2, QTableWidgetItem(f"{info.month:02d}월"))
        self._table.setItem(row, 3, QTableWidgetItem(info.pay_type or "(구분 없음)"))
        self._table.setItem(row, 4, QTableWidgetItem("대기"))
        self._table.item(row, 0).setData(1000, item_id)  # Qt.UserRole 대용

    def _on_add_files(self) -> None:
        paths = pick_saved_upload_files(self)
        for path in paths:
            self._add_file(Path(path))

    def _on_remove_selected(self) -> None:
        if self._pending:
            QMessageBox.information(self, "처리 중", "업로드가 끝난 뒤에 목록을 편집하세요.")
            return
        rows = sorted({index.row() for index in self._table.selectedIndexes()}, reverse=True)
        for row in rows:
            self._table.removeRow(row)
        # 예전엔 화면에서만 지워서, 지운 파일도 "시작"하면 그대로 올라갔다
        # (2026-09-29 발견) — 남은 행 기준으로 업로드 대상과 행 번호를 다시 맞춘다.
        items: dict[int, UploadFileInfo] = {}
        row_of_item: dict[int, int] = {}
        for row in range(self._table.rowCount()):
            item_id = self._table.item(row, 0).data(1000)
            if item_id is not None and item_id in self._items:
                items[item_id] = self._items[item_id]
                row_of_item[item_id] = row
        self._items = items
        self._row_of_item = row_of_item

    # ---------- 실행 ----------

    def _on_start(self) -> None:
        if not self._items:
            QMessageBox.information(self, "업로드할 파일 없음", "먼저 파일을 추가하세요.")
            return
        self._start_btn.setEnabled(False)
        self._start_btn.setText("처리 중…")
        self._pending = len(self._items)
        for item_id in self._items:
            row = self._row_of_item[item_id]
            self._table.setItem(row, 4, QTableWidgetItem("대기 중"))

        from core.settings import load_settings
        from core.wehago_upload import start_wehago_upload_batch

        settings = load_settings()
        start_wehago_upload_batch(
            list(self._items.values()),
            self._result_queue,
            max_concurrent=self._concurrency_spin.value(),
            wehago_id=settings.wehago_id,
            wehago_password=settings.wehago_password,
            expected_totals=self._expected_totals,
        )

    def _drain_result_queue(self) -> None:
        try:
            while True:
                kind, payload = self._result_queue.get_nowait()
                if kind == "wehago_upload_item_done":
                    self._on_item_done(payload)
                elif kind == "wehago_upload_batch_finished":
                    self._on_batch_finished(payload)
        except queue.Empty:
            pass

    def _find_row_for_outcome(self, outcome) -> int | None:
        for item_id, info in self._items.items():
            if info.path == outcome.path:
                return self._row_of_item.get(item_id)
        return None

    def _on_item_done(self, outcome) -> None:
        row = self._find_row_for_outcome(outcome)
        if row is not None:
            status = "완료" if outcome.ok else f"확인 필요 — {outcome.message}"
            self._table.setItem(row, 4, QTableWidgetItem(status))
        self._pending = max(0, self._pending - 1)

    def _on_batch_finished(self, error_message: str | None) -> None:
        self._start_btn.setEnabled(True)
        self._start_btn.setText("자동 업로드 시작")
        if error_message:
            QMessageBox.critical(self, "업로드 중단", error_message)
        else:
            QMessageBox.information(self, "완료", "위하고 자동 업로드가 끝났습니다. 상태 열을 확인하세요.")
