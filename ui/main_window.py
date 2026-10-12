"""배치 화면 — 여러 급여대장을 한 번에 올리면 사업장을 자동 인식해서(등록된
위하고 참고파일을 재사용, core/registry.py) 한꺼번에 처리한다.

한 번이라도 위하고 파일을 직접 골라 처리에 성공한 사업장은 자동으로
등록/갱신된다(core/batch.py의 `process_with_reference`, `auto_register=True`
기본값) — 버튼 없이 조용히 저장된다(사용자 지시). 다음부터 그 사업장의
급여대장은 원본만 올려도 자동 인식된다. 처음 보는 사업장이거나 인식이 애매한
파일만 "확인 필요"로 남아 위하고 파일을 그 파일에 한해 직접 고르면 된다 —
한 파일이 애매해도 나머지 배치는 그대로 처리된다.

처리는 전부 백그라운드 스레드(ui/batch_worker.py)에서 돌린다 — 여러 파일을
순서대로 열고 계산하는 게 파일당 수 초씩 걸릴 수 있어, 화면 스레드에서 직접
돌리면 창이 멈춘 것처럼 보인다(다른 프로젝트가 AI 호출로 실제로 겪었던 문제와
같은 원인).

**2026-09-08부터 QThread+Qt 시그널이 아니라 순수 threading.Thread+queue.Queue를
쓴다** — 자세한 이유는 ui/batch_worker.py 모듈 docstring 참고(반복 재현되던
네이티브 크래시가 PySide6 버전과 무관하게 항상 Qt의 크로스스레드 시그널 전달
경로에서 나서, 그 경로 자체를 안 타게 바꿈). 백그라운드 스레드는 결과를
`self._result_queue`에 넣기만 하고, 화면 갱신은 메인 스레드의 `self._poll_timer`
(QTimer)가 주기적으로 그 큐를 비우면서 한다 — 큐를 비우는 동안은 이미 메인
스레드 안이므로 언제나 `self.` 바운드 메서드를 그대로 호출해도 안전하다."""

from __future__ import annotations

import html
import queue
from datetime import date
from pathlib import Path

from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSplitter,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from core import registry
from core.batch import BatchItem, plan_batch_add
from core.engine import list_dest_sheet_candidates, list_source_sheet_candidates, supported_extensions
from core.export import export_wehago
from core.profile import (
    derive_client_label,
    parse_template_token,
    payroll_type_from_filename,
    payroll_type_from_sheet_name,
)
from core.reconcile import learned_mapping_key
from core.settings import load_settings, save_settings
from core.wehago_sync import SyncResult
from ui.batch_worker import start_batch_add, start_single_run, start_wehago_folder_sync
from ui.company_manager_dialog import CompanyManagerDialog
from ui.settings_dialog import SettingsDialog
from ui.wehago_upload_dialog import WehagoUploadDialog, pick_saved_upload_files
from ui.drop_zone import DropZone
from ui.review_table import ReviewTable
from ui.verification_dialog import build_mapping_table, build_verification_table, summary_text

_STATUS_TEXT = {
    "pending": "⏳ 처리 중…",
    "matched": "✅ 정상",
    "needs_reference": "⚠️ 확인 필요",
    "error": "❌ 오류",
}

# 사원확인이 필요할 때 큐 목록 맨 오른쪽에 띄우는 경고 아이콘 — 사용자 요청:
# "사원확인필요"라는 글자 문구 대신 얼굴 이모티콘 + 빨간색 느낌표로.
_EMPLOYEE_WARN_ICON = "😟❗"


def _amount_status_text(item: BatchItem) -> str:
    """큐 목록 행 왼쪽(상태+사업장명 칸)에 보여줄 금액 상태 문구.
    `summary.diff`(수당/공제/차인지급액 합계가 원본과 다른 사원 수)만 본다
    — 순수하게 "숫자가 안 맞는다"는 신호. 사원 확인 여부는 별도로
    `_needs_employee_check`가 판단해서 행 오른쪽에 아이콘으로만 표시한다
    (사용자 요청: 두 문제는 서로 다르니 한쪽이 다른 쪽 표시를 가리면 안 됨 —
    왼쪽 금액/오른쪽 사원 두 칸으로 완전히 분리)."""
    if item.status == "matched" and item.result is not None:
        return "⚠️ 금액확인" if item.result.summary.diff else "✅ 금액정상"
    return _STATUS_TEXT.get(item.status, item.status)


def _needs_employee_check(item: BatchItem) -> bool:
    """큐 목록 행 오른쪽에 `_EMPLOYEE_WARN_ICON`을 띄울지 여부.
    `summary.added`(원본엔 있는데 위하고 사원등록에 아예 없어서 신규 행으로
    추가한 사원)와 `summary.need_check`(위하고에 앞 3글자가 같은 이름 후보는
    있지만 생년월일 등으로 확정을 못 해 자동 매칭하지 않은 사원,
    `core/reconcile.py`의 "확인필요" 상태 — memo가 항상 "위하고에 ... 이름이
    있으나 자동 매칭되지 않았습니다" 류로, 금액이 아니라 **그 사람이 누구인지
    확정이 안 됐다는** 문제라 금액 쪽이 아니라 여기서 판단하는 게 맞다) 중
    하나라도 있으면 True."""
    if item.status == "matched" and item.result is not None:
        summary = item.result.summary
        return bool(summary.added or summary.need_check)
    return False


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("급여대장 자동 변환 — 일괄 처리")
        self.resize(1400, 860)

        # 큐 안에서 항목의 정체성은 화면상 위치(행 번호)가 아니라 추가될 때
        # 한 번 발급되는 고정 id(_next_item_id)로 관리한다 — 행 번호를 쓰면
        # 다른 항목이 아직 백그라운드에서 처리 중일 때 목록에서 하나를 지워
        # 나머지 행이 당겨지는 순간, 그 워커가 스폰될 때 붙잡아둔 옛 행
        # 번호로 완료 결과를 엉뚱한 자리에 덮어써버린다(사용자 요청으로
        # "처리 중에도 지울 수 있게" 해달라고 해서 이 방식으로 바꿈,
        # 2026-09-08). id는 한 번 배정되면 항목이 지워질 때까지 안 바뀐다.
        self._next_item_id: int = 0
        self._items: dict[int, BatchItem] = {}
        # {항목 id -> (행 위젯, 왼쪽 라벨, 오른쪽 라벨)} — 큐 목록 한 행을
        # "왼쪽(상태+사업장명, 늘어남) / 오른쪽(사원확인 경고 아이콘, 고정폭
        # 오른쪽 정렬)"로 나눠 보여주려고 setItemWidget으로 붙인 위젯.
        # QListWidgetItem 텍스트만으로는 한 줄 안에서 일부만 오른쪽 정렬할
        # 방법이 없어 위젯을 쓴다. 항목이 추가될 때 한 번만 만들고 이후로는
        # 라벨 텍스트만 갱신한다(_update_queue_row) — 이 위젯 자체를
        # deleteLater()하거나 다시 만드는 코드는 절대 넣지 않는다(다른 곳의
        # QTabWidget 위젯 삭제가 STATUS_HEAP_CORRUPTION을 일으킨 전례,
        # `_dispose_extra_tabs` 주석 참고 — 여기서도 같은 위험을 피하려고
        # "메모리가 좀 새는 것보다 앱이 죽는 게 낫다" 원칙을 그대로 따른다).
        self._queue_row_widgets: dict[int, tuple[QWidget, QCheckBox, QLabel, QLabel]] = {}
        # 사용자가 "확정" 체크박스를 켠 항목 id 집합 — 위하고 자동 업로드는
        # 여기 있는 것만 대상으로 한다(사용자 요청, 2026-09-21: "확정을
        # 지은것만"). `BatchItem`에 넣지 않고 따로 두는 이유는 재처리(수동
        # 매칭 적용/위하고 파일 재지정 등)마다 `self._items[item_id]`가 새
        # 객체로 통째로 바뀌는데, 그때마다 이전 확정 여부가 새 결과에도
        # 그대로 남으면 "확인 안 한 새 계산 결과가 조용히 업로드"되는
        # 위험한 상황이 생긴다 — `_update_queue_row`가 호출될 때마다(=새
        # 결과가 반영될 때마다) 무조건 지워서 매번 다시 확정하게 한다.
        self._confirmed_ids: set[int] = set()
        # {항목 id -> {소스열번호 -> 목적지키 또는 None}} — 항목매핑결과 탭의
        # "선택한 매칭 적용"으로 쌓이는 수동 오버라이드. 파일별로 독립적으로 유지.
        self._manual_overrides: dict[int, dict[int, str | None]] = {}
        self._combo_by_source_col: dict[int, tuple[QComboBox, str | None]] = {}
        # 백그라운드 스레드(ui/batch_worker.py)가 결과를 여기 넣는다 —
        # queue.Queue는 스레드 안전(내부적으로 락 사용)이라 어느 스레드에서
        # put()해도 안전하다. 화면 갱신은 이 큐를 절대 직접 읽지 않고,
        # 항상 메인 스레드의 _poll_timer가 주기적으로 비우면서 한다(Qt
        # 시그널을 크로스스레드로 쓰지 않기 위해 — 클래스 docstring 참고).
        self._result_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self._poll_timer = QTimer(self)
        self._poll_timer.timeout.connect(self._drain_result_queue)
        self._poll_timer.start(100)

        central = QWidget(self)
        outer = QVBoxLayout(central)
        outer.setContentsMargins(20, 16, 20, 16)
        outer.setSpacing(10)

        outer.addLayout(self._build_top_bar())

        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.addWidget(self._build_queue_panel())
        splitter.addWidget(self._build_detail_panel())
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([340, 1060])
        outer.addWidget(splitter, stretch=1)

        self.setCentralWidget(central)

    # ---------- 화면 구성 ----------

    def _build_top_bar(self) -> QHBoxLayout:
        row = QHBoxLayout()
        add_btn = QPushButton("급여대장 추가 (여러 개 선택 가능)")
        add_btn.clicked.connect(self._on_pick_files)
        row.addWidget(add_btn)

        self._drop_zone = DropZone()
        self._drop_zone.setFixedHeight(36)
        self._drop_zone.files_dropped.connect(self._on_files_added)
        zone_layout = QHBoxLayout(self._drop_zone)
        zone_layout.setContentsMargins(10, 4, 10, 4)
        zone_layout.addWidget(QLabel("여기로 파일을 끌어다 놓아도 됩니다."))
        row.addWidget(self._drop_zone, stretch=1)

        manage_btn = QPushButton("사업장 관리")
        manage_btn.clicked.connect(self._on_manage_companies)
        row.addWidget(manage_btn)

        self._sync_wehago_btn = QPushButton("위하고 양식 폴더 업데이트")
        self._sync_wehago_btn.setToolTip(load_settings().wehago_form_folder)
        self._sync_wehago_btn.clicked.connect(self._on_sync_wehago_folder)
        row.addWidget(self._sync_wehago_btn)

        settings_btn = QPushButton("환경설정")
        settings_btn.clicked.connect(self._on_open_settings)
        row.addWidget(settings_btn)

        download_selected_btn = QPushButton("선택 항목 다운로드")
        download_selected_btn.clicked.connect(self._on_download_selected)
        row.addWidget(download_selected_btn)

        download_all_btn = QPushButton("전체 다운로드")
        download_all_btn.clicked.connect(self._on_download_all)
        row.addWidget(download_all_btn)

        wehago_upload_btn = QPushButton("위하고 자동 업로드")
        wehago_upload_btn.setToolTip(
            "큐에서 '확정' 체크한 항목만 골라 저장 후 위하고에 브라우저 자동화로 올립니다."
        )
        wehago_upload_btn.clicked.connect(self._on_wehago_auto_upload)
        row.addWidget(wehago_upload_btn)

        saved_upload_btn = QPushButton("저장된 파일 업로드")
        saved_upload_btn.setToolTip(
            "이미 저장해둔 위하고 업로드용 파일을 골라 다시 저장하지 않고 바로 올립니다."
        )
        saved_upload_btn.clicked.connect(self._on_upload_saved_files)
        row.addWidget(saved_upload_btn)
        return row

    def _build_queue_panel(self) -> QWidget:
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(QLabel("처리 목록 (여러 개 선택 가능)"))
        self._queue_list = QListWidget()
        self._queue_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self._queue_list.currentRowChanged.connect(self._on_select_item)
        layout.addWidget(self._queue_list, stretch=1)

        remove_btn = QPushButton("선택 항목 목록에서 지우기")
        remove_btn.clicked.connect(self._on_remove_selected)
        layout.addWidget(remove_btn)
        return panel

    def _build_detail_panel(self) -> QWidget:
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)

        self._status_label = QLabel("왼쪽 위 \"급여대장 추가\"로 파일을 올리세요.")
        self._status_label.setWordWrap(True)
        self._status_label.setTextFormat(Qt.TextFormat.RichText)
        layout.addWidget(self._status_label)

        layout.addWidget(self._build_reference_row())

        self._tabs = QTabWidget()
        layout.addWidget(self._tabs, stretch=1)

        review_tab = QWidget()
        review_layout = QVBoxLayout(review_tab)
        review_layout.setContentsMargins(0, 0, 0, 0)
        self._table = ReviewTable(self)
        review_layout.addWidget(self._table, stretch=1)
        self._selection_summary_label = QLabel("")
        self._table.selection_summary_changed.connect(self._selection_summary_label.setText)
        review_layout.addWidget(self._selection_summary_label)
        self._tabs.addTab(review_tab, "검토표")

        bottom_row = QHBoxLayout()
        bottom_row.addStretch(1)
        download_btn = QPushButton("이 파일 다운로드")
        download_btn.clicked.connect(self._on_download_current)
        bottom_row.addWidget(download_btn)
        layout.addLayout(bottom_row)
        return panel

    def _build_reference_row(self) -> QFrame:
        frame = QFrame()
        row = QHBoxLayout(frame)
        row.setContentsMargins(0, 0, 0, 0)
        row.addWidget(QLabel("사업장명"))
        self._client_name_edit = QLineEdit()
        self._client_name_edit.setFixedWidth(220)
        row.addWidget(self._client_name_edit)
        pick_ref_btn = QPushButton("위하고 파일 선택")
        pick_ref_btn.clicked.connect(self._on_pick_reference_for_current)
        row.addWidget(pick_ref_btn)
        change_source_sheet_btn = QPushButton("원본 시트 변경")
        change_source_sheet_btn.clicked.connect(self._on_change_source_sheet)
        row.addWidget(change_source_sheet_btn)
        row.addStretch(1)
        frame.setVisible(False)
        self._reference_row = frame
        return frame

    # ---------- 파일 추가 ----------

    def _make_queue_row_widget(self, item_id: int, left_text: str) -> QWidget:
        """큐 목록 한 행을 체크박스(확정)/왼쪽(상태+사업장명)/오른쪽(사원확인
        경고 아이콘) 세 칸으로 보여주는 위젯을 만들어 `self._queue_row_widgets`에
        등록하고 돌려준다. 오른쪽 칸은 `_EMPLOYEE_WARN_ICON` 전용 고정폭 칸이라
        처음엔 빈 문자열로 시작 — 결과가 나오면 `_update_queue_row`가 채운다."""
        widget = QWidget()
        layout = QHBoxLayout(widget)
        layout.setContentsMargins(6, 2, 6, 2)
        layout.setSpacing(8)
        confirm_box = QCheckBox("확정")
        confirm_box.toggled.connect(lambda checked, item_id=item_id: self._on_confirm_toggled(item_id, checked))
        left_label = QLabel(left_text)
        right_label = QLabel("")
        right_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        layout.addWidget(confirm_box, 0)
        layout.addWidget(left_label, 1)
        layout.addWidget(right_label, 0)
        self._queue_row_widgets[item_id] = (widget, confirm_box, left_label, right_label)
        return widget

    def _on_confirm_toggled(self, item_id: int, checked: bool) -> None:
        if checked:
            self._confirmed_ids.add(item_id)
        else:
            self._confirmed_ids.discard(item_id)

    def _update_queue_row(self, item_id: int, left_text: str, warn: bool) -> None:
        """이미 만들어진 행 위젯의 라벨 텍스트만 갱신한다(위젯 자체는 절대
        새로 안 만듦 — `_queue_row_widgets` 주석 참고). 이 함수는 그 항목에
        새 처리 결과가 반영될 때마다 불리므로, 이전 "확정" 체크는 항상
        초기화한다 — 확정 당시와 다른 계산 결과가 확인 없이 그대로
        업로드되는 걸 막기 위함(클래스 docstring의 `_confirmed_ids` 주석
        참고)."""
        entry = self._queue_row_widgets.get(item_id)
        if entry is None:
            return
        _, confirm_box, left_label, right_label = entry
        confirm_box.setChecked(False)
        self._confirmed_ids.discard(item_id)
        left_label.setText(left_text)
        right_label.setText(_EMPLOYEE_WARN_ICON if warn else "")

    def _on_pick_files(self) -> None:
        exts = " ".join(f"*{e}" for e in sorted(supported_extensions()))
        paths, _ = QFileDialog.getOpenFileNames(self, "급여대장 선택", "", f"급여대장 파일 ({exts})")
        if paths:
            self._on_files_added(paths)

    def _on_files_added(self, path_strs: list[str]) -> None:
        # 한 워크북에 추급/급여/급+상 시트가 나뉘어 있거나, 아예 별도
        # 파일로 나뉘어 있으면(사용자 요청, 2026-09-08 — "시트 말고도
        # 엑셀파일이 나눠져있는경우도있어") 큐 항목 여러 개가 하나로 묶여
        # 처리된다 — `plan_batch_add`가 가벼운 구조/파일명 스캔만으로
        # 판단하므로(무거운 계산 자체는 여전히 아래 `start_batch_add`가
        # 배경 스레드로 돌림) 파일 수가 많아도 화면이 멈추지 않는다. 같은
        # 종류(예: "추급1"/"추급2", 또는 "...(추급1).xlsx"/"...(추급2).xlsx")
        # 로 나뉜 건 한 그룹으로 묶여서 합쳐진다(동명이인+생년월일 같으면
        # 항목별 금액 합산 — `merge_sources`).
        paths = [Path(s) for s in path_strs]
        groups = plan_batch_add(paths)

        indexed_paths: list[tuple[int, Path, str | None, list[tuple[Path, str | None]] | None, bool]] = []
        for group in groups:
            # 같은 파일 안에서 시트로만 나뉜 경우에만 시트 이름을 파일명보다
            # 우선하는 라벨링(was_split)을 켠다 — 별도 파일로 나뉜 경우는
            # 각자의 파일명이 이미 명확한 신호라 그럴 필요가 없다.
            same_file = len({p for p, _ in group}) == 1
            primary_path, primary_sheet = group[0]
            # "(급여)(급+상)" 파일을 사업장 수정본 양식 종류별로 나눈 항목
            # (`<양식:급+상>` 표시)도 시트 쪽 종류로 라벨을 붙인다.
            was_split = (len(group) > 1 and same_file) or parse_template_token(primary_sheet) is not None

            item_id = self._next_item_id
            self._next_item_id += 1
            self._items[item_id] = BatchItem(primary_path, "pending")
            if len(group) == 1:
                label = primary_path.name if primary_sheet is None else f"{primary_path.name} — {primary_sheet}"
            elif same_file:
                label = f"{primary_path.name} — {'+'.join(s for _, s in group if s)}"
            else:
                label = " + ".join(p.name for p, _ in group)
            list_item = QListWidgetItem()
            list_item.setData(Qt.ItemDataRole.UserRole, item_id)
            self._queue_list.addItem(list_item)
            self._queue_list.setItemWidget(
                list_item, self._make_queue_row_widget(item_id, f"{_STATUS_TEXT['pending']}  {label}")
            )

            if len(group) == 1:
                source_sheet, merge_sources = primary_sheet, None
            else:
                source_sheet, merge_sources = None, group
            indexed_paths.append((item_id, primary_path, source_sheet, merge_sources, was_split))
        if self._queue_list.currentRow() < 0 and self._items:
            self._queue_list.setCurrentRow(0)

        start_batch_add(indexed_paths, self._result_queue)

    def _drain_result_queue(self) -> None:
        """`self._poll_timer`가 100ms마다 메인 스레드 안에서 부른다 — 백그라운드
        스레드(ui/batch_worker.py)가 쌓아둔 결과를 전부 소비할 때까지 반복
        (`get_nowait()`가 비어있으면 `queue.Empty`로 바로 끝남 — 블로킹 없음)."""
        while True:
            try:
                kind, payload = self._result_queue.get_nowait()
            except queue.Empty:
                return
            if kind == "item_done":
                item_id, item = payload
                self._on_item_done(item_id, item)
            elif kind == "wehago_synced":
                self._on_wehago_folder_synced(payload)
            # "batch_add_finished"는 지금은 화면에서 딱히 할 일이 없어 무시.

    def _on_item_done(self, item_id: int, item: BatchItem) -> None:
        if item_id not in self._items:
            return  # 처리 도중 사용자가 목록에서 이미 지운 항목 — 결과는 조용히 버린다
        self._items[item_id] = item
        row = self._row_for_item_id(item_id)
        if row is None:
            return
        label = self._type_tagged_label(item.company_name or item.path.name, item)
        self._update_queue_row(item_id, f"{_amount_status_text(item)}  {label}", _needs_employee_check(item))
        if self._current_item_id() == item_id:
            # 재계산이 끝나자마자 이 자리에서 바로 검토표/탭을 다시 만들면,
            # 사용자가 방금 그 자리의 콤보박스(항목매핑결과 "직접 매칭")를
            # 조작한 직후라 아직 OS/IME 쪽에서 그 위젯과 관련된 처리가 안
            # 끝났을 수 있다 — 실사에서 바로 이 타이밍에 QTableWidget()
            # 생성 자체가 멈춰버리는 응답 없음이 재현됐다(2026-09-08).
            # 한 틱(QTimer.singleShot(0, ...))만 늦춰서 현재 진행 중인
            # 이벤트 처리가 먼저 다 끝나게 한다 — 그동안 다른 파일을
            # 클릭해서 선택이 바뀌었으면 그때 가서 다시 확인하고 조용히
            # 건너뛴다.
            QTimer.singleShot(0, lambda item_id=item_id: self._refresh_if_current(item_id))

    def _refresh_if_current(self, item_id: int) -> None:
        if self._current_item_id() == item_id and item_id in self._items:
            self._show_item(item_id)

    # ---------- 목록에서 파일 선택 ----------

    def _row_for_item_id(self, item_id: int) -> int | None:
        for row in range(self._queue_list.count()):
            if self._queue_list.item(row).data(Qt.ItemDataRole.UserRole) == item_id:
                return row
        return None

    def _item_id_for_row(self, row: int) -> int | None:
        if row < 0 or row >= self._queue_list.count():
            return None
        return self._queue_list.item(row).data(Qt.ItemDataRole.UserRole)

    def _on_select_item(self, row: int) -> None:
        item_id = self._item_id_for_row(row)
        if item_id is None:
            return
        self._show_item(item_id)

    def _current_item_id(self) -> int | None:
        return self._item_id_for_row(self._queue_list.currentRow())

    def _on_remove_selected(self) -> None:
        # 큐 항목의 정체성은 행 번호가 아니라 고정 id라서(클래스 docstring 겸
        # __init__ 주석 참고), 다른 파일이 아직 백그라운드에서 처리 중이어도
        # 안전하게 지울 수 있다 — 그 워커가 나중에 끝나서 `_on_item_done`이
        # 불려도 이미 지워진 id는 `self._items`에 없어 조용히 무시된다
        # (사용자 요청, 2026-09-08: 처리 중에도 지울 수 있게 해달라).
        rows = self._selected_rows()
        if not rows:
            QMessageBox.information(self, "선택 필요", "목록에서 지울 파일을 먼저 선택하세요.")
            return
        for row in sorted(rows, reverse=True):
            item_id = self._item_id_for_row(row)
            self._queue_list.takeItem(row)
            if item_id is not None:
                self._items.pop(item_id, None)
                self._manual_overrides.pop(item_id, None)
                # 위젯 자체는 지우지 않는다(__init__의 `_queue_row_widgets`
                # 주석 참고) — 파이썬 쪽 참조만 정리해서 못 쓰는 항목이 계속
                # 업데이트를 받으려고 하는 일이 없게 한다.
                self._queue_row_widgets.pop(item_id, None)
        if not self._items:
            self._status_label.setText("왼쪽 위 \"급여대장 추가\"로 파일을 올리세요.")
            self._reference_row.setVisible(False)
            self._clear_result_tabs()

    def _show_item(self, item_id: int) -> None:
        item = self._items[item_id]
        self._combo_by_source_col = {}

        if item.status == "pending":
            self._status_label.setText("처리 중입니다…")
            self._reference_row.setVisible(False)
            self._clear_result_tabs()
            return

        if item.status == "error":
            self._status_label.setText(f"<span style='color:#c62828;'>오류: {item.error}</span>")
            self._reference_row.setVisible(True)
            self._client_name_edit.setText(item.company_name or derive_client_label(item.path.stem))
            self._clear_result_tabs()
            return

        if item.status == "needs_reference":
            names = ", ".join(f"{c.company_name}({c.ratio:.0%})" for c in item.match_candidates[:3])
            hint = f" — 비슷한 후보: {names}" if names else ""
            self._status_label.setText(
                f"이 파일의 사업장을 자동으로 인식하지 못했습니다{hint}. "
                "위하고 파일을 직접 선택하면 그 사업장으로 등록되어 다음부터 자동 인식됩니다."
            )
            self._reference_row.setVisible(True)
            self._client_name_edit.setText(derive_client_label(item.path.stem))
            self._clear_result_tabs()
            return

        # matched
        self._reference_row.setVisible(True)
        self._client_name_edit.setText(item.company_name or "")
        summary = item.result.summary
        # 큐 목록의 _queue_status_text와 같은 원칙 — 금액(diff)과 사원확인
        # (added/need_check)은 서로 다른 문제라 배지를 따로 둔다.
        if summary.diff:
            amount_badge = "<span style='color:#ef6c00;font-weight:bold;'>⚠️ 금액확인</span>"
        else:
            amount_badge = "<span style='color:#2e7d32;font-weight:bold;'>✅ 금액정상</span>"
        employee_badge = (
            " <span style='color:#1565c0;font-weight:bold;'>🧑 사원확인필요</span>"
            if (summary.added or summary.need_check)
            else ""
        )
        # 처리 메모(저장된 수정본 양식으로 변환했는지, 덧붙이거나 비운 항목 등 —
        # core/source_template.py) — 사람이 결과를 볼 때 같이 보이게 한다.
        notes_html = "".join(
            f"<br><span style='color:#555;'>📝 {html.escape(n)}</span>" for n in item.result.notes
        )
        self._status_label.setText(
            f"{amount_badge}{employee_badge} — {item.company_name}: {summary_text(item.result)}{notes_html}"
        )
        self._table.load_records(item.result.records, item.result.profile, item.result.ambiguous_dest_keys)
        self._rebuild_result_tabs(item)

    def _dispose_extra_tabs(self) -> None:
        # QTabWidget.removeTab()은 탭 목록에서만 빼고 위젯 자체는 지우지
        # 않는다(Qt 문서에 명시) — 그래서 재계산할 때마다 예전 위젯(항목매핑결과
        # 탭의 콤보박스들 포함)이 부모 없이 메모리에 계속 남는다. 명시적으로
        # `deleteLater()`를 걸어 정리해봤는데(2026-09-08), 그 직후 이 앱이
        # STATUS_HEAP_CORRUPTION으로 강제 종료되는 게 Windows 이벤트 로그에
        # 잡혔다 — 콤보박스를 셀 위젯으로 쓰는 QTableWidget을 지울 때 걸리는
        # PySide6/Qt 쪽 문제로 추정(정확한 원인은 원격으로 더 못 팜). 메모리가
        # 좀 새는 것보다 앱이 죽는 게 훨씬 나쁘므로, deleteLater 호출은 다시
        # 뺐다 — removeTab만 하고 그대로 둔다(기존 동작으로 복귀).
        while self._tabs.count() > 1:
            self._tabs.removeTab(1)

    def _clear_result_tabs(self) -> None:
        self._dispose_extra_tabs()
        self._table.setRowCount(0)
        self._table.setColumnCount(0)

    def _rebuild_result_tabs(self, item: BatchItem) -> None:
        self._dispose_extra_tabs()
        result = item.result
        verification_table = build_verification_table(result.verification_rows)
        self._tabs.addTab(verification_table, f"검증결과 ({len(result.verification_rows)}명)")

        mapping_tab = QWidget()
        mapping_layout = QVBoxLayout(mapping_tab)
        mapping_layout.setContentsMargins(0, 0, 0, 0)
        mapping_table, self._combo_by_source_col = build_mapping_table(result.mapping_rows, result.profile)
        mapping_layout.addWidget(mapping_table, stretch=1)

        bottom_row = QHBoxLayout()
        hint = QLabel("미매칭 항목의 \"직접 매칭\" 칸에서 위하고 열을 고른 뒤 적용하면, 이 파일만 다시 계산합니다.")
        hint.setWordWrap(True)
        bottom_row.addWidget(hint, stretch=1)
        apply_button = QPushButton("선택한 매칭 적용")
        apply_button.clicked.connect(self._on_apply_mapping_tab)
        bottom_row.addWidget(apply_button)
        mapping_layout.addLayout(bottom_row)
        self._tabs.addTab(mapping_tab, f"항목매핑결과 ({len(result.mapping_rows)}건)")

    # ---------- 위하고 파일 수동 지정 / 시트 변경 ----------

    def _choose_sheet(self, path: Path, candidates: list, title: str) -> str | None:
        if len(candidates) <= 1:
            return candidates[0].name if candidates else None
        labels = [f"{c.name} (약 {c.employee_count}명)" for c in candidates]
        choice, ok = QInputDialog.getItem(
            self, title, f"{path.name}에 후보 시트가 여러 개 있습니다. 사용할 시트를 고르세요.", labels, 0, False
        )
        if not ok:
            return candidates[0].name
        return candidates[labels.index(choice)].name

    def _on_pick_reference_for_current(self) -> None:
        item_id = self._current_item_id()
        if item_id is None:
            return
        item = self._items[item_id]
        client_name = self._client_name_edit.text().strip() or derive_client_label(item.path.stem)
        path_str, _ = QFileDialog.getOpenFileName(self, "위하고 참고파일 선택", "", "Excel Files (*.xlsx)")
        if not path_str:
            return
        reference_path = Path(path_str)
        try:
            candidates = list_dest_sheet_candidates(reference_path)
        except Exception:
            candidates = []
        reference_sheet = self._choose_sheet(reference_path, candidates, "위하고 시트 선택")
        self._run_single(
            item_id, client_name, reference_path, reference_sheet=reference_sheet,
            auto_register=True, was_split=item.was_split, merge_sources=item.merge_sources,
        )

    def _on_change_source_sheet(self) -> None:
        item_id = self._current_item_id()
        if item_id is None:
            return
        item = self._items[item_id]
        try:
            candidates = list_source_sheet_candidates(item.path)
        except Exception as exc:
            QMessageBox.critical(self, "오류", f"시트 목록을 불러오지 못했습니다:\n{exc}")
            return
        if not candidates:
            QMessageBox.information(self, "시트 없음", "이 파일에서 급여 구조를 인식한 시트가 없습니다.")
            return
        labels = [f"{c.name} (약 {c.employee_count}명)" for c in candidates]
        choice, ok = QInputDialog.getItem(self, "원본 시트 변경", f"{item.path.name}에서 사용할 시트를 고르세요.", labels, 0, False)
        if not ok:
            return
        source_sheet = candidates[labels.index(choice)].name
        if item.status != "matched" or item.reference_path is None:
            QMessageBox.information(self, "위하고 파일 필요", "먼저 위하고 파일을 지정한 뒤에 시트를 바꿀 수 있습니다.")
            return
        client_name = self._client_name_edit.text().strip() or item.company_name
        self._manual_overrides.pop(item_id, None)
        self._run_single(
            item_id, client_name, item.reference_path, source_sheet=source_sheet,
            reference_sheet=item.reference_sheet, auto_register=False, company_id=item.company_id,
            was_split=item.was_split,
        )

    def _run_single(
        self,
        item_id: int,
        client_name: str,
        reference_path: Path,
        source_sheet: str | None = None,
        reference_sheet: str | None = None,
        manual_overrides: dict | None = None,
        auto_register: bool = True,
        company_id: int | None = None,
        was_split: bool = False,
        merge_sources: list[tuple[Path, str | None]] | None = None,
    ) -> None:
        path = self._items[item_id].path
        self._items[item_id] = BatchItem(path, "pending")
        self._update_queue_row(item_id, f"{_STATUS_TEXT['pending']}  {path.name}", False)
        if self._current_item_id() == item_id:
            self._status_label.setText("다시 계산 중입니다…")

        start_single_run(
            item_id, path, client_name, reference_path, self._result_queue,
            source_sheet=source_sheet, reference_sheet=reference_sheet,
            manual_overrides=manual_overrides, auto_register=auto_register, company_id=company_id,
            was_split=was_split, merge_sources=merge_sources,
        )

    # ---------- 수동 매칭 적용 ----------

    def _on_apply_mapping_tab(self) -> None:
        item_id = self._current_item_id()
        if item_id is None:
            return
        item = self._items[item_id]
        if item.status != "matched":
            return
        overrides: dict[int, str | None] = {}
        for source_col, (combo, original_dest_key) in self._combo_by_source_col.items():
            chosen = combo.currentData()
            if chosen != original_dest_key:
                overrides[source_col] = chosen
        if not overrides:
            return
        if item.company_id is not None:
            # 이 사업장에 한해 다음 달(그리고 이 앱을 새로 빌드/재배포한
            # 뒤)에도 같은 매칭을 기억해둔다(사용자 요청, 2026-09-08).
            # "선택 안 함"(dest_key=None)도 저장한다(사용자 요청,
            # 2026-09-14) — 예전엔 "일회성 판단"이라며 이 경우만 저장을
            # 건너뛰었는데, 그러면 자동매칭되던 항목을 매번 다시 손으로
            # 꺼줘야 했다. `registry.save_column_mapping`이 빈 문자열을
            # "항상 매칭하지 마라"는 뜻으로 저장해두면
            # `core/reconcile.py`가 build_map의 자동매칭 결과까지 무시하고
            # 강제로 미매칭 처리한다(그 함수 docstring 참고). 저장 키는
            # `learned_mapping_key(row.group, row.item_label)` — 그룹(수당/
            # 공제)까지 포함한다. "성과급"처럼 같은 텍스트가 수당 쪽에도
            # 공제 쪽에도 각각 존재하면서 서로 다른 위하고 열을 가리켜야
            # 하는 실제 사례가 있어서(업체001어르신주야간보호센터 실사,
            # 2026-09-08) — 그룹 없이 항목명만 키로 쓰면 한쪽을 고쳐 저장한
            # 게 다른 쪽까지 덮어써 버렸다.
            key_by_col = {
                row.source_col: learned_mapping_key(row.group, row.item_label) for row in item.result.mapping_rows
            }
            for source_col, dest_key in overrides.items():
                label = key_by_col.get(source_col)
                if label:
                    registry.save_column_mapping(item.company_id, label, dest_key or "")
        merged = {**self._manual_overrides.get(item_id, {}), **overrides}
        self._manual_overrides[item_id] = merged
        client_name = self._client_name_edit.text().strip() or item.company_name
        self._run_single(
            item_id, client_name, item.reference_path,
            source_sheet=item.source_sheet, reference_sheet=item.reference_sheet,
            manual_overrides=merged, auto_register=False, company_id=item.company_id,
            was_split=item.was_split, merge_sources=item.merge_sources,
        )

    # ---------- 다운로드 ----------

    def _type_tagged_label(self, base: str, item: BatchItem) -> str:
        """`base`(거래처명 계열 라벨) 뒤에 "(급여)"/"(추급)"/"(급+상)" 종류
        표시를 붙인다(사용자 요청, 2026-09-08) — 8월지급 급여를 이 세
        종류로 나눠 처리하는 관례가 있는데, 별도 파일로 나뉜 경우엔 원본
        파일명 끝의 "...급여대장(급여)" 표시를 그대로 옮겨 붙이고, 한
        워크북에 세 시트로 나뉜 경우엔(`core/batch.py`의 `plan_batch_add`가
        이미 시트별로 큐 항목을 쪼개놨다) 이 항목이 실제로 처리된 시트
        이름에서 종류를 판단한다.

        어느 쪽을 먼저 볼지는 `item.was_split`으로 정한다 — `False`(별도
        파일)면 파일명이 더 명시적인 신호라 그쪽을 먼저 본다. `True`(한
        워크북을 시트별로 쪼갠 경우)면 **시트 이름을 먼저** 본다: 이런
        파일은 파일명 자체가 "...급여대장(급+상)(추급)"처럼 여러 종류를
        한꺼번에 담고 있을 수 있어서(실사 확인, 2026-09-08 — 업체050
        복지센터), 파일명만 보면 이 항목이 정확히 어느 시트인지 구분이
        안 되고 두 쪼개진 항목이 전부 같은(예: 마지막 괄호) 라벨을 달아
        버린다."""
        if item.was_split and item.source_sheet:
            type_token = payroll_type_from_sheet_name(item.source_sheet)
            if type_token is None:
                type_token = payroll_type_from_filename(item.path.stem)
        else:
            type_token = payroll_type_from_filename(item.path.stem)
            if type_token is None and item.source_sheet:
                type_token = payroll_type_from_sheet_name(item.source_sheet)
        return f"{base}({type_token})" if type_token else base

    def _download_label(self, item: BatchItem) -> str:
        """다운로드 파일명에 쓸 라벨 — "{거래처명} N월귀속..." 관례에서 월을
        자동 인식해 "거래처명07"처럼 만든다(derive_client_label, 이미 사업장
        자동 인식/미등록 파일 표시에 쓰던 것과 동일한 함수). 같은 사업장을
        여러 달 치 한꺼번에 처리해도(전체 다운로드) 파일마다 원본 급여대장
        이름에서 그 달을 다시 뽑으므로 "거래처명06"/"07"/"08"처럼 자연히
        구분된다 — 등록된 사업장명(item.company_name)만 쓰면 달이 달라도
        전부 같은 이름이 되어 서로 덮어써 버리는 문제가 있었다(사용자 실사
        피드백, 2026-09-08). 여기에 추급/급여/급+상 종류 표시까지 더한다."""
        return self._type_tagged_label(derive_client_label(item.path.stem), item)

    def _on_download_current(self) -> None:
        item_id = self._current_item_id()
        if item_id is None:
            return
        item = self._items[item_id]
        if item.result is None:
            QMessageBox.warning(self, "다운로드 불가", "이 파일은 아직 처리되지 않았습니다.")
            return
        records = self._table.records()
        default_name = f"{date.today():%y%m%d}_{self._download_label(item)}.xlsx"
        save_path, _ = QFileDialog.getSaveFileName(self, "위하고 업로드용 파일 저장", default_name, "Excel Files (*.xlsx)")
        if not save_path:
            return
        try:
            export_wehago(item.result.profile, records, Path(save_path))
        except Exception as exc:
            QMessageBox.critical(self, "저장 오류", f"파일을 저장하는 중 오류가 발생했습니다:\n{exc}")
            return
        QMessageBox.information(self, "완료", f"저장되었습니다:\n{save_path}")

    def _export_items_to_folder(self, items: list[BatchItem], folder: str) -> tuple[list[str], list[str]]:
        saved, failed = [], []
        for item in items:
            name = f"{date.today():%y%m%d}_{self._download_label(item)}.xlsx"
            save_path = Path(folder) / name
            try:
                export_wehago(item.result.profile, item.result.records, save_path)
                saved.append(name)
            except Exception as exc:
                failed.append(f"{name}: {exc}")
        return saved, failed

    def _selected_rows(self) -> list[int]:
        return sorted({index.row() for index in self._queue_list.selectedIndexes()})

    def _selected_items(self) -> list[BatchItem]:
        ids = (self._item_id_for_row(row) for row in self._selected_rows())
        return [self._items[i] for i in ids if i is not None and i in self._items]

    def _on_download_selected(self) -> None:
        # 수당은 전부 인식됐지만 합계가 안 맞아 "확인 필요"로 표시되는 파일도
        # 계산 결과 자체는 있으므로 다운로드는 막지 않는다 — 상태 문자열이
        # 아니라 결과(item.result)가 있는지만 본다(사용자 요청, 2026-09-08).
        selected = [item for item in self._selected_items() if item.result is not None]
        if not selected:
            QMessageBox.information(
                self, "다운로드할 파일 없음", "왼쪽 목록에서 정상 처리된 파일을 먼저 선택하세요."
            )
            return
        folder = QFileDialog.getExistingDirectory(self, "저장할 폴더 선택")
        if not folder:
            return
        saved, failed = self._export_items_to_folder(selected, folder)
        message = f"{len(saved)}개 저장 완료."
        if failed:
            message += "\n실패:\n" + "\n".join(failed)
        QMessageBox.information(self, "선택 항목 다운로드", message)

    def _on_download_all(self) -> None:
        matched = [item for item in self._items.values() if item.result is not None]
        if not matched:
            QMessageBox.information(self, "다운로드할 파일 없음", "정상 처리된 파일이 없습니다.")
            return
        folder = QFileDialog.getExistingDirectory(self, "저장할 폴더 선택")
        if not folder:
            return
        saved, failed = self._export_items_to_folder(matched, folder)
        message = f"{len(saved)}개 저장 완료."
        if failed:
            message += "\n실패:\n" + "\n".join(failed)
        QMessageBox.information(self, "일괄 다운로드", message)

    # ---------- 사업장 관리 ----------

    def _on_manage_companies(self) -> None:
        dialog = CompanyManagerDialog(self)
        dialog.exec()

    # ---------- 환경설정 ----------

    def _on_open_settings(self) -> None:
        dialog = SettingsDialog(self)
        if dialog.exec():
            self._sync_wehago_btn.setToolTip(load_settings().wehago_form_folder)

    # ---------- 위하고 자동 업로드 ----------

    def _on_open_wehago_upload(
        self, initial_paths: list[Path] | None = None, auto_start: bool = False,
        expected_totals: dict[Path, int] | None = None,
    ) -> None:
        dialog = WehagoUploadDialog(
            self, initial_paths=initial_paths, auto_start=auto_start, expected_totals=expected_totals
        )
        dialog.exec()

    def _on_wehago_auto_upload(self) -> None:
        """"확정" 체크한 항목만 골라 설정된 폴더에 저장하고, 그 파일들로 바로
        업로드 자동화를 시작한다(사용자 요청, 2026-09-21) — 기존 "선택 항목
        다운로드"/"전체 다운로드"와는 별개의, 확정된 항목 전용 한 번에
        끝내는 버튼."""
        confirmed = [
            item
            for item_id, item in self._items.items()
            if item_id in self._confirmed_ids and item.result is not None
        ]
        if not confirmed:
            QMessageBox.information(
                self,
                "확정된 항목 없음",
                "큐 목록에서 업로드할 급여대장의 '확정' 체크박스를 먼저 켜세요.",
            )
            return

        settings = load_settings()
        folder = settings.upload_export_folder
        if not folder or not Path(folder).is_dir():
            folder = QFileDialog.getExistingDirectory(self, "위하고 업로드용 엑셀 저장 폴더 선택")
            if not folder:
                return
            settings.upload_export_folder = folder
            save_settings(settings)

        saved, failed = self._export_items_to_folder(confirmed, folder)
        if failed:
            QMessageBox.warning(self, "일부 저장 실패", "\n".join(failed))
        if not saved:
            return
        paths = [Path(folder) / name for name in saved]
        # 프로그램상의 차인지급액 합계 — 엑셀을 다시 열지 않고 이 값으로 대조
        totals: dict[Path, int] = {}
        for item in confirmed:
            name = f"{date.today():%y%m%d}_{self._download_label(item)}.xlsx"
            if name in saved:
                totals[Path(folder) / name] = int(round(sum(r.차인지급액 for r in item.result.records)))
        self._on_open_wehago_upload(initial_paths=paths, auto_start=True, expected_totals=totals)

    def _on_upload_saved_files(self) -> None:
        """이미 저장된 업로드용 파일로 바로 업로드 창을 연다(사용자 요청,
        2026-09-29 — 테스트할 때마다 같은 파일을 다시 저장하게 되는 문제).
        프로그램 계산 합계가 없으므로 기준 금액은 파일의 차인지급액 합계를
        읽어서 쓴다(`core/wehago_upload.py::_process_one`). 목록을 확인하고
        직접 "시작"을 누르도록 자동 시작은 하지 않는다."""
        paths = pick_saved_upload_files(self)
        if paths:
            self._on_open_wehago_upload(initial_paths=paths)

    # ---------- 위하고 양식 폴더 일괄 업데이트 ----------

    def _on_sync_wehago_folder(self) -> None:
        wehago_folder = Path(load_settings().wehago_form_folder)
        if not wehago_folder.is_dir():
            QMessageBox.warning(
                self, "폴더 없음", f"다음 폴더를 찾을 수 없습니다:\n{wehago_folder}\n환경설정에서 경로를 바꿀 수 있습니다."
            )
            return
        self._sync_wehago_btn.setEnabled(False)
        self._sync_wehago_btn.setText("업데이트 중…")

        start_wehago_folder_sync(self._result_queue, wehago_folder)

    def _on_wehago_folder_synced(self, result: SyncResult) -> None:
        self._sync_wehago_btn.setEnabled(True)
        self._sync_wehago_btn.setText("위하고 양식 폴더 업데이트")

        if result.folder_missing:
            QMessageBox.warning(
                self,
                "폴더 없음",
                f"다음 폴더를 찾을 수 없습니다:\n{load_settings().wehago_form_folder}",
            )
            return

        lines = [f"{len(result.registered)}개 사업장 등록/갱신 완료."]
        if result.skipped_files:
            lines.append(f"파일명에서 사업장명을 못 뽑아 건너뜀: {len(result.skipped_files)}개")
            lines.extend(f"  - {name}" for name in result.skipped_files[:10])
        if result.failed:
            lines.append(f"실패: {len(result.failed)}개")
            lines.extend(f"  - {name}: {err}" for name, err in result.failed[:10])
        QMessageBox.information(self, "위하고 양식 업데이트", "\n".join(lines))
