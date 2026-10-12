"""core/batch.py의 파일별 처리를 화면(UI) 스레드가 아니라 백그라운드에서
돌린다 — 여러 급여대장을 순서대로 열고 계산하는 게 파일당 수 초씩 걸릴 수
있어서, 화면 스레드에서 직접 돌리면 창이 "응답 없음"처럼 멈춘다.

**2026-09-08: QThread + Qt 시그널/슬롯에서 순수 파이썬 threading.Thread +
queue.Queue로 교체.** 예전엔 QObject 워커를 QThread로 옮겨서 완료를
`Signal`로 화면 스레드에 알렸는데, 반복적으로 재현되던 원인불명 네이티브
크래시(STATUS_STACK_BUFFER_OVERRUN, `ucrtbase.dll`의 정확히 같은 주소,
PySide6 6.6.3.1/6.10.3/6.11.2 세 버전 전부에서 동일하게 재현 — 버전
문제가 아님이 확인됨)의 크래시 스택이 매번 `shiboken6`/`Qt6Core.dll`의
크로스스레드 시그널 전달 경로를 지나고 있었다. 그 경로 자체를 아예 안
타도록, 백그라운드 작업은 이제 순수 `threading.Thread`로 돌리고 결과는
스레드 안전한 `queue.Queue`에 담아두기만 한다 — `ui/main_window.py`의
`QTimer`가 메인 스레드 안에서 주기적으로(같은 스레드 내 타이머라 Qt의
크로스스레드 큐잉을 안 씀) 큐를 비워서 화면을 갱신한다. `queue.Queue`는
파이썬 자체 락(threading.Lock/Condition) 기반이라 Qt/shiboken 코드를
전혀 거치지 않는다.

호출부(ui/main_window.py)는 각 `start_*` 함수가 돌려주는 결과를 기다릴
필요가 없다 — 결과는 전달받은 `result_queue`에 `(kind, payload)` 튜플로
쌓인다. daemon=True라 앱 종료 시 이 스레드들 때문에 프로세스가 안 끝나고
멈추는 일은 없다."""
from __future__ import annotations

import queue
import threading
from pathlib import Path

from core.batch import BatchItem, process_one, process_with_reference
from core.wehago_sync import DEFAULT_FOLDER, SyncResult, sync_from_folder

# result_queue에 쌓이는 (kind, payload) 중 kind 값:
#   "item_done"          — payload: (item_id: int, item: BatchItem)
#   "batch_add_finished" — payload: None (지금은 별도 처리 안 함, 훅 자리만)
#   "wehago_synced"      — payload: SyncResult


def start_batch_add(
    indexed_paths: list[tuple[int, Path, str | None, list[tuple[Path, str | None]] | None, bool]],
    result_queue: "queue.Queue[tuple[str, object]]",
) -> threading.Thread:
    """새로 추가된 파일들을 순서대로 자동 인식 + 처리한다. 파일 하나 끝날
    때마다 큐에 `("item_done", (item_id, item))`을 넣고, 전부 끝나면
    `("batch_add_finished", None)`을 넣는다. `merge_sources`(2개 이상)가
    있으면 그 (파일, 시트) 조합들을 합쳐서 이 항목 하나로 처리한다
    (추급1/추급2처럼 같은 종류로 나뉜 시트나 파일 — 사용자 요청,
    2026-09-08)."""

    def _run() -> None:
        for item_id, path, source_sheet, merge_sources, was_split in indexed_paths:
            item = process_one(
                path, source_sheet=source_sheet, was_split=was_split, merge_sources=merge_sources
            )
            result_queue.put(("item_done", (item_id, item)))
        result_queue.put(("batch_add_finished", None))

    thread = threading.Thread(target=_run, daemon=True, name="batch-add")
    thread.start()
    return thread


def start_single_run(
    item_id: int,
    path: Path,
    client_name: str,
    reference_path: Path,
    result_queue: "queue.Queue[tuple[str, object]]",
    source_sheet: str | None = None,
    reference_sheet: str | None = None,
    manual_overrides: dict | None = None,
    auto_register: bool = True,
    company_id: int | None = None,
    was_split: bool = False,
    merge_sources: list[tuple[Path, str | None]] | None = None,
) -> threading.Thread:
    """큐에 이미 있는 파일 하나를 (재)실행한다 — 위하고 파일을 수동으로
    지정했을 때, 또는 항목매핑결과 탭에서 수동 매칭을 적용해 다시 계산할
    때. 끝나면 큐에 `("item_done", (item_id, item))`을 넣는다(새로 추가된
    파일이 끝났을 때와 같은 종류라 main_window.py가 같은 핸들러로 처리)."""

    def _run() -> None:
        item: BatchItem = process_with_reference(
            path, client_name, reference_path,
            source_sheet=source_sheet, reference_sheet=reference_sheet,
            manual_overrides=manual_overrides, auto_register=auto_register,
            company_id=company_id, was_split=was_split, merge_sources=merge_sources,
        )
        result_queue.put(("item_done", (item_id, item)))

    thread = threading.Thread(target=_run, daemon=True, name="single-run")
    thread.start()
    return thread


def start_wehago_folder_sync(
    result_queue: "queue.Queue[tuple[str, object]]", folder: Path = DEFAULT_FOLDER
) -> threading.Thread:
    """`core/wehago_sync.py`로 위하고 양식 폴더 전체를 스캔/등록한다. 끝나면
    큐에 `("wehago_synced", result)`을 넣는다."""

    def _run() -> None:
        result: SyncResult = sync_from_folder(folder)
        result_queue.put(("wehago_synced", result))

    thread = threading.Thread(target=_run, daemon=True, name="wehago-sync")
    thread.start()
    return thread
