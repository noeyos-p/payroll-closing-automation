"""위하고 자동 업로드 — 이 앱(PayrollAutomation)이 만든 "위하고 업로드용"
엑셀 파일을 브라우저 자동화(Playwright)로 직접 위하고에 입력·저장한다.

`core/reconcile.py` 등 기존 엔진은 그대로 두고(파일 생성까지는 여전히
사람이 확인·확정한다), 그 다음 단계 — 확정된 파일을 위하고 화면에 직접
올리는 것만 자동화한다. 1.45(별도 배포 프로그램)와 달리 사원/수당·공제
매칭 판단은 전혀 하지 않는다 — 이미 완성된 파일을 위하고 자체의 "엑셀
불러오기" 기능에 그대로 먹이기만 한다.

## 실사로 확정한 전제 (2026-09-21, 실제 위하고 화면에서 직접 확인)

1. **동시 로그인 불가, 같은 세션 안 탭 추가는 가능.** 위하고는 계정당
   세션 1개만 허용한다 — 다른 브라우저에서 같은 계정으로 로그인하면
   기존 세션이 `#/login?type=expired`로 즉시 튕겨나가는 것을 실측
   확인했다. 반면 로그인된 세션 "안에서" 업체별 "급여" 버튼을 누르면
   `smarta.wehago.com`으로 향하는 **새 탭**이 뜨고, 그 탭의 컨텍스트
   (`cNum`/`taxNum`/`taxDate` 등)는 전부 URL 쿼리 파라미터에 담겨 있어
   탭마다 완전히 독립적이다(업체 A 탭에서 업체 B 탭으로 오가도 서로
   안 꼬이는 것 실측 확인). 그래서 이 모듈은 **브라우저 하나·로그인
   1번·업체 수만큼 탭**으로 동시 처리한다.

2. **급여자료입력 화면의 핵심 그리드는 RealGrid2(캔버스 렌더링)라
   접근성 트리/`innerText`로 전혀 안 읽힌다**(실측: `<canvas>` 대체
   텍스트조차 `document.body.innerText`에 안 잡힘). 그래서 이 모듈은:
   - **클릭**은 좌표 기반으로 한다 — 이 모듈이 항상 고정 뷰포트
     (`VIEWPORT`)로 브라우저를 띄우므로, 아래 좌표 상수들은 그 고정
     뷰포트를 전제로 실측한 값이다. 창 크기가 흔들리면 좌표가 전부
     어긋나므로 뷰포트를 절대 바꾸면 안 된다.
   - **읽기**는 좌표/OCR이 아니라 이 화면이 실제로 쓰는 그리드
     라이브러리 RealGrid2의 전역 JS API로 한다 — 어떤 그리드 영역을
     클릭해서 포커스를 준 뒤 `window.Grids.getActiveGrid()`로 그
     인스턴스를 얻고, `.getDataSource().getJsonRows()`로 실제 데이터
     행을 JSON 그대로 읽는다(캔버스는 화면 표시일 뿐, 내부 데이터
     모델은 평범한 JS 객체라 100% 정확하다 — 실측: 업체007
     8월 급+상 지급건에서 급여항목 그리드 합계 44,181,290 / 공제항목
     그리드 합계 4,495,380 / 화면에 표시된 차인지급액 39,685,910이
     그리드 데이터를 직접 합산한 값과 정확히 일치함을 확인).
     **OCR·스크린샷 인식은 전혀 안 쓴다.**

3. **"연말 입력된 사원"은 엑셀 업로드 시 자동 제외된다.** 사원코드연결
   확인창에 "연결되지 않은 사원 및 연말 입력된 사원은 제외하고
   변환됩니다"라는 문구가 항상 뜨는데, 몇 명이 왜 빠지는지는 이 팝업이
   알려주지 않는다. 그래서 이 모듈은 제외 사유를 추정하지 않는다 —
   변환 후 실제 차인지급액 합계가 엑셀 원본과 다르면 그냥 "확인 필요"로
   남기고 사람이 위하고 화면에서 직접 확인하게 한다(실측 사례: 중도퇴사
   사원의 "연말" 표시를 지우기 전엔 딱 그 사원 금액만큼 차이가 났고,
   지운 뒤 재실행하니 정확히 일치했다).

## 아직 실측 못 한 부분(주의)

- `_select_pay_type_and_open`이 지급구분 목록을 `getJsonRows()`로 읽을 때
  실제 필드명을 확인하지 못했다(실사 세션에서 그 달에 지급건이 언제나
  1개뿐이라 이 분기를 못 타봤다) — 그래서 특정 필드명에 기대지 않고
  각 행의 모든 값을 문자열로 합쳐 지급구분 텍스트가 포함되는지만 본다.
  지급건이 2개 이상인 달을 실제로 처음 자동 처리할 때 `logger`로 남는
  raw rows를 보고 이 함수를 다시 확인해야 한다.
- 공지 팝업 닫기(`_dismiss_known_popups`)는 이번에 실제로 뜬 것만
  처리한다 — 다른 업체에서 처음 보는 팝업이 뜨면 이 목록을 넓혀야 한다.
"""
from __future__ import annotations

import asyncio
import logging
import re
import threading
from dataclasses import dataclass
from pathlib import Path
from queue import Queue

logger = logging.getLogger(__name__)

VIEWPORT = {"width": 1600, "height": 900}
WEHAGO_MAIN_URL = "https://www.wehago.com/#/main"
SEARCH_PLACEHOLDER = "회사명, 사업자등록번호, 대표자명으로 검색하세요."

# --- 좌표 상수 ---
# 실사 확인(2026-09-21): 이 앱을 실제로 실행하는 PC마다 브라우저 뷰포트가
# 다르게 렌더링돼서(화면 해상도 등) 절대좌표가 구조적으로 안 맞는 사고를
# 여러 번 겪었다 — 8월을 노렸는데 1월이 선택되거나, 우측 지급구분을
# 노린 좌표가 완전히 좌측 귀속월 패널에 떨어지는 것까지 재현됨. 그래서
# **좌표 하드코딩에 의존하던 두 곳(귀속월 선택, 지급구분 선택)은 전부
# 없앴다** — 귀속월은 키보드 방향키(`_select_accrual_month`, 사용자가
# 직접 방향키로 실측 검증), 지급구분은 `document.querySelectorAll
# ('canvas')`로 화면에 실제 존재하는 그리드 위치를 그때그때 재서 찾는
# 방식(`_click_right_side_grid`)으로 교체됨. 아래 남은 좌표 상수들은
# 아직 이 문제를 안 겪어봤을 뿐 **같은 위험에 노출돼 있다** — 실패가
# 재현되면 같은 방식(런타임 측정 또는 키보드)으로 바꿔야 한다.
# "엑셀업로드" 미리보기 — 1행(헤더: 사원코드/사원명 등)
EXCEL_HEADER_ROW_XY = (446, 217)
# 급여자료입력 화면 우측 — 급여항목/공제항목 합계 그리드(각 아무 데이터 행)
ALLOWANCE_SUMMARY_XY = (1470, 300)
DEDUCTION_SUMMARY_XY = (1470, 830)

_GET_ACTIVE_GRID_ROWS_JS = """() => {
    const g = window.Grids && window.Grids.getActiveGrid && window.Grids.getActiveGrid();
    if (!g) return null;
    const ds = g.getDataSource();
    return ds.getJsonRows();
}"""


async def _press_escape(page) -> None:
    """Escape를 한 번이 아니라 **두 번** 누른다 — 사용자 방침(2026-09-21):
    "그냥 Esc를 한 번씩만 누르지 말고 2번씩 누르도록" (첫 번째가 안
    먹히거나, 중첩된 팝업이 하나 더 있는 경우를 대비). 아무것도 없는
    상태에서 두 번 눌러도 이 SPA에서 부작용이 없다는 전제(뒤로가기
    등으로 라우팅을 바꾸는 전역 핸들러가 없음, 지금까지 실사에서 확인된
    범위 내)."""
    await page.keyboard.press("Escape")
    await page.wait_for_timeout(150)
    await page.keyboard.press("Escape")


async def _viewport_scale(page) -> tuple[float, float]:
    """실측 좌표(전부 `VIEWPORT` 1600x900 기준)를 실제 이 페이지의 현재
    뷰포트 크기에 맞게 보정할 배율을 구한다.

    실사로 실제 겪은 문제(2026-09-21): 같은 회사·같은 화면인데도 실행할
    때마다 모달 안 요소들의 절대 좌표가 계속 달라졌다(귀속월 행 간격이
    28px 아닌 31px, 지급구분 헤더/데이터 간격이 예상과 다름, 심지어
    지급구분 더블클릭이 완전히 다른 패널(좌측 귀속월 목록)에 떨어짐) —
    단순 폰트 렌더링 오차치고는 너무 크고 방향성이 일관돼서, **실제
    뷰포트 크기 자체가 요청한 1600x900대로 안 나오고 있다**고 결론
    내렸다(이 앱을 실행하는 실제 PC의 화면 해상도가 작아서 `new_context
    (viewport=...)` 요청이 그대로 반영되지 않는 것으로 추정). 매 좌표
    클릭 전에 실제 `window.innerWidth`/`innerHeight`를 읽어 비율을
    곱하면, 뷰포트가 얼마나 작게 나오든 상대적 위치는 정확하다."""
    size = await page.evaluate("() => ({w: window.innerWidth, h: window.innerHeight})")
    return size["w"] / VIEWPORT["width"], size["h"] / VIEWPORT["height"]


async def _click_scaled(page, x: float, y: float, **kwargs) -> None:
    sx, sy = await _viewport_scale(page)
    await page.mouse.click(x * sx, y * sy, **kwargs)


async def _dblclick_scaled(page, x: float, y: float, **kwargs) -> None:
    sx, sy = await _viewport_scale(page)
    await page.mouse.dblclick(x * sx, y * sy, **kwargs)


@dataclass(frozen=True)
class UploadOutcome:
    path: Path
    company: str
    ok: bool
    message: str
    excel_total: int | None = None
    wehago_total: int | None = None


def _friendly_error(exc: Exception) -> str:
    """Playwright의 원본 예외 메시지는 "Target page, context or browser has
    been closed" + 브라우저 실행 커맨드라인 + call log 전체를 그대로 담고
    있어 회계사가 읽기엔 너무 길고 원인이 뭔지 알기 어렵다.

    실사 확인(2026-09-22): 새로 배포한 빌드로 처음 시도할 때, 동시 처리
    개수 기본값(5)으로 여러 업체를 한꺼번에 열자마자 떠 있던 Edge 창이
    갑자기 사라지며 이 오류가 났다 — 브라우저가 막 뜬 직후 탭을 여러 개
    동시에 여는 부하(또는 이 PC의 백신/보안 소프트웨어가 자동화 플래그가
    붙은 새 브라우저 프로세스를 감시하다 종료시키는 경우)로 인한 브라우저
    프로세스 자체의 비정상 종료로 추정된다 — 코드 로직 오류가 아니라
    환경 문제일 가능성이 높다. 원인을 코드로 확정할 수는 없으므로, 이
    메시지를 봤을 때 사용자가 취할 수 있는 구체적인 다음 행동(동시 처리
    개수를 낮춰서 재시도)을 안내한다."""
    text = str(exc)
    if "has been closed" in text:
        return (
            "브라우저가 처리 도중 예기치 않게 닫혔습니다 — 동시 처리 개수를 "
            "낮춰서(예: 1~2개) 다시 시도해보세요. 반복되면 이 PC의 백신/보안"
            f"소프트웨어가 자동화 브라우저를 종료시키고 있을 수 있습니다. (원본: {text.splitlines()[0]})"
        )
    return text


def read_excel_net_total(path: Path) -> int:
    """업로드할 파일 자체(이 앱이 만든 정리본)의 '차인지급액' 합계.

    `core/export.py`가 마지막 행에 "합계" 텍스트가 든 총계 행을 같이
    써두므로(A열이 사원코드가 아니라 "합계" 문자열), 그 행까지 사원
    행으로 잘못 더하면 합계가 정확히 2배로 부풀어 오른다(실사로 실제
    겪은 버그) — A열이 숫자로 시작하는 실제 사원 행만 더한다.

    `workbook.close()`를 반드시 호출한다 — 안 닫으면 파일 핸들이 열린 채로
    남아서, 같은 경로로 다시 저장(덮어쓰기)하려 할 때 윈도우에서 "다른
    프로세스가 사용 중"이라는 권한 오류가 난다(실사로 실제 겪은 버그,
    2026-09-21 — "일부 저장 실패"의 원인)."""
    import openpyxl

    workbook = openpyxl.load_workbook(path, data_only=True)
    try:
        worksheet = workbook.active
        headers = [cell.value for cell in worksheet[1]]
        try:
            net_pay_index = headers.index("차인지급액")
        except ValueError as exc:
            raise RuntimeError(f"'{path.name}'에 '차인지급액' 열이 없습니다.") from exc

        total = 0
        for row in worksheet.iter_rows(min_row=2, values_only=True):
            code = row[0]
            if code is None or str(code).strip() in ("", "합계"):
                continue
            value = row[net_pay_index]
            if isinstance(value, (int, float)):
                total += value
        return int(total)
    finally:
        workbook.close()


async def _get_grid_rows(page, click_x: int, click_y: int) -> list[dict]:
    """지정 좌표를 클릭해 그 위치의 RealGrid2 인스턴스를 활성화한 뒤,
    `getJsonRows()`로 실제 데이터 행을 그대로 읽는다(모듈 docstring의
    "실사로 확정한 전제" 2번 참고). 좌표는 `_click_scaled`로 실제
    뷰포트 크기에 맞게 보정한다(`_viewport_scale` 참고)."""
    await _click_scaled(page, click_x, click_y)
    await page.wait_for_timeout(150)
    rows = await page.evaluate(_GET_ACTIVE_GRID_ROWS_JS)
    return rows or []


def _sum_amount_fields(rows: list[dict]) -> int:
    total = 0
    for row in rows:
        for key in ("am_tax", "am_deductpay"):
            if key in row and isinstance(row[key], (int, float)):
                total += row[key]
                break
        else:
            for value in row.values():
                if isinstance(value, (int, float)):
                    total += value
                    break
    return int(total)


async def _dismiss_overlays(page, rounds: int = 8, settle_ms: int = 400, min_empty_rounds: int = 2) -> int:
    """위하고 사이트가 수시로 띄우는 안내/광고성 팝업을 전부 닫는다(2차
    인증 설정 안내, 요금제 광고, 공지사항 등) — 실측 확인: 한 번에
    최대 4개까지 서로 다른 팝업이 겹쳐 뜨기도 하고, 종류·개수는 접속할
    때마다 달라진다(사용자가 실제로 "급여자료입력 가서 창 뜨는 거
    두 개"라고 보고한 것도 이 케이스).

    **사용자 방침(2026-09-21): "웬만한 팝업창은 esc로 끄고, 안 되는
    것만 닫기 버튼을 누른다."** 이 UI 프레임워크의 모달 대부분이
    Escape로(저장 없이) 취소되는 걸 여러 화면에서 실측 확인했다(예:
    "근로소득 이동, 복사, 삭제" 모달의 "취소(esc)" 버튼, "수당 및
    공제등록" 스크린) — 그래서 매 라운드 Escape를 먼저 누르고, 그래도
    화면에 남아있는 것만 닫기 버튼으로 처리한다. 닫기 버튼은 문구가
    달라도 class 이름에 항상 "close"가 들어간다는 공통점이 있다(실측
    확인: `btn_close`/`close_btn`/대소문자 섞인 `Close` 등).

    팝업이 뒤늦게 뜨는 경우가 있어(실사로 확인한 배너 타이밍 버그) 매
    라운드마다 `settle_ms`만큼 대기한 뒤 다시 확인한다 — 화면에 하나도
    안 보이는 상태가 두 라운드 연속 나오면 더 없는 것으로 보고 끝낸다.
    실제로 닫기 버튼을 누른 개수를 돌려준다(호출부 로그/디버깅용, Escape
    로만 닫힌 것은 집계 안 함).

    `_search_and_open_payroll`(업체 검색 직전)과 `_open_payroll_entry`
    (급여자료입력 진입 직후) 두 군데에서 부른다 — 이 둘이 실제로 팝업이
    끼어들어 실패하던 지점이다."""
    closed = 0
    close_locator = page.locator('button[class*="close" i]:visible')
    empty_rounds = 0
    for _ in range(rounds):
        await _press_escape(page)
        await page.wait_for_timeout(settle_ms)
        try:
            count = await close_locator.count()
        except Exception:
            break
        if count == 0:
            empty_rounds += 1
            if empty_rounds >= min_empty_rounds:
                break
            continue
        empty_rounds = 0
        try:
            await close_locator.last.click(timeout=1500)
            closed += 1
        except Exception:
            break
    return closed


_KNOWN_INFO_POPUP_HEADINGS = (
    "국민연금 기준소득월액 상/하한액 변경 안내",
)


async def _dismiss_known_popups(page, attempts: int = 3) -> None:
    """급여자료입력 진입 직후 뜨는, **정확히 알려진 안내 팝업만** 닫는다
    (예: "국민연금 기준소득월액 상/하한액 변경 안내" — 이 팝업은 닫기
    버튼이 아니라 "확인" 버튼이라 `_dismiss_overlays`의 close-class
    패턴에 안 잡힌다).

    이전 구현은 이름이 "확인"인 버튼이면 뭐든 최대 5번까지 눌렀는데,
    이게 진짜 사고를 냈다 — 사용자가 실제로 겪은 화면: 이 정책 안내
    팝업을 닫은 뒤 자동화가 계속 "확인"을 눌러대다가 실제 데이터 설정
    화면인 **"수당 및 공제등록"**(비과세/감면 코드 설정)까지 열고, 그
    안의 "비과세 설정탭에 설정되지 않은 코드가 있습니다" 경고창까지
    진행시켜버렸다(스크린샷으로 확인, 2026-09-21) — 사업장의 실제
    수당/공제 설정을 건드릴 뻔한 위험한 상황이었다.

    그래서 이제는 **`_KNOWN_INFO_POPUP_HEADINGS`에 등록된 정확한 문구가
    화면에 보일 때만** 이 팝업을 닫는다 — 이 문구가 안 보이면 절대 아무
    버튼도 안 누르고 그냥 넘어간다(모르는 팝업을 추측해서 닫지 않는다는
    이 프로젝트의 기존 원칙과 동일). 사용자 방침(2026-09-21: "웬만한
    건 esc로") 대로 **Escape를 먼저 시도**하고, 그래도 안 닫히면(문구가
    여전히 보이면) "확인" 버튼을 누른다. 이후 단계(사원코드 연결/변환/
    삭제 경고)의 "확인" 버튼은 각 단계 함수가 직접, 정확한 타이밍에
    누른다 — 이 함수를 그 뒤에 또 부르면 안 된다."""
    for _ in range(attempts):
        closed_any = False
        for heading_text in _KNOWN_INFO_POPUP_HEADINGS:
            heading = page.get_by_text(heading_text, exact=False)
            if await heading.count() == 0:
                continue
            await _press_escape(page)
            await page.wait_for_timeout(300)
            if await heading.count() == 0:
                closed_any = True
                continue
            try:
                await page.get_by_role("button", name="확인", exact=True).first.click(timeout=1500)
                closed_any = True
                await page.wait_for_timeout(300)
            except Exception:
                pass
        if not closed_any:
            return


async def _ensure_t_edge_unused_tab(main_page) -> None:
    """"T edge 미사용" 탭이 실제로 선택된 상태인지 확인하고, 아니면
    눌러서 전환한다 — 실측 확인: 이 탭이 선택되면 그 탭의 `<li>`에
    `class="active"`가 붙는다(버튼 자신의 class는 선택 여부와 무관하게
    항상 동일). 그냥 클릭만 하고 넘어가면(이전 구현) 팝업에 막혀 클릭이
    실제로는 안 먹혔는데도 모르고 지나가서, 엉뚱한 탭(예: "T edge 사용")
    에서 검색해 "검색 결과가 없습니다"로 실패하는 걸 실사로 겪었다(사용자
    보고, 2026-09-21: "T edge 무조건 미사용으로 하도록 해야해"). 그래서
    클릭 후 `<li class="active">` 여부를 직접 확인하고, 아니면 팝업을
    다시 닫고 재시도한다 — 몇 번을 시도해도 안 되면 추측하지 말고
    에러로 정직하게 실패한다."""
    tab_button = main_page.get_by_role("button", name="T edge 미사용")
    tab_li = tab_button.locator("xpath=ancestor::li[1]")
    for _ in range(4):
        class_attr = await tab_li.get_attribute("class") or ""
        if "active" in class_attr.split():
            return
        await _dismiss_overlays(main_page)
        await tab_button.click(force=True)
        await main_page.wait_for_timeout(500)
    raise RuntimeError("'T edge 미사용' 탭으로 전환하지 못했습니다(팝업이 계속 가로막고 있을 수 있습니다).")


async def _search_and_open_payroll(main_page, company_query: str):
    """메인 화면에서 "T edge 미사용" 탭으로 전환 → 회사명 검색 → 같은
    이름이 여러 개면 괄호 안 연도가 가장 최신인 행의 '급여' 버튼을 눌러
    새 탭을 연다(사용자 확인, 2026-09-21: "업체007(2026)"처럼
    괄호 부분이 최신인 업체로 들어가면 됨).

    **2026-09-22 수정**: Enter 입력 뒤 고정 800ms만 기다리고 바로 결과
    개수를 확인했는데, 사용자가 "업체071"로 검색했을 때
    두 번 중 한 번은 실제로 결과가 있는데도 "검색 결과가 없습니다"로
    실패하는 걸 겪었다(같은 검색이 매번 똑같이 실패/성공하지 않고
    들쭉날쭉함) — 렌더링이 800ms보다 늦게 끝나는 경우가 있다는 뜻의
    전형적인 타이밍 문제(이 프로젝트가 배너/팝업 타이밍에서 이미 여러 번
    겪은 것과 같은 패턴). 고정 대기 한 번 대신, 결과가 실제로 나타날
    때까지 간격을 늘려가며 여러 번 확인한다."""
    await _ensure_t_edge_unused_tab(main_page)
    search_box = main_page.get_by_placeholder(SEARCH_PLACEHOLDER)
    await search_box.fill(company_query)
    await search_box.press("Enter")

    rows = main_page.locator("a", has_text=company_query)
    count = 0
    for wait_ms in (400, 400, 600, 800, 1000):
        await main_page.wait_for_timeout(wait_ms)
        count = await rows.count()
        if count > 0:
            break
    if count == 0:
        raise RuntimeError(f"'{company_query}' 검색 결과가 없습니다.")

    year_re = re.compile(r"\((\d{4})\)\s*$")
    best_index, best_year = 0, -1
    for i in range(count):
        text = (await rows.nth(i).inner_text()).strip()
        match = year_re.search(text)
        year = int(match.group(1)) if match else 0
        if year > best_year:
            best_index, best_year = i, year

    li = rows.nth(best_index).locator("xpath=ancestor::li[1]")
    async with main_page.expect_popup() as popup_info:
        await li.get_by_role("button", name="급여", exact=True).click()
    new_page = await popup_info.value
    await new_page.wait_for_load_state()
    return new_page


async def _dismiss_stray_settings_screen(page, attempts: int = 3) -> None:
    """이 회사 화면에 "수당 및 공제등록"(비과세/감면·사회보험 등 실제
    설정값을 담은 화면) 이 이미 열린 채로 나타나는 경우를 닫는다 —
    실측으로 같은 업체(업체007)에서 두 번 연속 정확히 같은 탭
    ("비과세/감면설정")이 열린 채로 나타났다(사용자 스크린샷, 2026-09-21
    두 차례) — 위하고가 업체별로 마지막에 보던 화면을 기억해뒀다가
    다시 열 때 복원하는 것으로 추정된다(이날 실사 중 이 화면을 실수로
    열어봤던 이력이 있다). 이 화면은 실제 수당/공제/비과세 설정값을
    담고 있어 어떤 버튼도 누르면 안 되므로, **Escape 키로만** 닫는다
    (이 UI 프레임워크가 "취소(esc)"를 지원하는 걸 다른 모달에서 이미
    확인함 — "확인"/"취소" 등 어떤 버튼도 클릭하지 않는다)."""
    for _ in range(attempts):
        heading = page.get_by_text("수당 및 공제등록", exact=True)
        if await heading.count() == 0:
            return
        await _press_escape(page)
        await page.wait_for_timeout(400)


async def _open_payroll_entry(page) -> None:
    """급여자료입력에 들어간 뒤, 뜰 수 있는 팝업들을 처리하기 전에
    **충분히 기다린다** — 사용자 지적(2026-09-21): "바로바로 하니까
    렉먹어서 그런 거 같은데, 팝업창 뜰 때까지 조금 기다렸다가 완벽히
    끄고 가도록 해줘". 이 PC가 느릴 때 팝업이 늦게 뜨는 걸 이미
    한 번(2차 인증 배너) 실측으로 겪었는데, 급여자료입력 진입 시에도
    같은 종류의 지연이 있을 수 있다고 보고 초기 대기를 넉넉히 늘리고,
    `_dismiss_overlays`도 이 지점에서만 더 오래(더 많은 라운드·더 긴
    간격) 확인하도록 한다."""
    await page.get_by_role("link", name="급여자료입력", exact=True).click()
    await page.wait_for_load_state()
    await page.wait_for_timeout(2000)
    await _dismiss_stray_settings_screen(page)
    await _dismiss_overlays(page, rounds=10, settle_ms=600, min_empty_rounds=3)
    await _dismiss_known_popups(page)


async def _select_accrual_month(page, month: int) -> None:
    """좌표 클릭 대신 키보드 이동으로 귀속월을 고른다 — 실사로 실제
    사고를 겪었다(2026-09-21: 8월을 노렸는데 렌더링 차이로 좌표가 안
    맞아 1월이 선택된 채로 멈춤). "근로소득 이동, 복사, 삭제" 팝업은
    열리면 "01"행이 기본 선택돼 있고, 방향키로 정확히 한 행씩 이동한다
    (사용자가 직접 방향키로 실측 확인) — 픽셀이 얼마나 다르게 렌더링되든
    상관없다.

    **2026-09-22 수정**: "항상 01행이 기본 선택"이라는 가정 자체가 깨질 수
    있다는 게 사용자 신고로 드러났다 — "같은 업체도 창을 두 개 띄워서
    다르게 처리한다"(예: 09(급여)/08(급+상) 두 파일이 같은 업체인데 한쪽
    귀속월이 엉뚱하게 들어감). 이 UI 프레임워크가 업체별로 "마지막에 보던
    화면"을 서버에 기억해뒀다가 복원하는 동작이 이미 다른 화면("수당 및
    공제등록", `_dismiss_stray_settings_screen` 참고)에서 확인된 적이
    있다 — 같은 업체를 여러 창에서 동시에 처리하면, 한 창이 이미 다른
    달로 이동해둔 상태가 나중에 여는 창의 모달 기본 선택 행에 영향을 줄
    가능성이 높다(계정/업체 단위로 서버가 마지막 선택을 기억하는 것으로
    보임). 그래서 이제 몇 번째 행에서 시작하든 상관없이 **먼저 ArrowUp을
    넉넉히 눌러 맨 위(01행)로 강제 이동**한 뒤 목표 월까지 내려간다 —
    그리드는 첫 행에서 더 눌러도 그대로 멈춰 있을 뿐이라 안전하다."""
    await page.get_by_role("button", name="지급일자", exact=True).click()
    # 변환 뒤 미저장 상태에서는 지급일자 대신 저장 확인창이 먼저 뜬다.
    for _ in range(30):
        save_prompt = page.get_by_text(re.compile(r"저장 후.*지급일자.*실행"))
        if await save_prompt.count() and await save_prompt.first.is_visible():
            await page.get_by_role("button", name="확인", exact=True).last.click()
        if await page.evaluate(_FIND_MODAL_CANVASES_JS):
            break
        await page.wait_for_timeout(300)
    else:
        raise RuntimeError("지급일자 목록이 열리지 않았습니다. 저장 확인창 또는 안내창을 확인하세요.")
    for _ in range(12):
        await page.keyboard.press("ArrowUp")
        await page.wait_for_timeout(60)
    for _ in range(month - 1):
        await page.keyboard.press("ArrowDown")
        await page.wait_for_timeout(120)
    await page.wait_for_timeout(300)


_FIND_MODAL_CANVASES_JS = """() => {
    // "근로소득 이동, 복사, 삭제" 팝업 안에 있는 canvas만 찾는다 — 페이지
    // 전체에서 찾으면 배경에 깔린 다른 그리드(사원목록/급여항목/공제항목
    // 등)까지 후보로 잡혀서 엉뚱한 걸 클릭하는 사고가 실제로 났다
    // (2026-09-21, 사용자 지적: "애초에 오른쪽으로 가질 않아").
    const heading = [...document.querySelectorAll('*')].find(
        el => el.textContent.trim() === '근로소득 이동, 복사, 삭제'
            && (!el.firstElementChild || el.firstElementChild.textContent.trim() !== '근로소득 이동, 복사, 삭제')
    );
    if (!heading) return [];
    let container = heading;
    for (let i = 0; i < 12 && container; i++) {
        if (container.querySelectorAll('canvas').length >= 2) break;
        container = container.parentElement;
    }
    if (!container) return [];
    const canvases = [...container.querySelectorAll('canvas')].filter(c => {
        const r = c.getBoundingClientRect();
        return r.width > 60 && r.height > 60;
    });
    canvases.sort((a, b) => a.getBoundingClientRect().left - b.getBoundingClientRect().left);
    return canvases.map(c => {
        const r = c.getBoundingClientRect();
        return {left: r.left, top: r.top, width: r.width, height: r.height};
    });
}"""


async def _click_right_side_grid(page) -> list[dict]:
    """"근로소득 이동, 복사, 삭제" 팝업의 우측(지급구분) 그리드를 좌표
    추정이 아니라 **화면에 실제로 존재하는 canvas 요소의 위치를 그때그때
    재서** 찾아 클릭하고, 그 결과로 읽힌 행을 돌려준다.

    실사로 실측 좌표 접근을 여러 번 고쳐도 계속 실패했다(2026-09-21) —
    귀속월 목록 줄 간격이 실행마다 다르고, 심지어 우측을 노린 좌표가
    완전히 좌측 패널에 떨어지기도 했다(사용자 스크린샷 두 차례). 이
    앱을 실행하는 PC마다 실제 화면/뷰포트 크기가 달라 절대 좌표가
    구조적으로 안 맞는 것으로 결론짓고, **좌표 하드코딩을 완전히
    포기**했다 — 대신 `_FIND_MODAL_CANVASES_JS`로 "근로소득 이동, 복사,
    삭제" 팝업 **안에 있는** 그리드들의 실제 위치만 읽어(페이지 전체를
    뒤지면 배경에 깔린 다른 그리드까지 걸려 엉뚱한 걸 클릭하는 사고가
    실제로 났다 — 사용자 지적: "애초에 오른쪽으로 가질 않아"), 가장
    오른쪽에 있는 것부터 클릭해보고 `getJsonRows()`로 읽은 행이 좌측
    귀속월 목록(항상 정확히 12행)이 아니면(=우측 지급구분 목록을
    제대로 찾았으면) 그걸 쓴다.

    **주의: 여기서는 절대 Escape를 누르면 안 된다.** 이전 구현은 "혹시
    남아있을 팝업부터 치운다"며 맨 앞에서 `_press_escape`를 불렀는데,
    Claude가 직접 라이브로 재현해서 확정한 진짜 원인이 바로 이거였다
    (2026-09-21) — Escape가 다른 팝업이 아니라 **지금 막 연 "근로소득
    이동, 복사, 삭제" 모달 자체를 "취소(esc)"로 닫아버리고** 있었다.
    그래서 그 뒤에 캔버스를 아무리 찾아도 모달이 이미 사라진 뒤라 항상
    0개였다 — 지금까지의 "타이밍 문제"라는 진단도 틀렸었다(재시도
    로직을 넣어도 매번 실패했던 이유). 이 함수에 들어올 시점엔 이미
    `_select_accrual_month`가 그 모달을 열어 원하는 월까지 이동해둔
    상태이므로, 여기서 팝업을 치울 필요 자체가 없다.

    모달이 아직 다 렌더링되기 전에 확인하면 후보가 0개로 나올 수 있어
    (렌더링 지연 자체는 여전히 가능성으로 남겨둠), 0개면 점점 더 오래
    기다리며 몇 번 더 확인한다."""
    rects: list[dict] = []
    for wait_ms in (0, 500, 1000, 1500):
        if wait_ms:
            await page.wait_for_timeout(wait_ms)
        rects = await page.evaluate(_FIND_MODAL_CANVASES_JS)
        if rects:
            break
    logger.warning("_click_right_side_grid: 모달 안 캔버스 후보=%r", rects)
    for rect in reversed(rects):  # 오른쪽부터 시도
        x = rect["left"] + rect["width"] / 2
        y = rect["top"] + min(40, rect["height"] / 4)
        await page.mouse.click(x, y)
        await page.wait_for_timeout(200)
        rows = await page.evaluate(_GET_ACTIVE_GRID_ROWS_JS) or []
        logger.warning("_click_right_side_grid: (%.0f,%.0f) 클릭 → %d행", x, y, len(rows))
        if rows and all("fg_pay" in row and "dt_pay" in row for row in rows):
            return rows
    raise RuntimeError(
        f"지급구분 목록 그리드를 화면에서 찾지 못했습니다(모달 안 캔버스 후보 {len(rects)}개 시도함)."
    )


def _classify_pay_type_text(text: str) -> str | None:
    """숫자 지급구분 코드가 없는 경우에만 사용하는 텍스트 정규화."""
    normalized = text.replace(" ", "")
    if normalized in ("급+상", "급상"):
        return "급+상"
    if "상여" in normalized:
        return "급+상"
    if "추급" in normalized or "추가" in normalized:
        return "추급"
    if "급여" in normalized:
        return "급여"
    return None


def _find_pay_type_row_index(
    rows: list[dict], pay_type: str | None, expected_total: int | None = None
) -> int | None:
    """지급구분 코드로 유일한 행을 찾는다. 기존 금액은 선택 기준으로 쓰지 않는다.

    파일에 종류가 없으면 지급건이 하나인 경우만 선택한다. expected_total은
    기존 호출부 호환용이며 업로드 결과 검증은 _process_one에서 수행한다.
    """
    if not rows:
        return None
    if pay_type:
        # 지급구분은 금액이 아니라 코드로 식별한다. 업로드 전 금액은 달라질 수 있다.
        codes = {"1": "급여", "2": "급+상", "5": "추급"}
        hits = []
        for i, row in enumerate(rows):
            code = str(row.get("fg_pay", "")).strip()
            kind = codes.get(code)
            if not code:
                kind = _classify_pay_type_text(" ".join(str(v) for v in row.values() if v is not None))
            if kind == pay_type:
                hits.append(i)
        return hits[0] if len(hits) == 1 else None
    # 종류가 없는 파일은 단일 지급건만 자동 선택한다.
    return 0 if len(rows) == 1 else None



async def _select_pay_type_and_open(page, pay_type: str | None, expected_total: int | None = None) -> None:
    rows = await _click_right_side_grid(page)
    if not rows:
        raise RuntimeError(
            "지급일자 목록에 해당 귀속월 지급건이 없습니다 — 전월 지급건 복사가 먼저 되어 있어야 합니다."
        )

    target_index = _find_pay_type_row_index(rows, pay_type, expected_total)
    if target_index is None:
        logger.warning(
            "지급구분 '%s'(기대 차인지급액=%s) 매칭 실패, raw rows=%r", pay_type, expected_total, rows
        )
        raise RuntimeError(
            f"지급구분 '{pay_type}'과 일치하는 항목을 찾지 못했습니다(지급구분이 없거나 같은 종류가 여러 건입니다)."
        )

    if str(rows[target_index].get("yn_close", "0")) == "1":
        raise RuntimeError("선택한 지급건이 마감되어 있습니다. 마감 상태를 확인한 후 다시 시도하세요.")

    # `_click_right_side_grid`가 이미 그 그리드를 클릭해서 포커스와 함께
    # 첫 행이 선택된 상태다(같은 RealGrid2 컴포넌트라 좌측 귀속월 목록과
    # 같은 방식으로 동작) — 목표 행까지 ArrowDown으로 이동한 뒤, 실측
    # 확인된 실제 **"실행" 버튼**을 누른다(사용자 확인, 2026-09-21: 행을
    # 체크하면 하단 툴바에 "실행" 버튼이 실제로 나타남 — 화면 안내 문구
    # "더블클릭 또는 실행하시면 해당 급여자료를 불러옵니다"의 "실행"이
    # 바로 이 버튼이었다). Enter 키보다 이 버튼을 직접 누르는 게 더
    # 확실하고, 몇 번째 행이든 방향키로 사원014 하면 항상 같은 방식으로
    # 처리된다.
    for _ in range(len(rows)):
        await page.keyboard.press("ArrowUp")
    for _ in range(target_index):
        await page.keyboard.press("ArrowDown")
        await page.wait_for_timeout(120)
    await page.get_by_role("button", name="실행", exact=True).click()
    await page.wait_for_load_state()
    await page.wait_for_timeout(500)


async def _upload_excel(page, file_path: Path, ui_lock: "asyncio.Lock") -> None:
    """더보기 아이콘 좌표를 클릭해 메뉴를 연 뒤 "엑셀 불러오기" 텍스트를
    누르는 대신, 그 메뉴 항목에 적혀 있던 단축키 **Ctrl+U**를 바로
    쓴다(사용자 확인, 2026-09-21) — 좌표 의존을 하나 더 없앤다."""
    # Ctrl+U는 화면 맨 앞 탭에서만 확실히 동작하므로, 이 키 입력과 파일 선택창
    # 처리까지만 한 탭씩 순서대로 — 그 탭을 앞으로 가져와 처리한다.
    async with ui_lock:
        await page.bring_to_front()
        await page.wait_for_timeout(500)
        async with page.expect_file_chooser(timeout=15000) as fc_info:
            await page.keyboard.press("Control+u")
        file_chooser = await fc_info.value
        await file_chooser.set_files(str(file_path))
        await page.wait_for_timeout(800)

    await _click_scaled(page, *EXCEL_HEADER_ROW_XY)
    await page.wait_for_timeout(300)
    await page.get_by_text("엑셀제목설정", exact=False).click()
    await page.wait_for_timeout(500)
    await page.get_by_role("button", name="확인", exact=True).last.click()
    await page.wait_for_timeout(800)


async def _confirm_and_convert(page) -> None:
    await page.get_by_role("button", name="변환", exact=True).click()
    await page.wait_for_timeout(500)
    # "연결되지 않은 사원 및 연말 입력된 사원은 제외하고 변환됩니다" 확인
    await page.get_by_role("button", name="확인", exact=True).last.click()
    await page.wait_for_timeout(500)
    # "현재 지급일자의 급여데이터를 삭제후 업로드됩니다" 경고 — 기존
    # 데이터가 없는 지급건에서도 항상 뜨는지 미확인이라, 안 뜨면 그냥
    # 넘어간다(예외를 삼키고 진행).
    try:
        await page.get_by_role("button", name="확인", exact=True).first.click(timeout=3000)
    except Exception:
        pass
    await page.wait_for_timeout(1000)


async def _read_wehago_net_total(page, month: int, pay_type: str | None) -> int:
    """저장 후 지급일자 목록을 다시 조회해 해당 지급건의 차인지급액을 읽는다.

    요약 패널의 고정 좌표는 다른 그리드를 활성화할 수 있으므로 쓰지 않는다.
    지급구분을 재확인하고 숫자 필드가 없으면 실패로 처리한다.
    """
    await _select_accrual_month(page, month)
    rows = await _click_right_side_grid(page)
    index = _find_pay_type_row_index(rows, pay_type, None)
    if index is None:
        raise RuntimeError("저장 결과의 지급건을 유일하게 확인하지 못했습니다.")
    amount = rows[index].get("am_orpay")
    if isinstance(amount, bool) or not isinstance(amount, (int, float)):
        raise RuntimeError("저장 결과에 차인지급액 값이 없습니다.")
    return int(amount)


async def _process_one(
    main_page,
    info,
    semaphore: "asyncio.Semaphore",
    search_lock: "asyncio.Lock",
    ui_lock: "asyncio.Lock",
    expected_total: int | None = None,
) -> UploadOutcome:
    async with semaphore:
        # 기준 합계는 프로그램이 계산한 차인지급액 합계를 우선 쓴다(사용자
        # 요청, 2026-09-22: 엑셀을 열어 확인하기 어려우면 프로그램상의 합계를
        # 보는 게 낫다 — 엑셀이 열려 있어도 영향 없음). 없을 때만 파일에서 읽는다.
        if expected_total is not None:
            excel_total = expected_total
        else:
            try:
                excel_total = read_excel_net_total(info.path)
            except Exception as exc:
                return UploadOutcome(info.path, info.company, False, f"엑셀 읽기 실패: {exc}")

        page = None
        stage = "업체 검색"
        try:
            # 검색창은 메인 탭 하나를 공유하므로, 동시에 여러 업체를
            # 검색하면 서로 입력을 덮어쓴다 — 이 부분만 잠가서 순서대로
            # 처리하고, 새 탭이 열린 뒤부터는 완전히 독립적이라 그대로
            # 동시에 진행한다.
            async with search_lock:
                page = await _search_and_open_payroll(main_page, info.company)
            await page.set_viewport_size(VIEWPORT)
            stage = "급여자료입력 열기"
            logger.warning("업로드 단계: %s / %s", info.path.name, stage)
            await _open_payroll_entry(page)
            # 사용자 확인(2026-09-22): 지급일자 선택 등은 동시에 해도 되고, 뒤에
            # 있는 탭에서 Ctrl+U(엑셀 불러오기)만 안 먹는다 — `_upload_excel`
            # 안에서 그 부분만 `ui_lock`으로 한 탭씩 처리한다.
            stage = "귀속월 선택"
            logger.warning("업로드 단계: %s / %s", info.path.name, stage)
            await _select_accrual_month(page, info.month)
            # 이미 위에서 구한 excel_total(프로그램이 계산한 차인지급액)을
            # 그대로 넘겨 지급구분 행을 금액으로 특정한다(2026-09-22 실사
            # 확정 — `_select_pay_type_and_open` docstring 참고).
            stage = "지급구분 선택"
            logger.warning("업로드 단계: %s / %s", info.path.name, stage)
            await _select_pay_type_and_open(page, info.pay_type, excel_total)
            stage = "엑셀 불러오기"
            logger.warning("업로드 단계: %s / %s", info.path.name, stage)
            await _upload_excel(page, info.path, ui_lock)
            stage = "변환 및 저장"
            logger.warning("업로드 단계: %s / %s", info.path.name, stage)
            await _confirm_and_convert(page)
            stage = "저장 결과 검증"
            logger.warning("업로드 단계: %s / %s", info.path.name, stage)
            wehago_total = await _read_wehago_net_total(page, info.month, info.pay_type)
        except Exception as exc:
            # 실패/예외 탭은 그대로 열어둔다 — 처리 끝에 자동으로 닫아버리면
            # 정확히 무엇 때문에 실패했는지 화면에서 확인할 수 없다는 실사
            # 피드백(2026-09-21). 사용자가 다 보고 나서 직접 닫으면 된다.
            return UploadOutcome(info.path, info.company, False, f"{stage} 중 오류: {_friendly_error(exc)}", excel_total)

        if wehago_total == excel_total:
            # 2026-09-22 사용자 요청: 금액까지 정확히 맞아 진짜로 성공한
            # 탭은 자동으로 닫는다 — 여러 업체를 동시에 돌리면 탭/브라우저
            # 프로세스가 계속 쌓여서 시스템이 느려진다는 실사 피드백(위
            # 렉 진단 대화 참고). 실패/불일치 탭만 계속 열어둬서 위
            # "실패 원인을 확인할 수 있어야 한다"는 원래 요구사항은 그대로
            # 지킨다 — 닫을지 말지를 성공 여부로만 가른다.
            try:
                await page.close()
            except Exception:
                pass  # 탭이 이미 닫혀있는 등 — 닫기 실패는 성공 결과 자체에 영향 없음
            return UploadOutcome(info.path, info.company, True, "성공", excel_total, wehago_total)
        return UploadOutcome(
            info.path,
            info.company,
            False,
            f"금액 불일치(프로그램 {excel_total:,} vs 위하고 {wehago_total:,}) — "
            "연말/퇴사 처리된 사원 자동 제외 등 원인을 위하고 화면에서 직접 확인하세요.",
            excel_total,
            wehago_total,
        )


async def _launch_browser(playwright):
    """설치된 Chrome을 우선 사용하고, 없으면 Chromium/Edge 순서로 시도한다."""
    launch_args = ["--disable-notifications"]
    errors: list[str] = []
    for channel in ("chrome", None, "msedge"):
        try:
            if channel is None:
                return await playwright.chromium.launch(headless=False, args=launch_args)
            return await playwright.chromium.launch(headless=False, args=launch_args, channel=channel)
        except Exception as exc:  # noqa: BLE001 - 다음 후보로 폴백
            errors.append(f"{channel or 'chromium'}: {str(exc).splitlines()[0]}")
    raise RuntimeError("실행 가능한 브라우저(Chromium/Edge/Chrome)를 찾지 못했습니다. " + " | ".join(errors))


async def _log_in(page, timeout_ms: int, wehago_id: str = "", wehago_password: str = "") -> None:
    """아이디/비밀번호가 둘 다 설정돼 있으면 로그인 폼을 직접 채워서
    자동 로그인하고, 없으면(기본값) 지금까지처럼 사람이 직접 로그인할
    때까지 기다린다(사용자 요청, 2026-09-21 — 환경설정에 저장된 값 사용,
    평문 저장은 `core/settings.py` 참고). 로그인 폼 필드는 실측
    확인(`https://www.wehago.com/#/login`): 아이디 입력칸 placeholder
    "아이디를 입력하세요.", 비밀번호 칸 "비밀번호를 입력하세요.", 버튼
    이름 "로그인"."""
    await page.goto(WEHAGO_MAIN_URL)
    if wehago_id and wehago_password:
        try:
            await page.get_by_placeholder("아이디를 입력하세요.").fill(wehago_id, timeout=5000)
            await page.get_by_placeholder("비밀번호를 입력하세요.").fill(wehago_password, timeout=5000)
            await page.get_by_role("button", name="로그인", exact=True).click(timeout=5000)
            await page.wait_for_timeout(800)
            # 중복 로그인(다른 브라우저/기기에서 이미 로그인 중) 확인창 —
            # 사용자 방침(2026-09-21): "중복 로그인은 무조건 확인 누르기"
            # — 기존 세션을 밀어내고 이 자동화 세션으로 로그인을 계속
            # 진행한다. 안 뜨면 그냥 넘어간다(예외를 삼킴).
            try:
                await page.get_by_role("button", name="확인", exact=True).first.click(timeout=3000)
            except Exception:
                pass
        except Exception:
            pass  # 이미 로그인돼 있거나(리다이렉트로 로그인 폼 자체가 없음) 폼 구조가 바뀐 경우 — 아래 대기로 폴백
    await page.wait_for_selector("text=담당 수임처", timeout=timeout_ms)


class _Session:
    """앱이 켜져 있는 동안 유지되는 위하고 브라우저 세션. 업로드 버튼을 누를
    때마다 브라우저를 새로 띄우고 다시 로그인하던 것을(사용자 요청,
    2026-09-22) — 이미 열려 있는 메인 화면(https://www.wehago.com/#/main,
    "T edge 미사용" 검색창이 보이는 상태)이 있으면 로그인·팝업 처리를 전부
    건너뛰고 그 탭을 그대로 재사용하도록 바꿨다. 브라우저를 처리 후에도
    계속 열어두려는 요구("절대 자동으로 안 닫는다")도 이 구조가 자연스럽게
    만족한다 — 전용 이벤트 루프 스레드가 앱 종료 때까지 살아 있다."""

    lock = threading.Lock()
    loop: "asyncio.AbstractEventLoop | None" = None
    thread: "threading.Thread | None" = None
    setup_lock: "asyncio.Lock | None" = None
    search_lock: "asyncio.Lock | None" = None
    ui_lock: "asyncio.Lock | None" = None
    playwright = None
    browser = None
    main_page = None


def _ensure_worker_loop() -> "asyncio.AbstractEventLoop":
    with _Session.lock:
        if _Session.loop is None or _Session.thread is None or not _Session.thread.is_alive():
            loop = asyncio.new_event_loop()

            def _serve() -> None:
                asyncio.set_event_loop(loop)
                loop.run_forever()

            thread = threading.Thread(target=_serve, daemon=True, name="wehago-upload")
            thread.start()
            _Session.loop, _Session.thread = loop, thread
            _Session.setup_lock = asyncio.Lock()
            _Session.search_lock = asyncio.Lock()  # 메인 탭 검색창은 세션 전체가 공유
            _Session.ui_lock = asyncio.Lock()
            _Session.playwright = _Session.browser = _Session.main_page = None
        return _Session.loop


def _forget_browser(*_args) -> None:
    _Session.browser = None
    _Session.main_page = None


async def _main_page_is_ready(page) -> bool:
    """이미 열려 있는 탭이 로그인된 메인 화면(검색창이 보이는 상태)인지."""
    try:
        if page.is_closed() or not page.url.startswith(WEHAGO_MAIN_URL):
            return False
        return await page.get_by_placeholder(SEARCH_PLACEHOLDER).is_visible()
    except Exception:
        return False


async def _get_main_page(login_timeout_ms: int, wehago_id: str, wehago_password: str):
    """세션의 메인 탭을 돌려준다. 열려 있는 로그인된 메인 화면이 있으면 로그인과
    팝업 처리 없이 그대로 재사용하고, 브라우저는 있는데 메인 화면이 아니면 그
    탭에서 로그인만 다시 하며, 브라우저가 없거나 닫혔으면 새로 띄운다."""
    browser, page = _Session.browser, _Session.main_page
    if browser is not None and browser.is_connected() and page is not None and not page.is_closed():
        if await _main_page_is_ready(page):
            logger.warning("기존 위하고 메인 화면 재사용 — 로그인/팝업 처리 생략")
            return page
        await _log_in(page, login_timeout_ms, wehago_id, wehago_password)
        return page

    if _Session.playwright is None:
        from playwright.async_api import async_playwright

        _Session.playwright = await async_playwright().start()
    # --disable-notifications 등은 `_launch_browser` 참고.
    browser = await _launch_browser(_Session.playwright)
    browser.on("disconnected", _forget_browser)
    context = await browser.new_context(viewport=VIEWPORT)
    page = await context.new_page()
    _Session.browser, _Session.main_page = browser, page
    await _log_in(page, login_timeout_ms, wehago_id, wehago_password)
    return page


def start_wehago_upload_batch(
    items: list,
    result_queue: "Queue",
    max_concurrent: int = 5,
    login_timeout_ms: int = 300_000,
    wehago_id: str = "",
    wehago_password: str = "",
    expected_totals: "dict[Path, int] | None" = None,
) -> threading.Thread:
    """`ui/batch_worker.py`와 같은 패턴 — 결과는 `queue.Queue`로만 전달한다(Qt
    시그널 미사용, 크로스스레드 크래시 전례). 완료마다
    `("wehago_upload_item_done", UploadOutcome)`, 전부 끝나면
    `("wehago_upload_batch_finished", 오류메시지 또는 None)`.

    브라우저/로그인 세션은 `_Session`이 앱 실행 중 계속 들고 있어서 두 번째
    호출부터는 열려 있는 메인 화면을 재사용한다. Playwright는 무거워서 실제로
    부르기 전까지 import하지 않는다."""

    async def _run_batch() -> None:
        try:
            async with _Session.setup_lock:
                main_page = await _get_main_page(login_timeout_ms, wehago_id, wehago_password)
        except Exception as exc:  # noqa: BLE001
            logger.warning("위하고 세션 준비 실패: %r", exc)
            result_queue.put(("wehago_upload_batch_finished", f"위하고 로그인/브라우저 준비 실패: {exc}"))
            return
        try:
            semaphore = asyncio.Semaphore(max_concurrent)
            totals = expected_totals or {}

            async def _wrapped(info) -> None:
                outcome = await _process_one(
                    main_page, info, semaphore, _Session.search_lock, _Session.ui_lock, totals.get(info.path)
                )
                result_queue.put(("wehago_upload_item_done", outcome))

            await asyncio.gather(*(_wrapped(info) for info in items))
            result_queue.put(("wehago_upload_batch_finished", None))
        except Exception as exc:  # noqa: BLE001
            result_queue.put(("wehago_upload_batch_finished", str(exc)))

    loop = _ensure_worker_loop()
    asyncio.run_coroutine_threadsafe(_run_batch(), loop)
    return _Session.thread

