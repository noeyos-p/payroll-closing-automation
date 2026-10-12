"""WEHAGO adapter for core/payroll_workflow.run_workflow (Playwright, async).

Verified elsewhere in this project (live, 2026-09-21..29, core/wehago_upload):
  company search → 급여 tab → 급여자료입력 → 지급일자 modal → accrual month by
  arrow keys → right grid rows {fg_pay, dt_pay, yn_close, am_orpay, ...} →
  "실행" opens a payment → Ctrl+U Excel upload → 변환 → save prompt on 지급일자.
Observed live 2026-10-08: Ctrl+G ledger export (focus hidden #Left_grid_line),
사원등록(인사) via menu search -> SWPM0108 tab, roster read from its RealGrid.
Observed live 2026-10-08: close = '완료' + confirm dialog (close_if_unchanged).
NOT observed: the payment copy screen flow and the temporary registration key
sequence. Those raise
``LiveBindingRequired`` instead of guessing; ``inspect_screen`` collects button
and input names (never values) to bind them in the live step.

Mutations are refused unless ``allow_mutation`` is True. Nothing here logs
names, resident numbers or passwords.
"""
from __future__ import annotations

import asyncio
from dataclasses import replace
from datetime import date, datetime
from hashlib import sha256
from pathlib import Path
import re
from typing import Mapping

from core.payroll_workflow import (
    Employee, OriginalPayroll, PaymentIdentity, PayrollSnapshot, TemporaryRegistration, normalize_name,
)
from core.wehago_export import ExportFormatError, WehagoExport, parse_export_filename, read_wehago_export

KIND_CODES = {"급여": "1", "급+상": "2", "추급": "5"}  # observed 2026-09-28
CLOSE_CONFIRM_TEXT = "현재 입력하는 급여등을 완료하시겠습니까?"  # observed 2026-10-08


class PayrollBrowserError(RuntimeError):
    """Safe diagnostic; contains no personnel values."""


class LiveBindingRequired(PayrollBrowserError):
    """A WEHAGO screen step that nobody has observed yet; do not guess it."""


def _key(text: str) -> str:
    return normalize_name(text)


def matching_payment_rows(rows: list[dict], payment: PaymentIdentity) -> list[int]:
    """Rows of the payment's kind whose dt_pay falls in its payment month."""
    code = KIND_CODES.get(payment.payment_kind)
    if code is None:
        raise PayrollBrowserError(f"알 수 없는 지급구분: {payment.payment_kind}")
    hits = []
    for index, row in enumerate(rows):
        if not isinstance(row, dict) or not {"fg_pay", "dt_pay", "yn_close"} <= row.keys():
            raise PayrollBrowserError("지급일자 목록 형식이 예상과 다릅니다(fg_pay/dt_pay/yn_close).")
        raw = str(row["dt_pay"] or "")
        if not re.fullmatch(r"\d{8}", raw):
            raise PayrollBrowserError("지급일자 형식이 예상과 다릅니다.")
        if str(row["fg_pay"]).strip() == code and int(raw[4:6]) == payment.payment_month:
            hits.append(index)
    return hits


def payment_with_date(payment: PaymentIdentity, row: dict) -> PaymentIdentity:
    raw = str(row["dt_pay"])
    dated = replace(payment, payment_date=date(int(raw[:4]), int(raw[4:6]), int(raw[6:])))
    if not payment.same_payment(dated):
        raise PayrollBrowserError("지급건 날짜가 대상과 다릅니다.")
    return dated


def is_closed(row: dict) -> bool:
    value = str(row["yn_close"]).strip()
    if value not in ("0", "1"):
        raise PayrollBrowserError("마감 표시(yn_close) 값이 0/1이 아닙니다.")
    return value == "1"


def export_revision(export: WehagoExport, closed: bool) -> str:
    parts = [f"{p.employee_code}|{p.amounts.gross}|{p.amounts.deductions}|{p.amounts.net}"
             for p in export.people]
    return sha256(("\n".join(sorted(parts)) + f"\nclosed={closed}").encode()).hexdigest()[:16]


def check_export_identity(filename: str, company_name: str, payment: PaymentIdentity) -> None:
    ident = parse_export_filename(filename)
    if (ident.year, ident.month) != (payment.year, payment.accrual_month):
        raise PayrollBrowserError("내려받은 급여대장의 귀속월이 대상과 다릅니다.")
    if _key(re.sub(r"\(\d{4}\)$", "", ident.company)) != _key(re.sub(r"\(\d{4}\)$", "", company_name)):
        raise PayrollBrowserError("내려받은 급여대장의 업체명이 업체리스트와 다릅니다. 같은 업체인지 확인이 필요합니다.")


ROSTER_FIELDS = ("cd_emp", "nm_krname", "da_retire")  # observed live 2026-10-08 (SWPM0108 grid)


def employees_from_rows(rows: list[dict], accrual_first_day: date) -> tuple[Employee, ...]:
    """사원등록(인사) RealGrid rows → employees. Only code/name/retire date are
    read; the grid also carries full resident numbers (no_social), which are
    never touched. The trailing blank new-entry row is skipped. An employee who
    retired before the accrual month is marked excluded."""
    employees, seen = [], set()
    for row in rows:
        if not isinstance(row, dict) or not {"cd_emp", "nm_krname"} <= row.keys():
            raise PayrollBrowserError("사원등록 표 형식이 예상과 다릅니다(cd_emp/nm_krname).")
        code, name = str(row["cd_emp"] or "").strip(), str(row["nm_krname"] or "").strip()
        if not code and not name:
            continue
        if not code or not name:
            raise PayrollBrowserError("사원등록 표에 코드나 이름이 빈 사원이 있습니다.")
        if code in seen:
            raise PayrollBrowserError("사원등록 표에 같은 사원코드가 반복됩니다.")
        seen.add(code)
        reason = None
        retired = row.get("da_retire")
        if retired not in (None, ""):
            day = _as_date(retired)
            if day is None:
                raise PayrollBrowserError("사원등록 퇴사일자 형식을 읽지 못했습니다.")
            if day < accrual_first_day:
                reason = f"퇴사({day.isoformat()})"
        employees.append(Employee(code, name, reason))
    if not employees:
        raise PayrollBrowserError("사원등록 표에서 사원을 읽지 못했습니다.")
    return tuple(employees)


def _as_date(value) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    match = re.fullmatch(r"(\d{4})[.\-/]?(\d{2})[.\-/]?(\d{2})", str(value).strip())
    return date(*map(int, match.groups())) if match else None


_INSPECT_JS = """() => {
  const visible = el => { const r = el.getBoundingClientRect(); return r.width > 0 && r.height > 0; };
  const text = el => (el.innerText || el.value === undefined ? (el.innerText || '') : '').trim().slice(0, 40);
  return {
    url: location.pathname + location.hash,
    buttons: [...document.querySelectorAll('button')].filter(visible).map(b => b.innerText.trim()).filter(Boolean).slice(0, 200),
    inputs: [...document.querySelectorAll('input')].filter(visible).map(i => ({id: i.id, type: i.type, placeholder: i.placeholder})).slice(0, 100),
    headings: [...document.querySelectorAll('h1,h2,h3,h4,.tit,.title')].filter(visible).map(text).filter(Boolean).slice(0, 50),
    grids: [...document.querySelectorAll('[id$="_grid"],[id$="_grid_line"],[id$="_grid_view"]')].map(g => g.id).slice(0, 50),
  };
}"""


async def inspect_screen(page) -> dict:
    """Names only (buttons, input ids, headings, grid ids) — never field values."""
    return await page.evaluate(_INSPECT_JS)


class WehagoPayrollAdapter:
    """One adapter per run. ``companies`` maps business number → list name."""

    def __init__(self, main_page, companies: Mapping[str, str], download_dir: Path, *,
                 allow_mutation: bool = False, search_lock: asyncio.Lock | None = None,
                 ui_lock: asyncio.Lock | None = None, export_closed: bool = False):
        self.main_page = main_page
        self.companies = dict(companies)
        self.download_dir = Path(download_dir)
        self.download_dir.mkdir(parents=True, exist_ok=True)
        self.allow_mutation = allow_mutation
        # Closed payments are skipped (user rule) without export unless this
        # verification switch is on.
        self.export_closed = export_closed
        self.search_lock = search_lock or asyncio.Lock()
        self.ui_lock = ui_lock or asyncio.Lock()
        self._pages: dict[str, object] = {}
        self._plans: dict[tuple, object] = {}
        self._created: dict[tuple, PaymentIdentity] = {}
        self._registered: dict[tuple, dict[str, str]] = {}
        self._rosters: dict[int, object] = {}
        # Payments this run just closed: the next read must export to prove
        # the persisted amounts are still the reconciled ones.
        self._verify_after_close: set[tuple] = set()
        # Live-step evidence: screen names (no values) and payment-list rows
        # (company-level codes/dates/totals only), keyed by step label.
        self.observations: list[dict] = []

    async def _observe(self, page, label: str, **extra) -> None:
        try:
            screen = await inspect_screen(page)
        except Exception as exc:  # observation must never break the run
            screen = {"error": type(exc).__name__}
        self.observations.append({"step": label, "screen": screen, **extra})

    # --- wiring -----------------------------------------------------------
    def register_upload_plan(self, payment: PaymentIdentity, plan) -> None:
        self._plans[payment.key] = plan

    def _require_mutation(self, what: str) -> None:
        if not self.allow_mutation:
            raise PayrollBrowserError(f"읽기 전용 실행: {what}을(를) 하지 않았습니다.")

    async def _company_page(self, payment: PaymentIdentity):
        page = self._pages.get(payment.company_id)
        if page is not None and not page.is_closed():
            return page
        from core import wehago_upload as wu
        name = self.companies.get(payment.company_id)
        if not name:
            raise PayrollBrowserError("업체리스트에 없는 사업자번호입니다.")
        async with self.search_lock:
            page = await _open_company(self.main_page, name, payment.company_id, payment.year)
        await page.set_viewport_size(wu.VIEWPORT)
        await wu._open_payroll_entry(page)
        await self._observe(page, "급여자료입력 진입")
        self._pages[payment.company_id] = page
        return page

    async def _payment_rows(self, page, payment: PaymentIdentity) -> list[dict]:
        from core import wehago_upload as wu
        await wu._select_accrual_month(page, payment.accrual_month)
        try:
            return await wu._click_right_side_grid(page)
        except RuntimeError:
            # An accrual month without payments shows an empty right grid
            # (observed 2026-10-08/12). Prove "empty" from the left month list:
            # the target month row must show 0 people; otherwise re-raise.
            rects = await page.evaluate(wu._FIND_MODAL_CANVASES_JS)
            if not rects:
                raise
            left = rects[0]
            await page.mouse.click(left["left"] + left["width"] / 2, left["top"] + 40)
            await page.wait_for_timeout(300)
            months = await page.evaluate(wu._GET_ACTIVE_GRID_ROWS_JS) or []
            target = [r for r in months if isinstance(r, dict)
                      and str(r.get("ym_rvrs", ""))[-2:] == f"{payment.accrual_month:02d}"]
            if len(months) == 12 and len(target) == 1 and int(target[0].get("num") or 0) == 0:
                return []
            raise

    async def _open_row(self, page, rows: list[dict], index: int) -> None:
        for _ in range(len(rows)):
            await page.keyboard.press("ArrowUp")
        for _ in range(index):
            await page.keyboard.press("ArrowDown")
            await page.wait_for_timeout(120)
        await page.get_by_role("button", name="실행", exact=True).click()
        await page.wait_for_load_state()
        await page.wait_for_timeout(800)

    async def _export(self, page, payment: PaymentIdentity) -> WehagoExport:
        """1.52 path: focus 사원표 (#Left_grid_line) then Ctrl+G 급여대장 내보내기."""
        async with self.ui_lock:
            await page.bring_to_front()
            grid = page.locator("#Left_grid_line")
            if await grid.count() != 1:
                raise PayrollBrowserError("사원표(#Left_grid_line)를 하나로 찾지 못했습니다.")
            # Live 2026-10-08: a hidden aria-hidden <input> (keyboard sink), so focus, never click.
            await grid.focus()
            try:
                async with page.expect_download(timeout=30_000) as info:
                    await page.keyboard.press("Control+g")
            except Exception:
                raise PayrollBrowserError("급여대장 내보내기(Ctrl+G)가 30초 안에 시작되지 않았습니다.") from None
            download = await info.value
        name = download.suggested_filename
        target = self.download_dir / f"{payment.company_id}_{payment.payment_kind}_{datetime.now():%H%M%S}_{name}"
        await download.save_as(str(target))
        check_export_identity(name, self.companies[payment.company_id], payment)
        try:
            return read_wehago_export(target)
        except ExportFormatError as exc:
            raise PayrollBrowserError(f"내려받은 급여대장 확인 실패: {exc}") from None

    # --- PayrollAdapter -----------------------------------------------------
    async def list_closed(self, payment: PaymentIdentity) -> bool:
        return False  # The company list O is handled by the runner before this.

    async def read_payroll(self, payment: PaymentIdentity) -> PayrollSnapshot:
        page = await self._company_page(payment)
        rows = await self._payment_rows(page, payment)
        await self._observe(page, "지급일자 목록", rows=[
            {k: r.get(k) for k in ("fg_pay", "nm_pay", "dt_pay", "num", "yn_close", "am_orpay", "fg_submit")}
            for r in rows], row_keys=sorted({k for r in rows for k in r}))
        hits = matching_payment_rows(rows, payment)
        if not hits:
            await page.keyboard.press("Escape")  # closes only the 지급일자 modal (취소(esc))
            return PayrollSnapshot(payment, (), None, False, revision="absent")
        if len(hits) > 1:
            await page.keyboard.press("Escape")
            raise PayrollBrowserError("같은 지급월·지급구분의 지급건이 여러 개입니다.")
        row = rows[hits[0]]
        dated = payment_with_date(payment, row)
        closed = is_closed(row)
        if closed and not self.export_closed and payment.key not in self._verify_after_close:
            await page.keyboard.press("Escape")
            return PayrollSnapshot(dated, (), None, True, True, "closed")
        await self._open_row(page, rows, hits[0])
        await self._observe(page, "지급건 열기")
        export = await self._export(page, dated)
        return PayrollSnapshot(dated, export.people, export.totals, True, closed,
                               export_revision(export, closed))

    async def list_employees(self, payment: PaymentIdentity) -> tuple[Employee, ...]:
        page = await self._company_page(payment)
        roster = await self._roster_page(page)
        rows = await _read_roster_rows(roster)
        return employees_from_rows(rows, payment.join_date)

    async def _roster_page(self, page):
        cached = self._rosters.get(id(page))
        if cached is not None and not cached.is_closed():
            await cached.reload()
            await cached.wait_for_timeout(2500)
        else:
            cached = await _open_roster(page)
            self._rosters[id(page)] = cached
        from core import wehago_upload as wu
        await wu._dismiss_known_popups(cached)
        await self._observe(cached, "사원등록")
        return cached

    async def register_temporary(self, payment: PaymentIdentity,
                                 request: TemporaryRegistration) -> Employee:
        self._require_mutation("임시 사원등록")
        raise LiveBindingRequired("사원등록 신규 입력 순서는 아직 라이브로 확인하지 않았습니다.")

    async def copy_previous_payment(self, payment: PaymentIdentity) -> PaymentIdentity:
        self._require_mutation("전월 지급건 복사")
        raise LiveBindingRequired("지급일자 '복사' 화면은 아직 라이브로 확인하지 않았습니다.")

    async def upload_to_created(self, original: OriginalPayroll) -> bool:
        self._require_mutation("급여 업로드")
        if self._created.get(original.payment.key) != original.payment:
            return False  # Only a payment this run created may be filled.
        plan = self._plans.get(original.payment.key)
        if plan is None:
            raise PayrollBrowserError("업로드 파일이 준비되지 않았습니다.")
        from core import wehago_upload as wu
        codes = {normalize_name(p.name): p.employee_code for p in original.people if p.employee_code}
        path = plan.write(self.download_dir / f"upload_{original.payment.company_id}_{original.payment.payment_kind}.xlsx",
                          codes)
        page = await self._company_page(original.payment)
        await wu._upload_excel(page, path, self.ui_lock)
        await wu._confirm_and_convert(page)
        return True

    async def close_if_unchanged(self, persisted: PayrollSnapshot) -> bool:
        """Observed live 2026-10-08 (업체048): payroll screen button
        '완료' → WSC_LUXDialog "현재 입력하는 급여등을 완료하시겠습니까?" [취소][확인]
        → 확인; the button then reads '해제' and the payment row yn_close=1.
        The payment is re-read and re-exported first; any change refuses."""
        self._require_mutation("마감")
        current = await self.read_payroll(persisted.payment)  # leaves that payment open
        if current.closed or current.revision != persisted.revision:
            return False
        page = await self._company_page(persisted.payment)
        flat = "".join((await page.inner_text("body")).split())
        payment = persisted.payment
        expected = (f"{payment.year}.{payment.accrual_month:02d}", payment.payment_date.strftime("%Y.%m.%d"))
        if not all(text in flat for text in expected):
            raise PayrollBrowserError("열린 급여 화면의 귀속월/지급일이 대상과 다릅니다. 마감하지 않았습니다.")
        button = page.get_by_role("button", name="완료", exact=True)
        visible = [i for i in range(await button.count()) if await button.nth(i).is_visible()]
        if len(visible) != 1:
            raise PayrollBrowserError("'완료' 버튼을 하나로 찾지 못했습니다. 마감하지 않았습니다.")
        await button.nth(visible[0]).click()
        await page.wait_for_timeout(1500)
        dialog = page.locator(".WSC_LUXDialog").filter(has_text=CLOSE_CONFIRM_TEXT)
        if await dialog.count() != 1:
            raise PayrollBrowserError("완료 확인창이 예상과 다릅니다. 확인을 누르지 않았습니다.")
        await dialog.get_by_role("button", name="확인", exact=True).click()
        await page.wait_for_timeout(2500)
        self._verify_after_close.add(payment.key)
        return True


_RESULT_ROWS_JS = r"""() => [...document.querySelectorAll('a[id^="tooltip_"]')].flatMap(a => {
  let r = a.parentElement;
  const hasPay = el => [...el.querySelectorAll('button')].some(b => b.innerText.trim() === '급여');
  while (r && r !== document.body) {
    if (hasPay(r) && /\d{3}-\d{2}-\d{5}/.test(r.innerText)) break;
    r = r.parentElement;
  }
  if (!r || r === document.body) return [];
  return [{company: a.innerText.trim(), businessNumber: r.innerText.match(/\d{3}-\d{2}-\d{5}/)[0], anchor: a.id}];
})"""


def choose_company_row(rows: list[dict], business_number: str, accrual_year: int) -> dict:
    """Exact business number (1.52 rule); among same-number rows the one whose
    "(YYYY)" equals the accrual year, else the single row without a year."""
    same = [r for r in rows if re.sub(r"\D", "", r["businessNumber"]) == business_number]
    if not same:
        raise PayrollBrowserError("업체명·사업자번호가 모두 맞는 수임처를 찾지 못했습니다.")
    def year(r):
        m = re.search(r"\((\d{4})\)\s*$", r["company"])
        return int(m.group(1)) if m else None
    chosen = [r for r in same if year(r) == accrual_year] or [r for r in same if year(r) is None]
    if len(chosen) != 1:
        years = sorted({year(r) for r in same if year(r)})
        raise PayrollBrowserError(f"사업자번호가 같은 수임처 중 {accrual_year}년을 하나로 정하지 못했습니다(검색된 연도 {years}).")
    return chosen[0]


async def _open_company(main_page, company_name: str, business_number: str, accrual_year: int):
    """Search by name; read every result row in ONE evaluate (rows re-render
    while loading — reading them one by one failed live on 2026-10-08) and wait
    until two consecutive reads agree, then open the chosen row's 급여 tab."""
    from core import wehago_upload as wu
    await wu._ensure_t_edge_unused_tab(main_page)
    box = main_page.get_by_placeholder(wu.SEARCH_PLACEHOLDER)
    await box.fill(company_name)
    await box.press("Enter")
    previous = None
    rows: list[dict] = []
    for wait_ms in (500, 500, 700, 900, 1200, 1500):
        await main_page.wait_for_timeout(wait_ms)
        rows = await main_page.evaluate(_RESULT_ROWS_JS)
        if rows and rows == previous:
            break
        previous = rows
    else:
        if not rows:
            raise PayrollBrowserError("수임처 검색 결과가 없습니다.")
        raise PayrollBrowserError("수임처 검색 결과가 안정되지 않았습니다.")
    chosen = choose_company_row(rows, business_number, accrual_year)
    anchor = main_page.locator(f'[id="{chosen["anchor"]}"]')
    row = anchor.locator('xpath=ancestor::*[.//button[normalize-space()="급여"]][1]')
    button = row.get_by_role("button", name="급여", exact=True)
    if await button.count() != 1:
        raise PayrollBrowserError("선택한 수임처의 급여 버튼이 하나가 아닙니다.")
    async with main_page.expect_popup() as popup:
        await button.click()
    page = await popup.value
    await page.wait_for_load_state()
    return page


async def _open_roster(page):
    """Payroll tab → menu search '사원등록' → exactly '사원등록(인사)' (live
    2026-10-08: 5 hits incl. 부서/회계/일용직/종교인 variants; nested li>a with
    the same text, the innermost is clicked) → new tab SWPM0108."""
    from core import wehago_upload as wu
    search = page.locator("#gnb_search")
    if await search.count() != 1:
        raise PayrollBrowserError("메뉴 검색창(#gnb_search)을 찾지 못했습니다.")
    await search.fill("사원등록")
    await page.wait_for_timeout(1500)
    items = page.locator("li, a").filter(has_text="사원등록(인사)")
    hits = [i for i in range(await items.count()) if await items.nth(i).is_visible()
            and (await items.nth(i).inner_text()).replace(" ", "").strip() == "사원등록(인사)"]
    if not hits:
        raise PayrollBrowserError("메뉴 검색에서 '사원등록(인사)'를 찾지 못했습니다.")
    async with page.context.expect_page() as info:
        await items.nth(hits[-1]).click()
    roster = await info.value
    await roster.set_viewport_size(wu.VIEWPORT)
    await roster.wait_for_load_state()
    await roster.wait_for_timeout(3000)
    if "SWPM0108" not in roster.url:
        raise PayrollBrowserError("사원등록(인사) 화면이 열리지 않았습니다.")
    return roster


_ROSTER_CANVAS_JS = """() => [...document.querySelectorAll('canvas')].map(c => {
  const r = c.getBoundingClientRect(); return {left: r.left, top: r.top, width: r.width, height: r.height}; })
  .filter(r => r.width > 60 && r.height > 60).sort((a, b) => a.left - b.left)"""

# Only these fields leave the page; full resident numbers stay in the browser.
_ROSTER_ROWS_JS = """() => {
  const g = window.Grids && window.Grids.getActiveGrid && window.Grids.getActiveGrid();
  if (!g) return null;
  return g.getDataSource().getJsonRows().map(r => ({cd_emp: r.cd_emp, nm_krname: r.nm_krname, da_retire: r.da_retire}));
}"""


async def _read_roster_rows(roster) -> list[dict]:
    canvases = await roster.evaluate(_ROSTER_CANVAS_JS)
    if not canvases:
        raise PayrollBrowserError("사원등록 표를 찾지 못했습니다.")
    left = canvases[0]
    await roster.mouse.click(left["left"] + 120, left["top"] + 40)  # focus only; single click
    await roster.wait_for_timeout(400)
    rows = await roster.evaluate(_ROSTER_ROWS_JS)
    await roster.keyboard.press("Escape")
    if not isinstance(rows, list):
        raise PayrollBrowserError("사원등록 표 데이터를 읽지 못했습니다.")
    return rows
