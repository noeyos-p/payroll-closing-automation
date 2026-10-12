"""Live WEHAGO steps observed on 2026-10-08 (업체002/업체003), as plain functions.

Every step verifies the screen before clicking and raises ``LiveStepError``
instead of guessing. Observations are recorded in CLAUDE.md (2026-10-08 절).
These functions mutate WEHAGO; callers decide when they are allowed.
"""
from __future__ import annotations

from datetime import date


class LiveStepError(RuntimeError):
    pass


LUX_JS = """() => [...document.querySelectorAll('.WSC_LUXDialog')].filter(e => e.getBoundingClientRect().width > 50)
  .map(e => ({text: e.innerText.trim().slice(0, 300), buttons: [...e.querySelectorAll('button')].map(b => b.innerText.trim())}))"""
ACTIVE_JS = "() => { const e = document.activeElement; return {tag: e.tagName, id: e.id, value: e.value}; }"
ROWS_JS = "() => { const g = window.Grids && window.Grids.getActiveGrid && window.Grids.getActiveGrid(); return g ? g.getDataSource().getJsonRows() : null; }"
CHECKED_ROWS_JS = "() => window.Grids.getActiveGrid().getCheckedRows()"
CHECK_ONLY_JS = "(i) => { const g = window.Grids.getActiveGrid(); g.checkAll(false); g.checkRow(i, true); }"
VALUES_JS = """() => [...document.querySelectorAll('input, .fakeinput, .fake_inputbox')]
  .map(i => (i.value || i.innerText || '').trim()).filter(v => v)"""
CLOSE_CONFIRM_TEXT = "현재 입력하는 급여등을 완료하시겠습니까?"


def _flat(text: str) -> str:
    return "".join(text.split())


async def screen_text(page) -> str:
    values = await page.evaluate(VALUES_JS)
    return _flat(await page.inner_text("body")) + "|" + "|".join(_flat(v) for v in values)


async def _visible_one(locator, what: str):
    hits = [i for i in range(await locator.count()) if await locator.nth(i).is_visible()]
    if len(hits) != 1:
        raise LiveStepError(f"{what}을(를) 하나로 찾지 못했습니다({len(hits)}개).")
    return locator.nth(hits[0])


async def wait_gone(page, text: str, seconds: int = 90) -> None:
    probe = page.get_by_text(text)
    for _ in range(seconds):
        await page.wait_for_timeout(1000)
        if not (await probe.count() and await probe.first.is_visible()):
            return
    raise LiveStepError(f"'{text}' 상태가 {seconds}초 안에 끝나지 않았습니다.")


async def export_ledger(page, save_dir, prefix: str, attempts: int = 3):
    """Ctrl+G from the hidden #Left_grid_line; retried because right after a
    payment opens the first press can be ignored (observed 2026-10-08)."""
    from pathlib import Path
    for _ in range(attempts):
        await page.bring_to_front()
        await page.wait_for_timeout(2000)
        await page.locator("#Left_grid_line").focus()
        try:
            async with page.expect_download(timeout=15000) as info:
                await page.keyboard.press("Control+g")
            download = await info.value
        except Exception:
            continue
        target = Path(save_dir) / f"{prefix}_{download.suggested_filename}"
        await download.save_as(str(target))
        return target, download.suggested_filename
    raise LiveStepError("급여대장 내보내기(Ctrl+G)가 시작되지 않았습니다.")


async def copy_payment(page, year: int, source_month: int, target_month: int, kind_code: str) -> str:
    """지급일자 modal: select source month → 복사 → set target month via its
    calendar → check the single source row → 확인. Returns the proposed dt_pay.
    Precondition: the target month has no payment (caller checked)."""
    from core import wehago_upload as wu
    await wu._select_accrual_month(page, source_month)
    await (await _visible_one(page.get_by_role("button", name="복사", exact=True), "'복사' 버튼")).click()
    await page.wait_for_timeout(2000)
    # The copy dialog nests two .dialog_content boxes; only the outer one holds 취소(esc)/확인.
    sub = (page.locator(".dialog_content").filter(has_text="→")
           .filter(has=page.get_by_role("button", name="확인", exact=True)).last)
    fields = sub.locator(".fakeinput")
    if await fields.count() != 2:
        raise LiveStepError("복사 창의 귀속월 칸이 예상과 다릅니다.")
    src_text = f"{year}.{source_month:02d}"
    if _flat(await fields.nth(0).inner_text()) != src_text:
        raise LiveStepError("복사 창의 원본 귀속월이 대상과 다릅니다.")
    box = await fields.nth(1).bounding_box()
    await page.mouse.click(box["x"] + box["width"] + 14, box["y"] + box["height"] / 2)  # calendar icon
    await page.wait_for_timeout(800)
    await (await _visible_one(page.get_by_text(f"{target_month}월", exact=True), f"월 선택기 {target_month}월")).click()
    await page.wait_for_timeout(1500)
    flat = _flat(await sub.inner_text())
    if f"{src_text}달력열기→{year}.{target_month:02d}" not in flat:
        raise LiveStepError("복사 창의 원본→대상 귀속월을 확인하지 못했습니다.")
    canv = await page.evaluate("""() => [...document.querySelectorAll('canvas')].map(c => { const r = c.getBoundingClientRect();
        return {left: r.left, top: r.top, width: r.width, height: r.height}; })
        .filter(r => r.width > 200 && r.width < 320 && r.height > 200 && r.height < 300).sort((a, b) => a.left - b.left)""")
    if len(canv) != 2:
        raise LiveStepError("복사 창의 지급건 목록 두 개를 찾지 못했습니다.")
    left, right = canv
    await page.mouse.click(right["left"] + 12, right["top"] + 12)
    await page.wait_for_timeout(400)
    proposed = await page.evaluate(ROWS_JS) or []
    await page.mouse.click(left["left"] + 12, left["top"] + 12)  # header click focuses the source list
    await page.wait_for_timeout(400)
    source_rows = await page.evaluate(ROWS_JS) or []
    same = [i for i, row in enumerate(source_rows) if str(row.get("fg_pay")) == kind_code]
    if len(same) != 1:
        raise LiveStepError("복사할 전월 같은 종류 지급건이 하나가 아닙니다.")
    # Check only that kind's row: 업체005 had 급여 and 급+상 last month (2026-10-12).
    await page.evaluate(CHECK_ONLY_JS, same[0])
    await page.wait_for_timeout(400)
    if await page.evaluate(CHECKED_ROWS_JS) != [same[0]]:
        raise LiveStepError("복사할 전월 지급건만 선택하지 못했습니다.")
    targets = [row for row in proposed if str(row.get("fg_pay")) == kind_code]
    if len(targets) != 1:
        raise LiveStepError("위하고가 제안한 대상 지급건을 확인하지 못했습니다.")
    dt_pay = str(targets[0].get("dt_pay"))
    await (await _visible_one(sub.get_by_role("button", name="확인", exact=True), "복사 '확인'")).click()
    await page.wait_for_timeout(1500)
    if await page.evaluate(LUX_JS):
        raise LiveStepError("복사 확인 후 예상하지 못한 안내창이 떴습니다.")
    await wait_gone(page, "데이터를 불러오고 있습니다", 90)
    cancel = page.get_by_role("button", name="취소(esc)", exact=True)
    hits = [i for i in range(await cancel.count()) if await cancel.nth(i).is_visible()]
    if hits:
        await cancel.nth(hits[-1]).click()
        await page.wait_for_timeout(1000)
    return dt_pay


DUPLICATE_RRN_TEXT = "동일한 주민등록번호가 존재합니다"


async def confirm_duplicate_rrn(roster) -> bool:
    """'주민번호 중복' window (not a WSC_LUXDialog; observed 2026-10-12 on the 2nd
    temp registration with 000000-0000000). User rule: confirm and continue.
    Clicks only that window's own 확인."""
    notice = roster.get_by_text(DUPLICATE_RRN_TEXT)
    if not (await notice.count() and await notice.first.is_visible()):
        return False
    box = roster.locator("div").filter(has=notice.first).filter(
        has=roster.get_by_role("button", name="확인", exact=True)).last
    await (await _visible_one(box.get_by_role("button", name="확인", exact=True), "주민번호 중복 '확인'")).click()
    await roster.wait_for_timeout(800)
    return True


async def register_employee(roster, code: str, name: str, join: date) -> None:
    """사원등록(인사): new grid row → code → name (F2 first: typing Hangul
    straight into the cell drops the first syllable) → 직원정보 입력 dialog:
    0. 내국인, 주민번호 000000-0000000, 입사일 → 등록(tab). A same-RRN notice
    is confirmed (user rule); any other dialog stops."""
    await roster.bring_to_front()
    canv = await roster.evaluate("""() => [...document.querySelectorAll('canvas')].map(c => { const r = c.getBoundingClientRect();
        return {left: r.left, top: r.top, width: r.width, height: r.height}; }).filter(r => r.width > 60 && r.height > 60).sort((a,b)=>a.left-b.left)""")
    await roster.mouse.click(canv[0]["left"] + 120, canv[0]["top"] + 40)
    await roster.wait_for_timeout(400)
    await roster.keyboard.press("Control+End")
    await roster.wait_for_timeout(400)
    await roster.keyboard.press("Home")
    await roster.wait_for_timeout(300)
    cur_js = """() => { const g = window.Grids.getActiveGrid(); const c = g.getCurrent(); const ds = g.getDataSource();
        const r = c.dataRow >= 0 && c.dataRow < ds.getRowCount() ? ds.getJsonRow(c.dataRow) : {};
        return {field: c.fieldName, cd_emp: r.cd_emp, nm_krname: r.nm_krname}; }"""
    cur = await roster.evaluate(cur_js)
    active = await roster.evaluate(ACTIVE_JS)
    if not (cur["field"] == "cd_emp" and not cur["cd_emp"] and active["id"] == "gridOne_line"
            and set(active["value"] or "_") <= {"_"}):
        raise LiveStepError("사원등록 신규 입력 행이 비어 있지 않습니다. 기존 사원을 수정하지 않았습니다.")
    await roster.keyboard.type(code)
    await roster.keyboard.press("Enter")
    await roster.wait_for_timeout(1200)
    cur = await roster.evaluate(cur_js)
    if cur["cd_emp"] != code or cur["field"] != "nm_krname" or await roster.evaluate(LUX_JS):
        raise LiveStepError("사원코드 입력이 확인되지 않았습니다.")
    await roster.keyboard.press("F2")
    await roster.wait_for_timeout(400)
    if (await roster.evaluate(ACTIVE_JS))["id"] != "gridOne_line":
        raise LiveStepError("이름 입력칸이 열리지 않았습니다.")
    await roster.keyboard.press("Control+a")
    await roster.keyboard.type(name)
    await roster.keyboard.press("Enter")
    await roster.wait_for_timeout(1500)
    cur = await roster.evaluate(cur_js)
    if cur["nm_krname"] != name:
        raise LiveStepError("이름이 정확히 입력되지 않았습니다.")
    body = _flat(await roster.inner_text("body"))
    if "직원정보입력" not in body:
        raise LiveStepError("'직원정보 입력' 창이 뜨지 않았습니다.")
    options = roster.get_by_text("0. 내국인", exact=True)
    vis = [i for i in range(await options.count()) if await options.nth(i).is_visible()]
    if vis:
        await options.nth(vis[-1]).click()
        await roster.wait_for_timeout(400)
    rrn = await _visible_one(roster.get_by_text("주민번호를 입력해주세요.", exact=True), "주민번호 칸")
    box = await rrn.bounding_box()
    await roster.mouse.click(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
    await roster.wait_for_timeout(500)
    if (await roster.evaluate(ACTIVE_JS))["id"] != "no_social":
        raise LiveStepError("주민번호 입력칸이 열리지 않았습니다.")
    await roster.keyboard.press("Control+a")
    await roster.keyboard.type("0000000000000")
    await roster.keyboard.press("Tab")
    await roster.wait_for_timeout(1000)
    await confirm_duplicate_rrn(roster)
    if await roster.evaluate(LUX_JS):
        dialog = (await roster.evaluate(LUX_JS))[-1]
        if not any(w in dialog["text"] for w in ("주민", "동일", "중복")) or "확인" not in dialog["buttons"]:
            raise LiveStepError("주민번호 입력 후 예상하지 못한 안내창이 떴습니다.")
        await roster.locator(".WSC_LUXDialog").last.get_by_role("button", name="확인", exact=True).click()
        await roster.wait_for_timeout(800)
    fields = await roster.evaluate("""() => [...document.querySelectorAll('[id="da_enter"]')].map(e => { const r = e.getBoundingClientRect();
        return {x: r.left, y: r.top, w: r.width, h: r.height}; }).filter(r => r.w > 0)""")
    dialog_field = [f for f in fields if f["y"] > 400]
    if len(dialog_field) != 1:
        raise LiveStepError("입사일 칸을 찾지 못했습니다.")
    f = dialog_field[0]
    # Needs two clicks to open; a double click alone missed on the 2nd registration
    # in a row (업체049재가, 2026-10-12), so click until the input really has focus.
    for _ in range(4):
        await roster.mouse.click(f["x"] + min(40, f["w"] / 3), f["y"] + f["h"] / 2)
        await roster.wait_for_timeout(500)
        if (await roster.evaluate(ACTIVE_JS))["id"] == "da_enter":
            break
    else:
        raise LiveStepError("입사일 입력칸이 열리지 않았습니다.")
    await roster.keyboard.press("Control+a")
    await roster.keyboard.type(join.strftime("%Y%m%d"))
    await roster.keyboard.press("Enter")
    await roster.wait_for_timeout(800)
    body = _flat(await roster.inner_text("body"))
    for text in (code, name, join.strftime("%Y.%m.%d"), "000000-0000000", "0.내국인"):
        if _flat(text) not in body:
            raise LiveStepError(f"직원정보 입력 값 확인 실패: {text}")
    await (await _visible_one(roster.get_by_role("button", name="등록(tab)", exact=True), "'등록(tab)' 버튼")).click()
    for _ in range(3):
        await roster.wait_for_timeout(1500)
        if await confirm_duplicate_rrn(roster):
            continue
        dialogs = await roster.evaluate(LUX_JS)
        if not dialogs:
            return
        d = dialogs[-1]
        if any(w in d["text"] for w in ("주민", "동일", "중복")) and "확인" in d["buttons"]:
            await roster.locator(".WSC_LUXDialog").last.get_by_role("button", name="확인", exact=True).click()
            continue
        raise LiveStepError("등록 후 예상하지 못한 안내창: " + d["text"][:80])


async def close_payment(page, company: str, year: int, accrual_month: int, payment_date: date, kind_label: str) -> None:
    """'완료' → confirm dialog text check → 확인 (observed twice)."""
    flat = await screen_text(page)
    expected = [_flat(company), f"{year}.{accrual_month:02d}", payment_date.strftime("%Y.%m.%d"), _flat(kind_label)]
    missing = [t for t in expected if t not in flat]
    if missing:
        raise LiveStepError(f"열린 급여 화면이 대상과 다릅니다: {missing}")
    await (await _visible_one(page.get_by_role("button", name="완료", exact=True), "'완료' 버튼")).click()
    await page.wait_for_timeout(1500)
    dialog = page.locator(".WSC_LUXDialog").filter(has_text=CLOSE_CONFIRM_TEXT)
    if await dialog.count() != 1:
        raise LiveStepError("완료 확인창이 예상과 다릅니다. 확인을 누르지 않았습니다.")
    await dialog.get_by_role("button", name="확인", exact=True).click()
    await page.wait_for_timeout(2500)
