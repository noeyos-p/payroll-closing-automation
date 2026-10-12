"""Adapted narrow subset of employee_entry.py (observations 2026-10-08).

Source: neighboring '위하고 사원등록 자동입력', _click_cell, _type_name,
_fill_employee_dialog, register_new_employee. Removed matching, rejoining,
code allocation, all existing employee edits and broad notice dismissal.
0000000000000 acceptance and duplicate-SSN dialog remain live UNVERIFIED.
Only a duplicate resident-number notice may be acknowledged (user authorized).
No automatic retirement/rejoin/year-end notices or generic confirm buttons.
"""
from __future__ import annotations

import re

from ..payroll_browser import PayrollBrowserError, UnknownDataError


async def _one(locator, label):
    if await locator.count() != 1 or not await locator.is_visible():
        raise UnknownDataError(f"Temporary registration: {label} absent/ambiguous.")
    return locator


async def acknowledge_duplicate_ssn(page):
    marker = page.get_by_text(re.compile(r"동일.*주민|주민.*동일|주민.*중복")).filter(visible=True)
    count = await marker.count()
    if not count:
        return False
    if count != 1:
        raise UnknownDataError("Duplicate SSN notice ambiguous; no acknowledgement.")
    popup = marker.locator("xpath=ancestor::*[.//button[normalize-space()='확인']][1]")
    if await popup.count() != 1:
        raise UnknownDataError("Duplicate SSN notice container unavailable.")
    bounds = await popup.bounding_box()
    if not bounds or bounds["width"] > 1000 or bounds["height"] > 800:
        raise UnknownDataError("Duplicate SSN notice is not an isolated popup.")
    button = await _one(popup.get_by_role("button", name="확인", exact=True), "duplicate SSN confirm")
    await button.click()
    await marker.wait_for(state="hidden", timeout=5000)
    return True


async def register_temporary(page, grid_js, rows, code, name, join_day):
    try:
        await _register(page, grid_js, rows, code, name, join_day)
    except PayrollBrowserError:
        raise
    except Exception:
        # Playwright exception logs may expose the full mask-editor value.
        raise UnknownDataError("Temporary registration failed; inspect locally, do not retry blindly. Sensitive call log suppressed.") from None


async def _register(page, grid_js, rows, code, name, join_day):
    if any(str(r.get("cd_emp") or "") == code for r in rows):
        raise PayrollBrowserError("Existing employee code: no registration or edits allowed.")
    # Do not reuse a partial row or silently replace its existing code.
    blanks = [i for i, r in enumerate(rows) if not r.get("cd_emp") and not r.get("nm_krname")]
    if len(blanks) != 1 or blanks[0] != len(rows) - 1:
        raise UnknownDataError("Unique final empty employee row not established.")
    idx = blanks[0]
    state = await page.evaluate("""async ({expression, idx, code}) => {
        const g = await (0, eval)('(' + expression + ')')();
        if (!g) return null;
        const rows = g.getDataSource().getJsonRows();
        if (rows.some(r => String(r.cd_emp || '') === code)) return null;
        if (idx !== rows.length - 1 || rows[idx].cd_emp || rows[idx].nm_krname) return null;
        window.__payrollTemporaryGrid = g;
        const item = g.getItemIndex(idx);
        g.setCurrent({itemIndex: item, column: 'cd_emp'});
        g.setTopItem(Math.max(0, item - 3));
        return g.getCurrent();
    }""", {"expression": grid_js, "idx": idx, "code": code})
    if not state or state.get("dataRow") != idx or state.get("column") != "cd_emp":
        raise UnknownDataError("Blank employee Code cell not selected; no input.")
    await page.keyboard.type(code, delay=60)
    await page.keyboard.press("Enter")
    await page.wait_for_timeout(400)
    state_js = """() => {
        const g = window.__payrollTemporaryGrid;
        return {rows:g.getDataSource().getJsonRows(), current:g.getCurrent()};
    }"""
    state = await page.evaluate(state_js)
    if (str(state["rows"][idx].get("cd_emp") or "") != code
            or state["current"].get("dataRow") != idx
            or state["current"].get("column") != "nm_krname"):
        raise UnknownDataError("Temporary Code input unverified; no retry/next-code guessing.")
    temporary_name = name + "(임시)"
    for _ in range(3):
        await page.keyboard.type(temporary_name, delay=60)
        await page.wait_for_timeout(200)
        value = await page.evaluate("() => document.activeElement && document.activeElement.value")
        if value == temporary_name:
            await page.keyboard.press("Enter")
            break
        if value:
            await page.keyboard.press("Control+A")
            await page.keyboard.press("Backspace")
    else:
        raise UnknownDataError("Temporary name editor not verified.")
    title = page.get_by_text("직원정보 입력", exact=True).filter(visible=True)
    await title.wait_for(state="visible", timeout=10000)
    await _one(title, "employee dialog heading")
    dialog = title.locator("xpath=ancestor::*[.//button[normalize-space()='등록(tab)']][1]")
    await _one(dialog, "employee dialog")
    resident = dialog.locator("input[id='no_social']")
    if await resident.count() != 1:
        raise UnknownDataError("Resident editor is absent/ambiguous.")
    display = resident.locator("xpath=ancestor::div[contains(@class,'LS_ngh_input2')][1]").locator(".fake_inputbox")
    await (await _one(display, "resident display")).click()
    await page.keyboard.press("0")
    editor = dialog.locator("input[id='no_social']").filter(visible=True)
    await editor.wait_for(state="visible", timeout=5000)
    await _one(editor, "resident editor")
    await editor.press("Control+A")
    await page.keyboard.type("0" * 13, delay=80)
    if re.sub(r"\D", "", await editor.input_value()) != "0" * 13:
        raise UnknownDataError("Temporary resident-number editor value unverified.")
    await editor.press("Enter")
    await page.wait_for_timeout(500)
    await acknowledge_duplicate_ssn(page)
    field = await _one(dialog.locator(".fake_inputbox[id='da_enter']"), "join-date display")
    await field.click()
    digits = join_day.strftime("%Y%m%d")
    await page.keyboard.press(digits[0])
    await page.wait_for_timeout(300)
    active = await _one(dialog.locator("input:focus"), "join-date editor")
    await active.press("Control+A")
    await page.keyboard.type(digits, delay=80)
    await page.wait_for_timeout(200)
    if re.sub(r"\D", "", await field.inner_text()) != digits:
        raise UnknownDataError("First day of accrual month unverified; not registering.")
    await (await _one(dialog.get_by_role("button", name="등록(tab)", exact=True), "register button")).click()
    await page.wait_for_timeout(300)
    await acknowledge_duplicate_ssn(page)
    await title.wait_for(state="hidden", timeout=15000)
    # No escape of unrelated dialogs; persisted values verified by parent adapter.
