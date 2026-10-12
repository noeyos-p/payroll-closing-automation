from pathlib import Path

import openpyxl

from core.wehago_upload import _classify_pay_type_text, _find_pay_type_row_index, read_excel_net_total


def _write_sample(path: Path) -> None:
    """`core/export.py`가 실제로 만드는 형태를 흉내낸다 — 마지막 행에 A열이
    "합계"인 총계 행이 같이 있다(실사로 확인, 2026-09-21). 이 총계 행까지
    사원 행으로 잘못 더하면 실제로 합계가 정확히 2배로 부풀어 오르는 버그가
    있었다 — 이 테스트가 그 회귀를 고정한다."""
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(["사원코드", "사원명", "기본급", "차인지급액"])
    sheet.append(["10", "사원025", 2_610_000, 4_084_250])
    sheet.append(["11", "사원026", 1_795_680, 2_886_700])
    sheet.append(["합계", "", 4_405_680, 6_970_950])
    workbook.save(path)


def test_read_excel_net_total_excludes_the_totals_row(tmp_path):
    path = tmp_path / "260921_테스트회사08(급여).xlsx"
    _write_sample(path)

    total = read_excel_net_total(path)

    assert total == 4_084_250 + 2_886_700  # "합계" 행(6,970,950)은 두 번 더해지면 안 됨


def test_read_excel_net_total_requires_net_pay_column(tmp_path):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(["사원코드", "사원명", "기본급"])
    sheet.append(["10", "사원025", 2_610_000])
    path = tmp_path / "no_net_pay_column.xlsx"
    workbook.save(path)

    try:
        read_excel_net_total(path)
    except RuntimeError as exc:
        assert "차인지급액" in str(exc)
    else:
        raise AssertionError("차인지급액 열이 없으면 RuntimeError가 나야 한다")


def test_classify_pay_type_text_distinguishes_bonus_from_plain_salary():
    """2026-09-22 버그 회귀 고정: 화면 문구 "급여+상여"는 파일명 토큰
    "급+상"의 글자 그대로 부분문자열이 아니라서(급,+,상 vs 급,여,+,상,여)
    옛 로직(`pay_type in row_text`)으로는 절대 못 찾았다. "급여"도
    "급여+상여"의 부분문자열이라 순서상 먼저 걸리면 엉뚱하게 매칭될 수
    있었다."""
    assert _classify_pay_type_text("2 급여+상여") == "급+상"
    assert _classify_pay_type_text("1 급여") == "급여"
    assert _classify_pay_type_text("3 추급") == "추급"
    assert _classify_pay_type_text("아무 상관 없는 텍스트") is None


# 2026-09-22: 실제 "업체079" 라이브 재현에서 그대로 가져온
# raw rows — getJsonRows()는 지급구분 텍스트를 전혀 안 주고(fg_pay/nm_pay는
# 숫자 코드뿐) am_orpay(차인지급액)만 실제 금액을 담고 있다는 걸 실측으로
# 확정했다. 급여 파일의 차인지급액(프로그램 계산값 8,331,530)이 fg_pay='1'
# 행과, 급+상 파일(7,338,970)이 fg_pay='2' 행과, 추급 파일(4,822,500)이
# fg_pay='5' 행과 원 단위까지 정확히 일치했다.
_LIVE_DONURI_ROWS = [
    {"fg_pay": "1", "nm_pay": "1", "dt_pay": "20260925", "num": 4, "yn_close": "0",
     "am_allowpay": 9658410, "am_orpay": 8331530, "fg_submit": ""},
    {"fg_pay": "2", "nm_pay": "2", "dt_pay": "20260925", "num": 10, "yn_close": "0",
     "am_allowpay": 7802740, "am_orpay": 7338970, "fg_submit": ""},
    {"fg_pay": "5", "nm_pay": "5", "dt_pay": "20260925", "num": 6, "yn_close": "0",
     "am_allowpay": 5072990, "am_orpay": 4822500, "fg_submit": ""},
]


def test_find_pay_type_row_index_matches_by_amount_when_rows_have_no_readable_text():
    """2026-09-22 버그 회귀 고정: 텍스트 매칭(`_classify_pay_type_text`)만으로는
    이 실제 데이터에서 급여/급+상/추급 전부 못 찾았다(로그로 재현 확인) — 이제
    `expected_total`(프로그램이 계산한 차인지급액)로 세 파일 모두 정확히
    서로 다른 행을 찾아야 한다."""
    assert _find_pay_type_row_index(_LIVE_DONURI_ROWS, "급여", 8_331_530) == 0
    assert _find_pay_type_row_index(_LIVE_DONURI_ROWS, "급+상", 7_338_970) == 1
    assert _find_pay_type_row_index(_LIVE_DONURI_ROWS, "추급", 4_822_500) == 2


def test_find_pay_type_row_index_falls_back_to_text_when_amount_is_unknown():
    rows = [{"note": "1 급여"}, {"note": "2 급여+상여"}]
    assert _find_pay_type_row_index(rows, "급+상", expected_total=None) == 1


def test_find_pay_type_row_index_accepts_changed_amount():
    assert _find_pay_type_row_index(_LIVE_DONURI_ROWS, "급여", expected_total=999) == 0


def test_find_pay_type_row_index_rejects_wrong_type_even_for_single_row():
    assert _find_pay_type_row_index([_LIVE_DONURI_ROWS[0]], "급+상", expected_total=None) is None


def test_pay_type_never_uses_another_types_matching_amount():
    assert _find_pay_type_row_index(_LIVE_DONURI_ROWS, "급여", 7338970) == 0
    assert _find_pay_type_row_index([], None, 0) is None
    assert _find_pay_type_row_index(_LIVE_DONURI_ROWS, None, 8331530) is None
    assert _find_pay_type_row_index([_LIVE_DONURI_ROWS[0]] * 2, "급여", 8331530) is None


def test_read_saved_total_uses_payment_row_instead_of_coordinate_sums(monkeypatch):
    import asyncio
    from unittest.mock import AsyncMock
    from core import wehago_upload as module
    month = AsyncMock()
    rows = AsyncMock(return_value=[{'fg_pay': '1', 'am_orpay': 36010960}, {'fg_pay': '2', 'am_orpay': 1845610}])
    monkeypatch.setattr(module, '_select_accrual_month', month)
    monkeypatch.setattr(module, '_click_right_side_grid', rows)
    page = object()
    assert asyncio.run(module._read_wehago_net_total(page, 8, '급여')) == 36010960
    month.assert_awaited_once_with(page, 8)


def test_read_saved_total_rejects_missing_amount(monkeypatch):
    import asyncio
    import pytest
    from unittest.mock import AsyncMock
    from core import wehago_upload as module
    monkeypatch.setattr(module, '_select_accrual_month', AsyncMock())
    monkeypatch.setattr(module, '_click_right_side_grid', AsyncMock(return_value=[{'fg_pay': '2'}]))
    with pytest.raises(RuntimeError, match='차인지급액'):
        asyncio.run(module._read_wehago_net_total(object(), 8, None))


def test_month_selection_confirms_only_known_save_prompt(monkeypatch):
    import asyncio
    from unittest.mock import AsyncMock, MagicMock
    from core import wehago_upload as module
    page = MagicMock()
    prompt = MagicMock()
    prompt.count = AsyncMock(return_value=1)
    prompt.first.is_visible = AsyncMock(return_value=True)
    page.get_by_text.return_value = prompt
    page.get_by_role.return_value.click = AsyncMock()
    page.get_by_role.return_value.last.click = AsyncMock()
    page.evaluate = AsyncMock(return_value=[{'left': 1}])
    page.keyboard.press = AsyncMock()
    page.wait_for_timeout = AsyncMock()
    asyncio.run(module._select_accrual_month(page, 8))
    page.get_by_role.return_value.last.click.assert_awaited_once()
    assert page.keyboard.press.await_count == 19


def test_closed_payroll_does_not_execute(monkeypatch):
    import asyncio
    import pytest
    from unittest.mock import AsyncMock, MagicMock
    from core import wehago_upload as module
    monkeypatch.setattr(module, '_click_right_side_grid', AsyncMock(return_value=[{'fg_pay': '2', 'yn_close': '1'}]))
    page = MagicMock()
    with pytest.raises(RuntimeError, match='마감'):
        asyncio.run(module._select_pay_type_and_open(page, None))
    page.get_by_role.assert_not_called()
