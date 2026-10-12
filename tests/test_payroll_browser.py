import asyncio
from datetime import date, datetime

import openpyxl
import pytest

from core.payroll_browser import (
    LiveBindingRequired, PayrollBrowserError, WehagoPayrollAdapter, check_export_identity,
    employees_from_rows, is_closed, matching_payment_rows, payment_with_date,
)
from core.payroll_workflow import PaymentIdentity, TemporaryRegistration

PAYMENT = PaymentIdentity("1234567890", 2026, 9, 10, None, "급여")

# Shape observed live on 2026-09-22 (values changed).
ROWS = [
    {"fg_pay": "1", "nm_pay": "1", "dt_pay": "20261025", "num": 4, "yn_close": "0", "am_orpay": 10},
    {"fg_pay": "2", "nm_pay": "2", "dt_pay": "20261025", "num": 3, "yn_close": "1", "am_orpay": 20},
    {"fg_pay": "5", "nm_pay": "5", "dt_pay": "20261010", "num": 1, "yn_close": "0", "am_orpay": 30},
    {"fg_pay": "1", "nm_pay": "1", "dt_pay": "20260925", "num": 4, "yn_close": "1", "am_orpay": 40},
]


def test_rows_are_selected_by_kind_code_and_payment_month():
    assert matching_payment_rows(ROWS, PAYMENT) == [0]
    assert matching_payment_rows(ROWS, PaymentIdentity("1234567890", 2026, 9, 10, None, "추급")) == [2]
    assert matching_payment_rows(ROWS, PaymentIdentity("1234567890", 2026, 9, 11, None, "급여")) == []
    with pytest.raises(PayrollBrowserError):
        matching_payment_rows([{"fg_pay": "1"}], PAYMENT)


def test_row_date_and_close_flag():
    assert payment_with_date(PAYMENT, ROWS[0]).payment_date == date(2026, 10, 25)
    assert not is_closed(ROWS[0]) and is_closed(ROWS[1])
    with pytest.raises(PayrollBrowserError):
        is_closed({"yn_close": "Y"})


def test_export_identity_checks_period_and_company():
    check_export_identity("가나다센터 - 202609 (1).xlsx", "가나다 센터", PAYMENT)
    check_export_identity("가나다센터(2026) - 202609.xlsx", "가나다센터", PAYMENT)
    with pytest.raises(PayrollBrowserError, match="귀속월"):
        check_export_identity("가나다센터 - 202608.xlsx", "가나다센터", PAYMENT)
    with pytest.raises(PayrollBrowserError, match="업체명"):
        check_export_identity("다른센터 - 202609.xlsx", "가나다센터", PAYMENT)


def test_roster_rows_keep_codes_skip_blank_and_mark_retired():
    rows = [{"cd_emp": "0023", "nm_krname": "홍길동", "da_retire": None},
            {"cd_emp": "01", "nm_krname": "김철수", "da_retire": "20260831"},
            {"cd_emp": "7", "nm_krname": "사원015", "da_retire": "20260915"},
            {"cd_emp": None, "nm_krname": None, "da_retire": None}]
    employees = {e.code: e for e in employees_from_rows(rows, date(2026, 9, 1))}
    assert set(employees) == {"0023", "01", "7"}
    assert employees["01"].excluded_reason and not employees["7"].excluded_reason
    with pytest.raises(PayrollBrowserError):
        employees_from_rows(rows + [{"cd_emp": "0023", "nm_krname": "중복", "da_retire": None}], date(2026, 9, 1))
    with pytest.raises(PayrollBrowserError):
        employees_from_rows([{"code": "1"}], date(2026, 9, 1))


def test_read_only_adapter_refuses_every_mutation(tmp_path):
    adapter = WehagoPayrollAdapter(None, {"1234567890": "가나다센터"}, tmp_path)

    async def attempts():
        for call in (adapter.copy_previous_payment(PAYMENT),
                     adapter.register_temporary(PAYMENT, TemporaryRegistration("가(임시)", date(2026, 9, 1))),
                     adapter.close_if_unchanged(None)):
            with pytest.raises(PayrollBrowserError, match="읽기 전용"):
                await call

    asyncio.run(attempts())


def test_unobserved_steps_never_guess_even_when_mutation_allowed(tmp_path):
    adapter = WehagoPayrollAdapter(None, {}, tmp_path, allow_mutation=True)

    async def attempts():
        with pytest.raises(LiveBindingRequired):
            await adapter.copy_previous_payment(PAYMENT)

    asyncio.run(attempts())


def test_upload_refused_for_payment_not_created_by_this_run(tmp_path):
    from core.payroll_workflow import Amounts, OriginalPayroll
    adapter = WehagoPayrollAdapter(None, {}, tmp_path, allow_mutation=True)
    original = OriginalPayroll(PaymentIdentity("1234567890", 2026, 9, 10, date(2026, 10, 25), "급여"),
                               (), Amounts.net_only(0))
    assert asyncio.run(adapter.upload_to_created(original)) is False


def test_choose_company_row_by_number_and_year():
    from core.payroll_browser import choose_company_row
    rows = [{"company": f"업체011 ({y})", "businessNumber": "100-00-00099", "anchor": f"tooltip_{y}"}
            for y in (2022, 2025, 2026)]
    rows.append({"company": "업체011 (2026)", "businessNumber": "111-11-11111", "anchor": "x"})
    assert choose_company_row(rows, "1000000099", 2026)["anchor"] == "tooltip_2026"
    with pytest.raises(PayrollBrowserError, match="2022, 2025, 2026"):
        choose_company_row(rows, "1000000099", 2024)
    with pytest.raises(PayrollBrowserError, match="찾지 못"):
        choose_company_row(rows, "9999999999", 2026)
