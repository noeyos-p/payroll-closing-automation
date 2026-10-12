from decimal import Decimal
from pathlib import Path
import pytest
from core.payroll_source_check import strict_won, payment_period, read_source_net


def test_money_does_not_round_or_turn_missing_into_zero():
    assert strict_won('1,001') == Decimal(1001)
    assert strict_won('(1,001)') == Decimal(-1001)
    for value in (None, True, '1.2', 'NaN', '#REF!'):
        with pytest.raises(ValueError): strict_won(value)


def test_accrual_first_day_basis_and_year_rollover():
    assert payment_period(Path('센터 9월귀속 10월지급 급여대장.xlsx'), 2026, 10) == (2026, 9)
    assert payment_period(Path('센터 12월귀속 1월지급 급여대장.xlsx'), 2027, 1) == (2026, 12)
    with pytest.raises(ValueError):
        payment_period(Path('센터 8월귀속 9월지급 급여대장.xlsx'), 2026, 10)


def test_raw_source_net_does_not_use_destination_or_learned_template():
    src = Path(__file__).parent/'fixtures'/'golden'/'sample3-source.xlsx'
    people, total = read_source_net(src)
    assert people and total == sum((p.net for p in people), Decimal(0))
    assert all(p.source == str(src) for p in people)


@pytest.mark.parametrize('fmt', ['#,##0', '#,##0;[Red]\(#,##0\);#,##0', '_-* #,##0_-;-* #,##0_-;_-* "-"_-;_-@_-', '0'])
def test_integer_display_rounds_like_excel(fmt):
    from core.payroll_source_check import displayed_won
    assert displayed_won(310678.5, fmt) == (Decimal(310679), Decimal('310678.5'))
    assert displayed_won(233008.875, fmt)[0] == Decimal(233009)
    assert displayed_won(839613.171195, fmt)[0] == Decimal(839613)
    assert displayed_won(1000, fmt) == (Decimal(1000), None)


@pytest.mark.parametrize('fmt', ['#,##0.00', 'General', None, '0.0'])
def test_decimal_display_is_refused(fmt):
    from core.payroll_source_check import displayed_won
    with pytest.raises(ValueError):
        displayed_won(310678.5, fmt)
