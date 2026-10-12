from decimal import Decimal
from pathlib import Path
from types import SimpleNamespace

import pytest

from core.payroll_prepare import NO_CODE, PrepareError, UploadPlan, cross_check, prepare, won
from core.payroll_source_check import SourceNet
from core.payroll_workflow import PaymentIdentity
from core.wehago_schema import WehagoRecord

PAYMENT = PaymentIdentity("1234567890", 2026, 9, 10, None, "급여")


def rec(code, name, net):
    return WehagoRecord(사원코드=code, 사원명=name, 지급액계=net + 10, 공제액계=10, 차인지급액=net)


def src(name, net):
    return SourceNet(name, Decimal(net), "x", 1)


def item(records, need_check=0, added=0, status="matched"):
    summary = SimpleNamespace(need_check=need_check, added=added, removed=0)
    result = SimpleNamespace(records=records, summary=summary, profile=None)
    return SimpleNamespace(status=status, result=result, error=None)


def run(records, sources, **kwargs):
    total = sum((p.net for p in sources), Decimal(0))
    return prepare(Path("a.xlsx"), PAYMENT, db_path=Path("db"),
                   process=lambda path, db_path: item(records, **kwargs),
                   read_net=lambda path: (sources, total), ledger_sheets=lambda path: ["Sheet1"])


def test_won_rounds_half_up_like_excel():
    assert won(310678.5) == Decimal(310679) and won(233008.875) == Decimal(233009)
    assert won(330121.92) == Decimal(330122)


def test_matched_codes_and_no_code_rows_become_registration_targets():
    prepared = run([rec("001", "홍길동", 100), rec(NO_CODE, "김신규", 50)],
                   [src("홍길동", 100), src("김신규", 50)], added=1)
    codes = {p.name: p.employee_code for p in prepared.original.people}
    assert codes == {"홍길동": "001", "김신규": None}
    assert prepared.original.totals.net == Decimal(150)
    assert all(p.amounts.gross is None for p in prepared.original.people)
    assert any("임시 등록" in n for n in prepared.notes)


def test_per_person_net_difference_blocks_even_with_equal_total():
    with pytest.raises(PrepareError, match="홍길동"):
        run([rec("001", "홍길동", 90), rec("002", "김철수", 60)], [src("홍길동", 100), src("김철수", 50)])


def test_count_mismatch_and_ambiguous_names_block():
    with pytest.raises(PrepareError, match="인원"):
        run([rec("001", "홍길동", 100)], [src("홍길동", 100), src("김철수", 0)])
    with pytest.raises(PrepareError, match="애매"):
        run([rec("001", "홍길동", 100)], [src("홍길동", 100)], need_check=1)


def test_unconverted_or_duplicate_codes_block():
    with pytest.raises(PrepareError, match="변환 안 됨"):
        run([], [], status="needs_reference")
    with pytest.raises(PrepareError, match="중복"):
        run([rec("001", "가", 1), rec("001", "나", 2)], [src("가", 1), src("나", 2)])


def test_unpaired_names_need_identical_net_multiset():
    # Ledger name differs from WEHAGO name (e.g. suffix), but nets agree.
    assert not cross_check([rec("001", "홍길동A", 100)], [src("홍길동", 100)])
    assert cross_check([rec("001", "홍길동A", 101)], [src("홍길동", 100)])


def test_upload_plan_fills_registered_codes_and_refuses_missing(tmp_path):
    plan = UploadPlan(None, [rec("001", "홍길동", 100), rec(NO_CODE, "김 신규", 50)], Path("a.xlsx"))
    filled = plan.with_codes({"김신규": "0099"})
    assert [r.사원코드 for r in filled] == ["001", "0099"]
    assert plan.records[1].사원코드 == NO_CODE  # original plan untouched
    with pytest.raises(PrepareError):
        plan.with_codes({})


def test_several_ledger_sheets_are_held_for_a_person():
    with pytest.raises(PrepareError, match="24시간, 36시간"):
        prepare(Path("a.xlsx"), PAYMENT, db_path=Path("db"), process=None, read_net=None,
                ledger_sheets=lambda path: ["24시간", "36시간"])
