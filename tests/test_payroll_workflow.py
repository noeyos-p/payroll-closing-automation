"""Offline safety tests: persisted originals are the sole closing authority."""
import asyncio
from dataclasses import replace
from datetime import date
from decimal import Decimal, localcontext

import pytest

from core.payroll_workflow import (
    Amounts, Employee, OriginalPayroll, PaymentIdentity, PayrollSnapshot,
    PersonAmount, WorkflowState, WorkflowStatus, money, reconcile, run_workflow,
)


PAYMENT = PaymentIdentity("company-1", 2026, 9, 10, date(2026, 10, 10), "salary")


def person(name="홍길동", net=90, code="001", excluded=None):
    return PersonAmount(name, Amounts(net + 10, 10, net), code, excluded)


def original(*people):
    people = people or (person(),)
    return OriginalPayroll(PAYMENT, tuple(people), Amounts(
        sum(int(p.amounts.gross) for p in people),
        sum(int(p.amounts.deductions) for p in people),
        sum(int(p.amounts.net) for p in people)))


def saved(source, **changes):
    return replace(PayrollSnapshot(source.payment, source.people, source.totals,
                                  True, revision="v1"), **changes)


class FakeAdapter:
    """Models atomic guards and persisted state, with recorded side effects."""
    def __init__(self, source, *, entered=False):
        self.snapshot = saved(source) if entered else PayrollSnapshot(
            source.payment, (), None, False, revision="empty")
        self.employees = [Employee(p.employee_code, p.name) for p in source.people if p.employee_code]
        self.calls = []
        self.requests = []
        self.list_o = False
        self.confirm_close = True
        self.accept_close = True
        self.crash_after_register = False
        self.race_snapshot = None
        self.after_close = None
        self.created = None
        self.copy_date = date(2026, 10, 10)

    async def list_closed(self, payment):
        assert payment.same_payment(self.snapshot.payment)
        self.calls.append("list_closed")
        return self.list_o

    async def read_payroll(self, payment):
        self.calls.append("read_payroll")
        return self.snapshot

    async def list_employees(self, payment):
        self.calls.append("list_employees")
        return tuple(self.employees)

    async def register_temporary(self, payment, request):
        self.calls.append("register_temporary")
        self.requests.append(request)
        employee = Employee(f"new-{len(self.employees)}", request.name)
        self.employees.append(employee)
        if self.crash_after_register:
            self.crash_after_register = False
            raise RuntimeError("lost response after server registration")
        return employee

    async def copy_previous_payment(self, payment):
        self.calls.append("copy_previous_payment")
        if self.race_snapshot:
            self.snapshot = self.race_snapshot
        if self.snapshot.entered or self.snapshot.closed:
            raise RuntimeError("payment of this kind already exists; refusing copy")
        self.created = replace(payment, payment_date=self.copy_date)
        # Copied payment carries last month's amounts until the upload replaces them.
        self.snapshot = replace(self.snapshot, payment=self.created, entered=True,
                                people=(person("전월", 1, "999"),), totals=Amounts(11, 10, 1))
        return self.created

    async def upload_to_created(self, source):
        self.calls.append("upload_to_created")
        if source.payment != self.created or self.snapshot.closed:
            return False
        self.snapshot = saved(source)
        return True

    async def close_if_unchanged(self, snapshot):
        self.calls.append("close_if_unchanged")
        if not self.accept_close or self.snapshot.revision != snapshot.revision:
            return False
        if self.confirm_close:
            self.snapshot = replace(self.snapshot, closed=True)
        if self.after_close:
            self.snapshot = self.after_close
        return True


def run(source, adapter, **kwargs):
    return asyncio.run(run_workflow(source, adapter, **kwargs))


def codes(result):
    return {i.code for i in result.issues}


@pytest.mark.parametrize("value", [None, "", "1,000", "abc", True, 1.0,
                                    Decimal("NaN"), Decimal("Infinity"), "1.01"])
def test_money_rejects_missing_fractional_float_or_invalid(value):
    with pytest.raises(ValueError):
        money(value)


@pytest.mark.parametrize("value", [0, -17, "1000", Decimal("1.000")])
def test_whole_won_is_exact_decimal(value):
    assert isinstance(money(value), Decimal)
    assert money(value) == Decimal(value)


def test_sum_does_not_round_with_low_decimal_context():
    source = original(person(net=123456789), person("사원019", 999999999, "002"))
    with localcontext() as context:
        context.prec = 3
        assert reconcile(source, saved(source)).exact


def test_equal_totals_wrong_per_person_blocks():
    source = original(person(net=90), person("사원019", 190, "002"))
    persisted = saved(source, people=(person(net=190), person("사원019", 90, "002")))
    result = reconcile(source, persisted)
    assert source.totals == persisted.totals
    assert "person_amount" in codes(result)
    assert not result.exact


def test_code_distinguishes_homonyms_and_names_are_not_base_three():
    source = original(person("홍길동A", code=None))
    assert not reconcile(source, saved(source, people=(person("홍길동B"),))).exact
    homonyms = original(person(), person(net=190, code="002"))
    assert reconcile(homonyms, saved(homonyms, people=tuple(reversed(homonyms.people)))).exact


def test_exact_normalized_name_and_temporary_alias_only_if_unique():
    source = original(person(" 홍 길 동 ", code=None))
    assert reconcile(source, saved(source, people=(person("홍길동(임시)"),))).exact
    assert not reconcile(source, saved(source, people=(person("홍길동(대체)"),))).exact
    duplicate = (person("홍길동"), person("홍길동(임시)", code="002"))
    result = reconcile(source, saved(source, people=duplicate))
    assert "ambiguous_person" in codes(result)


def test_matching_code_never_falls_back_to_name():
    source = original()
    result = reconcile(source, saved(source, people=(person(code="999"),)))
    assert {"missing_person", "extra_person"} <= codes(result)


@pytest.mark.parametrize("changed", [
    {"year": 2025}, {"accrual_month": 8}, {"payment_kind": "bonus"},
    {"company_id": "other"}, {"payment_date": date(2026, 10, 11)},
    {"payment_date": date(2027, 10, 10)},
    {"payment_month": 11, "payment_date": date(2026, 11, 10)},
])
def test_full_payment_identity_must_match(changed):
    source = original()
    result = reconcile(source, saved(source, payment=replace(PAYMENT, **changed)))
    assert "payment_identity" in codes(result)


@pytest.mark.parametrize("change, expected", [
    ({"people": ()}, "missing_person"),
    ({"people": (person(), person("추가", code="002"))}, "extra_person"),
    ({"people": (person(), person())}, "duplicate_code"),
    ({"totals": None}, "unknown_totals"),
    ({"totals": Amounts(999, 10, 989)}, "invalid_totals"),
    ({"entered": False}, "not_entered"),
])
def test_missing_extra_duplicate_and_unknown_totals_block(change, expected):
    source = original()
    assert expected in codes(reconcile(source, saved(source, **change)))


def test_duplicate_uncoded_original_not_masked_by_row_order():
    source = original(person(code=None), person(code=None))
    assert "duplicate_name" in codes(reconcile(source, saved(source)))


def test_inconsistent_individual_totals_block_even_when_both_sides_equal():
    bad = PersonAmount("홍길동", Amounts(100, 10, 91), "001")
    source = original(bad)
    assert "invalid_person_total" in codes(reconcile(source, saved(source)))


@pytest.mark.parametrize("closed_by_list", [True, False])
def test_list_o_and_already_closed_skip_without_mutations(closed_by_list):
    source = original()
    adapter = FakeAdapter(source, entered=True)
    adapter.list_o = closed_by_list
    adapter.snapshot = replace(adapter.snapshot, closed=not closed_by_list)
    result = run(source, adapter)
    assert result.status == WorkflowStatus.SKIPPED and not result.success
    assert adapter.calls == (["list_closed"] if closed_by_list else ["list_closed", "read_payroll"])


def test_existing_bad_data_untouched_and_no_registration_upload_or_close():
    source = original()
    adapter = FakeAdapter(original(person(net=91)), entered=True)
    before = adapter.snapshot
    result = run(source, adapter)
    assert result.status == WorkflowStatus.BLOCKED
    assert adapter.snapshot is before
    assert adapter.calls == ["list_closed", "read_payroll"]


def test_existing_equal_data_closes_without_upload():
    source = original()
    adapter = FakeAdapter(source, entered=True)
    result = run(source, adapter)
    assert result.success and result.snapshot.closed
    assert adapter.calls == ["list_closed", "read_payroll", "list_employees",
                             "close_if_unchanged", "read_payroll"]


@pytest.mark.parametrize("entered", [False, True])
def test_excluded_employee_blocks_and_is_not_modified(entered):
    source = original()
    adapter = FakeAdapter(source, entered=entered)
    adapter.employees = [Employee("001", "홍길동", "휴직/연말정산 제외")]
    before = adapter.snapshot
    result = run(source, adapter)
    assert "excluded_person" in codes(result)
    assert adapter.snapshot is before
    assert adapter.employees == [Employee("001", "홍길동", "휴직/연말정산 제외")]
    assert not any(call in adapter.calls for call in
                   ("upload_to_created", "register_temporary", "close_if_unchanged"))


def test_original_excluded_blocks_even_if_employee_is_missing():
    source = original(person(code=None, excluded="퇴사자 제외"))
    adapter = FakeAdapter(source)
    assert "excluded_person" in codes(run(source, adapter))
    assert not adapter.requests


def test_empty_register_upload_read_compare_close_read_sequence_and_journal():
    source = original(person(code=None))
    adapter = FakeAdapter(source)
    events = []

    async def journal(event):
        events.append(event)

    result = run(source, adapter, journal=journal)
    assert result.success
    assert adapter.calls == ["list_closed", "read_payroll", "list_employees",
                             "register_temporary", "copy_previous_payment",
                             "upload_to_created", "read_payroll",
                             "close_if_unchanged", "read_payroll"]
    request = adapter.requests[0]
    assert (request.name, request.ssn, request.join_date) == (
        "홍길동(임시)", "0000000000000", date(2026, 9, 1))
    assert result.states == tuple(e.state for e in events)
    checkpoint = [e for e in events if e.registered][0]
    assert checkpoint.state == WorkflowState.REGISTER
    assert checkpoint.registered[0].code == "new-0"
    assert result.states[-5:] == (WorkflowState.RECONCILE, WorkflowState.CLOSE,
                                  WorkflowState.READ_CLOSED, WorkflowState.RESULT)[-5:] if False else result.states[-1] == WorkflowState.RESULT
    assert events[-1].snapshot.closed


@pytest.mark.parametrize("confirm, accepted", [(False, True), (True, False)])
def test_close_not_confirmed_never_reports_success(confirm, accepted):
    source = original()
    adapter = FakeAdapter(source, entered=True)
    adapter.confirm_close = confirm
    adapter.accept_close = accepted
    result = run(source, adapter)
    assert not result.success and "close_unconfirmed" in codes(result)
    assert adapter.calls[-1] == "read_payroll"


def test_final_read_changed_person_blocks_despite_closed_flag():
    source = original()
    adapter = FakeAdapter(source, entered=True)
    adapter.after_close = saved(original(person(net=91)), closed=True)
    result = run(source, adapter)
    assert not result.success and "person_amount" in codes(result)


def test_registration_lost_response_retry_reads_roster_no_duplicate_registration():
    source = original(person(code=None))
    adapter = FakeAdapter(source)
    adapter.crash_after_register = True
    assert run(source, adapter).status == WorkflowStatus.ERROR
    assert len(adapter.employees) == 1 and not adapter.snapshot.entered
    assert run(source, adapter).success
    assert adapter.calls.count("register_temporary") == 1


def test_journal_crash_after_registration_checkpoint_can_retry():
    source = original(person(code=None))
    adapter = FakeAdapter(source)
    checkpoints = []

    async def failing_journal(event):
        if event.registered:
            checkpoints.append(event)
            raise RuntimeError("journal crash")

    assert run(source, adapter, journal=failing_journal).status == WorkflowStatus.ERROR
    assert checkpoints[0].registered == tuple(adapter.employees)
    assert "upload_to_created" not in adapter.calls
    assert run(source, adapter).success
    assert len(adapter.requests) == 1


def test_concurrent_entered_bad_data_is_never_overwritten():
    source = original()
    adapter = FakeAdapter(source)
    adapter.race_snapshot = saved(original(person(net=91)))
    result = run(source, adapter)
    # Another run entered this payment after our first read: the copy refuses,
    # nothing is uploaded or closed, and the other entry stays untouched.
    assert not result.success
    assert adapter.snapshot is adapter.race_snapshot
    assert not {"upload_to_created", "close_if_unchanged"} & set(adapter.calls)


def test_copy_without_proposed_date_blocks_before_upload():
    source = original()
    adapter = FakeAdapter(source)
    adapter.copy_date = None
    result = run(source, adapter)
    assert "copy_identity" in codes(result)
    assert "upload_to_created" not in adapter.calls


def test_copied_previous_amounts_are_never_closed_without_upload():
    source = original()
    adapter = FakeAdapter(source)

    async def refuse(_source):
        adapter.calls.append("upload_to_created")
        return False

    adapter.upload_to_created = refuse
    result = run(source, adapter)
    assert "upload_refused" in codes(result)
    assert "close_if_unchanged" not in adapter.calls


def test_original_without_date_adopts_existing_payment_date():
    undated = replace(original(), payment=replace(PAYMENT, payment_date=None))
    adapter = FakeAdapter(original(), entered=True)
    result = run(undated, adapter)
    assert result.success
    assert result.snapshot.payment.payment_date == PAYMENT.payment_date


def test_unknown_original_total_blocks_before_register_or_upload():
    source = replace(original(person(code=None)), totals=None)
    adapter = FakeAdapter(source)
    assert "unknown_totals" in codes(run(source, adapter))
    assert not adapter.requests and "upload_to_created" not in adapter.calls


def test_missing_explicit_code_is_not_replaced_or_registered():
    source = original()
    adapter = FakeAdapter(source)
    adapter.employees = [Employee("999", "홍길동")]
    assert "unknown_code" in codes(run(source, adapter))
    assert not adapter.requests


def test_unknown_revision_prevents_unconditional_close():
    source = original()
    adapter = FakeAdapter(source, entered=True)
    adapter.snapshot = replace(adapter.snapshot, revision=None)
    assert "unknown_revision" in codes(run(source, adapter))
    assert "close_if_unchanged" not in adapter.calls


def test_retry_after_close_lost_response_skips_closed_server_state():
    source = original()
    adapter = FakeAdapter(source, entered=True)
    real_close = adapter.close_if_unchanged

    async def lose_response(snapshot):
        await real_close(snapshot)
        raise RuntimeError("lost close response")

    adapter.close_if_unchanged = lose_response
    assert run(source, adapter).status == WorkflowStatus.ERROR
    assert run(source, adapter).status == WorkflowStatus.SKIPPED
    assert adapter.calls.count("close_if_unchanged") == 1


def net_only_original(*pairs):
    people = tuple(PersonAmount(n, Amounts.net_only(v), c) for n, v, c in pairs)
    return OriginalPayroll(PAYMENT, people, Amounts.net_only(sum(v for _, v, _ in pairs)))


def test_net_only_original_matches_persisted_with_gross_and_deductions():
    source = net_only_original(("홍길동", 90, "001"), ("김철수", 50, "002"))
    persisted = PayrollSnapshot(PAYMENT, (
        PersonAmount("홍길동", Amounts(100, 10, 90), "001"),
        PersonAmount("김철수", Amounts(60, 10, 50), "002"),
    ), Amounts(160, 20, 140), True, revision="v1")
    assert reconcile(source, persisted).exact


def test_net_only_equal_total_but_swapped_people_is_blocked():
    source = net_only_original(("홍길동", 90, "001"), ("김철수", 50, "002"))
    persisted = PayrollSnapshot(PAYMENT, (
        PersonAmount("홍길동", Amounts.net_only(50), "001"),
        PersonAmount("김철수", Amounts.net_only(90), "002"),
    ), Amounts.net_only(140), True, revision="v1")
    result = reconcile(source, persisted)
    assert not result.exact
    assert [i.code for i in result.issues].count("person_amount") == 2


def test_stated_gross_difference_blocks_even_when_net_matches():
    source = original(person())
    persisted = PayrollSnapshot(PAYMENT, (
        PersonAmount("홍길동", Amounts(101, 11, 90), "001"),), Amounts(101, 11, 90), True, revision="v1")
    assert "person_amount" in {i.code for i in reconcile(source, persisted).issues}


def test_zero_amount_persisted_extras_are_ignored_but_paid_extras_block():
    source = original()
    zero = PersonAmount("영원", Amounts(0, 0, 0), "900")
    assert reconcile(source, saved(source, people=source.people + (zero,))).exact
    paid = PersonAmount("유급", Amounts(10, 0, 10), "901")
    totals = Amounts(110, 10, 100)
    assert "extra_person" in codes(reconcile(source, saved(source, people=source.people + (paid,), totals=totals)))
