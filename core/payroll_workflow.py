"""Pure payroll comparison and an injected, retryable closing workflow.

Adapters must read persisted data, never a screen preview. A payment that
already exists is this month's entry: it is compared, never overwritten (user
decision 2026-10-08). Only a payment the adapter itself created in this run
via ``copy_previous_payment`` (WEHAGO 복사 of last month's same kind, whose
proposed payment date is adopted) may be filled by ``upload_to_created``.
``close_if_unchanged`` must compare the supplied persisted revision and refuse
changed/closed data. There
are deliberately no employment, leave, year-end, overwrite or reopen methods.
Journal storage belongs to the caller. Persist events before returning from the
async callback; retry always re-reads server state, even after a lost response.
"""
from __future__ import annotations

from dataclasses import dataclass, replace
from datetime import date
from decimal import Decimal, InvalidOperation
from enum import Enum
import unicodedata
from typing import Awaitable, Callable, Protocol


def money(value: Decimal | int | str) -> Decimal:
    """Parse exact whole won. Missing values, floats and rounding are forbidden."""
    if isinstance(value, bool) or not isinstance(value, (Decimal, int, str)):
        raise ValueError("금액은 누락 없이 Decimal, 정수 또는 숫자 문자열이어야 합니다")
    try:
        result = Decimal(value)
    except (InvalidOperation, ValueError):
        raise ValueError("유효하지 않은 금액") from None
    if not result.is_finite() or result != result.to_integral_value():
        raise ValueError("금액은 반올림 없는 정수 원 단위여야 합니다")
    return result


def normalize_name(name: str) -> str:
    """Keep the entire name (including birth digits and meaningful parentheses)."""
    return "".join(unicodedata.normalize("NFKC", name).split())


def _alias(name: str) -> str:
    key = normalize_name(name)
    return key[:-4] if key.endswith("(임시)") else key


@dataclass(frozen=True)
class PaymentIdentity:
    company_id: str
    year: int  # Accrual year; payment_date can fall in the following year.
    accrual_month: int
    payment_month: int
    # None until WEHAGO states it: an existing payment row, or the date WEHAGO
    # proposes when copying last month's payment.
    payment_date: date | None
    payment_kind: str  # Explicit adapter kind/code, never inferred from amounts.

    def __post_init__(self) -> None:
        if not self.company_id.strip() or not self.payment_kind.strip():
            raise ValueError("사업장과 지급구분이 필요합니다")
        date(self.year, self.accrual_month, 1)
        if not 1 <= self.payment_month <= 12 or (
                self.payment_date is not None and self.payment_date.month != self.payment_month):
            raise ValueError("지급월과 지급일자가 일치해야 합니다")

    @property
    def key(self) -> tuple:
        return (self.company_id, self.year, self.accrual_month, self.payment_month, self.payment_kind)

    def same_payment(self, other: "PaymentIdentity") -> bool:
        """Same key; dates must agree whenever both sides state one."""
        return self.key == other.key and (
            self.payment_date is None or other.payment_date is None
            or self.payment_date == other.payment_date)

    @property
    def join_date(self) -> date:
        return date(self.year, self.accrual_month, 1)


@dataclass(frozen=True)
class Amounts:
    """Net (실지급액) is mandatory: it is the user's closing criterion.

    Gross/deductions are optional (None = not read, NEVER zero). The original
    ledger reader only yields net, so they are compared only when both sides
    state them.
    """
    gross: Decimal | None
    deductions: Decimal | None
    net: Decimal

    def __post_init__(self) -> None:
        object.__setattr__(self, "net", money(self.net))
        for key in ("gross", "deductions"):
            if getattr(self, key) is not None:
                object.__setattr__(self, key, money(getattr(self, key)))

    @classmethod
    def net_only(cls, net: Decimal | int | str) -> "Amounts":
        return cls(None, None, net)

    @property
    def consistent(self) -> bool:
        if self.gross is None or self.deductions is None:
            return True
        return int(self.gross) - int(self.deductions) == int(self.net)

    def matches(self, other: "Amounts") -> bool:
        """Exact net, plus gross/deductions wherever BOTH sides stated them."""
        if self.net != other.net:
            return False
        return all(getattr(self, key) is None or getattr(other, key) is None
                   or getattr(self, key) == getattr(other, key)
                   for key in ("gross", "deductions"))


@dataclass(frozen=True)
class PersonAmount:
    name: str
    amounts: Amounts
    employee_code: str | None = None
    excluded_reason: str | None = None

    def __post_init__(self) -> None:
        if not normalize_name(self.name):
            raise ValueError("원본 직원의 정확한 이름이 필요합니다")
        if self.employee_code is not None:
            object.__setattr__(self, "employee_code", self.employee_code.strip() or None)


@dataclass(frozen=True)
class OriginalPayroll:
    payment: PaymentIdentity
    people: tuple[PersonAmount, ...]
    totals: Amounts | None  # None means unknown, NEVER zero or a computed substitute.

    def __post_init__(self) -> None:
        object.__setattr__(self, "people", tuple(self.people))


@dataclass(frozen=True)
class PayrollSnapshot:
    payment: PaymentIdentity
    people: tuple[PersonAmount, ...]
    totals: Amounts | None
    entered: bool
    closed: bool = False
    revision: str | None = None

    def __post_init__(self) -> None:
        object.__setattr__(self, "people", tuple(self.people))


@dataclass(frozen=True)
class Issue:
    code: str
    note: str


@dataclass(frozen=True)
class Reconciliation:
    issues: tuple[Issue, ...]
    matched_count: int = 0

    @property
    def exact(self) -> bool:
        return not self.issues


def _sum(people: tuple[PersonAmount, ...]) -> Amounts:
    # Integer arithmetic avoids Decimal's ambient precision silently rounding sums.
    # A component is summed only when every person states it; otherwise unknown.
    def total(key: str) -> Decimal | None:
        values = [getattr(p.amounts, key) for p in people]
        if key != "net" and any(v is None for v in values):
            return None
        return Decimal(sum(int(v) for v in values))
    return Amounts(total("gross"), total("deductions"), total("net"))


def _validate(people: tuple[PersonAmount, ...], totals: Amounts | None,
              label: str) -> list[Issue]:
    issues: list[Issue] = []
    if not people:
        issues.append(Issue("empty_people", f"{label}: 직원 내역 없음"))
    codes = [p.employee_code for p in people if p.employee_code]
    if len(codes) != len(set(codes)):
        issues.append(Issue("duplicate_code", f"{label}: 중복 사원코드"))
    aliases: dict[str, list[PersonAmount]] = {}
    for person in people:
        aliases.setdefault(_alias(person.name), []).append(person)
        if person.excluded_reason:
            issues.append(Issue("excluded_person", f"{label}: {person.name}: {person.excluded_reason}"))
        if not person.amounts.consistent:
            issues.append(Issue("invalid_person_total", f"{label}: {person.name}: 지급-공제≠실수령"))
    for group in aliases.values():
        # Codes distinguish real homonyms; any uncoded duplicate remains ambiguous.
        if len(group) > 1 and any(not p.employee_code for p in group):
            issues.append(Issue("duplicate_name", f"{label}: 이름/임시 별칭 중복 {group[0].name}"))
    if totals is None:
        issues.append(Issue("unknown_totals", f"{label}: 원본/저장 합계 미확인"))
    elif not totals.consistent or not totals.matches(_sum(people)):
        issues.append(Issue("invalid_totals", f"{label}: 개인별 합계와 명시 합계 불일치"))
    return issues


def _candidates(person: PersonAmount, others: tuple[PersonAmount, ...]) -> list[int]:
    if person.employee_code:
        return [i for i, other in enumerate(others) if other.employee_code == person.employee_code]
    # Alias uniqueness is checked over ALL rows, not merely unmatched rows.
    return [i for i, other in enumerate(others) if _alias(other.name) == _alias(person.name)]


def _all_zero(person: PersonAmount) -> bool:
    a = person.amounts
    return a.net == 0 and a.gross in (None, 0) and a.deductions in (None, 0)


def reconcile(original: OriginalPayroll, persisted: PayrollSnapshot) -> Reconciliation:
    """Compare originals to persisted people AND all three stated totals.

    A persisted person absent from the original whose amounts are all zero is
    a payment-list member with nothing paid; it is ignored (user decision
    2026-10-08, observed live: 4 such people in a manually closed payment)."""
    issues = _validate(original.people, original.totals, "원본")
    issues += _validate(persisted.people, persisted.totals, "위하고 저장")
    if not original.payment.same_payment(persisted.payment) or persisted.payment.payment_date is None:
        issues.append(Issue("payment_identity", "사업장/연도/귀속월/지급월/일자/지급구분 불일치"))
    if not persisted.entered:
        issues.append(Issue("not_entered", "저장된 급여 입력 확인 안 됨"))
    used: set[int] = set()
    for person in original.people:
        candidates = _candidates(person, persisted.people)
        if len(candidates) != 1:
            issues.append(Issue("missing_person" if not candidates else "ambiguous_person",
                                f"{person.name}: 개인 식별 불가 ({len(candidates)}건)"))
            continue
        index = candidates[0]
        if index in used:
            issues.append(Issue("duplicate_match", f"{person.name}: 동일 직원 재매칭"))
            continue
        used.add(index)
        if not person.amounts.matches(persisted.people[index].amounts):
            issues.append(Issue("person_amount", f"{person.name}: 원본 {person.amounts}, 저장 {persisted.people[index].amounts}"))
    for index, person in enumerate(persisted.people):
        if index not in used and not _all_zero(person):
            issues.append(Issue("extra_person", f"저장 내역에만 있는 직원: {person.name}"))
    if (original.totals is None or persisted.totals is None
            or not original.totals.matches(persisted.totals)):
        issues.append(Issue("total_amount", "원본과 위하고 저장 총액 불일치"))
    return Reconciliation(tuple(issues), len(used))


@dataclass(frozen=True)
class Employee:
    code: str
    name: str
    excluded_reason: str | None = None

    def __post_init__(self) -> None:
        if not self.code.strip() or not normalize_name(self.name):
            raise ValueError("사원 목록에 코드와 정확한 이름이 필요합니다")
        object.__setattr__(self, "code", self.code.strip())


@dataclass(frozen=True)
class TemporaryRegistration:
    name: str
    join_date: date
    ssn: str = "0000000000000"


class PayrollAdapter(Protocol):
    async def list_closed(self, payment: PaymentIdentity) -> bool: ...
    async def read_payroll(self, payment: PaymentIdentity) -> PayrollSnapshot: ...
    async def list_employees(self, payment: PaymentIdentity) -> tuple[Employee, ...]: ...
    async def register_temporary(self, payment: PaymentIdentity,
                                 request: TemporaryRegistration) -> Employee: ...
    async def copy_previous_payment(self, payment: PaymentIdentity) -> PaymentIdentity: ...
    async def upload_to_created(self, original: OriginalPayroll) -> bool: ...
    async def close_if_unchanged(self, persisted: PayrollSnapshot) -> bool: ...


class WorkflowState(str, Enum):
    LIST_CLOSED = "preflight_list_closed"
    EXISTING_CLOSED = "existing_closed"
    EXISTING_ENTERED = "existing_entered_compare_never_upload"
    CODE_CHECKS = "missing_code_checks"
    REGISTER = "register_temporary"
    CREATE_PAYMENT = "copy_previous_payment"
    UPLOAD = "upload_to_created"
    READ_PERSIST = "read_persist"
    RECONCILE = "reconcile"
    CLOSE = "close"
    READ_CLOSED = "read_closed"
    RESULT = "result"


class WorkflowStatus(str, Enum):
    SKIPPED = "skipped"
    BLOCKED = "blocked"
    CLOSED = "closed"
    ERROR = "error"


@dataclass(frozen=True)
class JournalEvent:
    state: WorkflowState
    payment: PaymentIdentity
    snapshot: PayrollSnapshot | None = None
    registered: tuple[Employee, ...] = ()
    note: str = ""


Journal = Callable[[JournalEvent], Awaitable[None]]


@dataclass(frozen=True)
class WorkflowResult:
    status: WorkflowStatus
    payment: PaymentIdentity
    issues: tuple[Issue, ...]
    states: tuple[WorkflowState, ...]
    snapshot: PayrollSnapshot | None = None

    @property
    def success(self) -> bool:
        return self.status == WorkflowStatus.CLOSED


async def run_workflow(original: OriginalPayroll, adapter: PayrollAdapter, *,
                       journal: Journal | None = None) -> WorkflowResult:
    """Skip closed payments; compare entered data; upload only an empty payment.

    A crash/cancellation may leave a completed registration/upload/close on the
    server. No saved local state is trusted on retry. Exceptions return ERROR;
    cancellation propagates. Journal failure prevents the next mutation.
    """
    states: list[WorkflowState] = []
    snapshot: PayrollSnapshot | None = None
    registered: list[Employee] = []
    comparison_original = original

    async def step(state: WorkflowState, note: str = "") -> None:
        states.append(state)
        if journal:
            await journal(JournalEvent(state, original.payment, snapshot, tuple(registered), note))

    async def finish(status: WorkflowStatus, issues: tuple[Issue, ...] = ()) -> WorkflowResult:
        await step(WorkflowState.RESULT, status.value)
        return WorkflowResult(status, original.payment, issues, tuple(states), snapshot)

    try:
        await step(WorkflowState.LIST_CLOSED)
        if await adapter.list_closed(original.payment):
            return await finish(WorkflowStatus.SKIPPED, (Issue("list_closed", "회사 목록 O/마감 표시: 건너뜀"),))
        await step(WorkflowState.EXISTING_CLOSED)
        snapshot = await adapter.read_payroll(original.payment)
        if not original.payment.same_payment(snapshot.payment) or (
                snapshot.entered and snapshot.payment.payment_date is None):
            return await finish(WorkflowStatus.BLOCKED, (Issue("payment_identity", "조회한 지급건 식별 불일치"),))
        if snapshot.closed:
            return await finish(WorkflowStatus.SKIPPED, (Issue("already_closed", "이미 마감된 지급건: 건너뜀"),))
        await step(WorkflowState.EXISTING_ENTERED)
        if snapshot.entered:
            # Existing payment = this month's entry: compare only, never overwrite.
            comparison_original = replace(original, payment=snapshot.payment)
            await step(WorkflowState.RECONCILE)
            comparison = reconcile(comparison_original, snapshot)
            if not comparison.exact:
                return await finish(WorkflowStatus.BLOCKED, comparison.issues)
            await step(WorkflowState.CODE_CHECKS)
            employees = tuple(await adapter.list_employees(original.payment))
            issues = []
            for person in snapshot.people:
                matches = [e for e in employees if
                           (e.code == person.employee_code if person.employee_code
                            else _alias(e.name) == _alias(person.name))]
                if len(matches) != 1:
                    issues.append(Issue("ambiguous_employee", f"{person.name}: 사원 명부 확인 불가"))
                elif matches[0].excluded_reason:
                    issues.append(Issue("excluded_person", f"{person.name}: {matches[0].excluded_reason}"))
            if issues:
                return await finish(WorkflowStatus.BLOCKED, tuple(issues))
        else:
            await step(WorkflowState.CODE_CHECKS)
            issues = _validate(original.people, original.totals, "원본")
            if snapshot.people or snapshot.totals is not None:
                issues.append(Issue("uncertain_empty", "미입력 표시지만 저장 내역/합계 존재"))
            employees = tuple(await adapter.list_employees(original.payment))
            if len({e.code for e in employees}) != len(employees):
                issues.append(Issue("duplicate_code", "사원 목록의 코드 중복"))
            resolved: list[PersonAmount] = []
            missing: list[PersonAmount] = []
            for person in original.people:
                matches = [e for e in employees if
                           (e.code == person.employee_code if person.employee_code
                            else _alias(e.name) == _alias(person.name))]
                if len(matches) > 1:
                    issues.append(Issue("ambiguous_employee", f"{person.name}: 사원 목록 중복"))
                elif matches:
                    employee = matches[0]
                    if employee.excluded_reason:
                        issues.append(Issue("excluded_person", f"{person.name}: {employee.excluded_reason}"))
                    resolved.append(replace(person, employee_code=employee.code))
                elif person.employee_code:
                    issues.append(Issue("unknown_code", f"{person.name}: 지정 코드 {person.employee_code} 조회 안 됨"))
                else:
                    missing.append(person)
            if issues:
                return await finish(WorkflowStatus.BLOCKED, tuple(issues))
            for person in missing:
                await step(WorkflowState.REGISTER, f"등록 요청: {person.name}")
                request = TemporaryRegistration(_alias(person.name) + "(임시)", original.payment.join_date)
                employee = await adapter.register_temporary(original.payment, request)
                if (normalize_name(employee.name) != normalize_name(request.name)
                        or employee.excluded_reason or employee.code in {e.code for e in employees}
                        or employee.code in {e.code for e in registered}):
                    return await finish(WorkflowStatus.BLOCKED, (Issue("registration_identity", "임시등록 응답 식별/제외 상태 오류"),))
                registered.append(employee)
                # Persist completed registration BEFORE upload; server roster also
                # recovers a crash between registration and this callback.
                await step(WorkflowState.REGISTER, f"등록 확인: {employee.code}")
                resolved.append(replace(person, employee_code=employee.code))
            upload = replace(original, people=tuple(resolved))
            issues = _validate(upload.people, upload.totals, "사원코드 확정 원본")
            if issues:
                return await finish(WorkflowStatus.BLOCKED, tuple(issues))
            await step(WorkflowState.CREATE_PAYMENT)
            created = await adapter.copy_previous_payment(original.payment)
            if not original.payment.same_payment(created) or created.payment_date is None:
                return await finish(WorkflowStatus.BLOCKED, (Issue(
                    "copy_identity", "전월 복사로 만든 지급건의 업체/귀속/지급월/종류/지급일 확인 실패"),))
            upload = replace(upload, payment=created)
            comparison_original = upload
            await step(WorkflowState.CREATE_PAYMENT, f"복사 지급일 {created.payment_date.isoformat()}")
            await step(WorkflowState.UPLOAD)
            if not await adapter.upload_to_created(upload):
                return await finish(WorkflowStatus.BLOCKED, (Issue(
                    "upload_refused", "복사로 만든 지급건이 아니거나 상태가 바뀌어 업로드하지 않음"),))
            await step(WorkflowState.READ_PERSIST)
            snapshot = await adapter.read_payroll(created)
            await step(WorkflowState.RECONCILE)
            comparison = reconcile(upload, snapshot)
            if not comparison.exact:
                return await finish(WorkflowStatus.BLOCKED, comparison.issues)
        if snapshot.closed:
            return await finish(WorkflowStatus.SKIPPED, (Issue("already_closed", "다른 실행에서 이미 마감됨"),))
        if not snapshot.revision:
            return await finish(WorkflowStatus.BLOCKED, (Issue("unknown_revision", "저장 버전 미확인: 안전한 마감 불가"),))
        await step(WorkflowState.CLOSE)
        accepted = await adapter.close_if_unchanged(snapshot)
        await step(WorkflowState.READ_CLOSED)
        confirmed = await adapter.read_payroll(comparison_original.payment)
        # Require both exact final persisted data and an independently read flag.
        final_comparison = reconcile(comparison_original, confirmed)
        snapshot = confirmed
        if not accepted or not confirmed.closed or not final_comparison.exact:
            return await finish(WorkflowStatus.BLOCKED, final_comparison.issues +
                                (Issue("close_unconfirmed", "마감 후 저장 내역/마감 상태 확인 실패"),))
        return await finish(WorkflowStatus.CLOSED)
    except Exception as exc:
        # A journal outage cannot produce success, even if a mutation completed.
        states.append(WorkflowState.RESULT)
        return WorkflowResult(WorkflowStatus.ERROR, original.payment,
                              (Issue("adapter_or_journal_error", f"{type(exc).__name__}: {exc}"),),
                              tuple(states), snapshot)
