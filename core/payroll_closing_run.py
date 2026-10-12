"""Folder → company list → per-kind workflow → H/N update and result rows.

Rules (user, 2026-10-08; see CLAUDE.md):
* A company whose list H is already O is skipped entirely.
* Source files are matched to the list by the exact company name in the file
  name (B column, whitespace-insensitive). No fuzzy matching; an unmatched or
  duplicated name is reported and never written to the list.
* Each payroll kind (급여/급+상/추급) is its own payment and closes on its own.
  A file without a kind in its name is treated as 급여 and says so in the note.
* The original ledger's per-person net (실지급액) is read independently of the
  learned conversion (core/payroll_source_check.py).

The adapter is injected (core/payroll_workflow.PayrollAdapter); this module
never touches a browser. ``write_list`` False keeps the list untouched (dry run).
"""
from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path
import re
import unicodedata
from typing import Callable

from core.payroll_company_list import (
    ClosureResult, CompanyList, CompanyListError, read_company_list, update_company_list,
)
from core.payroll_workflow import (
    Amounts, Issue, OriginalPayroll, PaymentIdentity, PayrollAdapter, PersonAmount,
    WorkflowResult, WorkflowStatus, run_workflow,
)
from core.profile import payroll_types_in_filename

_NAME_RE = re.compile(r"^(?P<company>.+?)\s*(?P<am>\d{1,2})\s*월\s*귀속\s*(?P<pm>\d{1,2})\s*월\s*지급")
_SUFFIXES = (".xlsx", ".xlsm")


def _key(name: str) -> str:
    return "".join(unicodedata.normalize("NFKC", name).split())


@dataclass(frozen=True)
class SourceJob:
    path: Path
    company_name: str
    kind: str
    accrual_year: int
    accrual_month: int
    payment_month: int
    kind_assumed: bool = False


@dataclass
class JobOutcome:
    job: SourceJob | None
    path: Path
    business_number: str | None
    status: str  # closed / skipped / blocked / error / unmatched
    notes: list[str] = field(default_factory=list)
    result: WorkflowResult | None = None


def discover_jobs(folder: Path, payment_year: int, payment_month: int) -> tuple[list[SourceJob], list[JobOutcome]]:
    jobs, problems = [], []
    for path in sorted(Path(folder).iterdir()):
        if not path.is_file() or path.suffix.lower() not in _SUFFIXES or path.name.startswith("~$"):
            continue
        match = _NAME_RE.match(unicodedata.normalize("NFC", path.stem))
        if not match:
            problems.append(JobOutcome(None, path, None, "unmatched", ["파일명에서 업체명/귀속월/지급월을 읽지 못함"]))
            continue
        am, pm = int(match["am"]), int(match["pm"])
        if pm != payment_month or not 1 <= am <= 12:
            problems.append(JobOutcome(None, path, None, "unmatched", [f"지급월 {pm}월: 대상({payment_month}월) 아님"]))
            continue
        kinds = payroll_types_in_filename(path.stem)
        if len(kinds) > 1:
            problems.append(JobOutcome(None, path, None, "blocked",
                                       ["한 파일에 여러 지급구분: 종류별 파일로 나눠야 자동 처리"]))
            continue
        accrual_year = payment_year - 1 if am > pm else payment_year
        jobs.append(SourceJob(path, match["company"].strip(), kinds[0] if kinds else "급여",
                              accrual_year, am, pm, kind_assumed=not kinds))
    return jobs, problems


ReadNet = Callable[[Path], tuple[list, Decimal]]


def build_original(job: SourceJob, business_number: str, read_net: ReadNet) -> tuple[OriginalPayroll, int]:
    """Returns the original and how many people had sub-won net rounded for display."""
    people, total = read_net(job.path)
    payment = PaymentIdentity(business_number, job.accrual_year, job.accrual_month,
                              job.payment_month, None, job.kind)
    rounded = sum(1 for p in people if getattr(p, "raw", None) is not None)
    return OriginalPayroll(payment, tuple(PersonAmount(p.name, Amounts.net_only(p.net)) for p in people),
                           Amounts.net_only(total)), rounded


def _note(outcome: JobOutcome) -> str:
    parts = list(outcome.notes)
    if outcome.result:
        parts += [i.note for i in outcome.result.issues]
    return " / ".join(dict.fromkeys(p for p in parts if p))


async def run_folder(folder: Path, list_path: Path, payment_year: int, payment_month: int,
                     adapter: PayrollAdapter, *, read_net: ReadNet | None = None,
                     write_list: bool = False, journal=None,
                     db_path: Path | None = None) -> tuple[list[JobOutcome], object]:
    """With db_path each job is converted by the registered engine and
    cross-checked against the original (core/payroll_prepare.py); the upload
    plan is handed to the adapter. Without it only the original net is used
    (enough for already-entered payments and for offline tests)."""
    if read_net is None:
        from core.payroll_source_check import read_source_net as read_net
    snapshot: CompanyList = read_company_list(list_path)
    jobs, outcomes = discover_jobs(folder, payment_year, payment_month)
    by_company: dict[str, list[JobOutcome]] = {}
    for job in jobs:
        try:
            row = snapshot.find_company(company_name=job.company_name)
        except CompanyListError:
            matches = [r for r in snapshot.rows if _key(r.company_name) == _key(job.company_name)]
            if len(matches) != 1:
                outcomes.append(JobOutcome(job, job.path, None, "unmatched",
                                           [f"업체리스트에서 '{job.company_name}'을(를) 하나로 찾지 못함({len(matches)}건)"]))
                continue
            row = matches[0]
        if row.business_number is None:
            outcomes.append(JobOutcome(job, job.path, None, "unmatched", ["업체리스트 사업자번호 없음"]))
            continue
        outcome = JobOutcome(job, job.path, row.business_number, "blocked")
        if job.kind_assumed:
            outcome.notes.append("파일명에 지급구분 없음: 급여로 처리")
        if row.completed:
            outcome.status = "skipped"
            outcome.notes.append("업체리스트 이미 O: 건너뜀")
        else:
            try:
                if db_path is not None:
                    from core.payroll_prepare import prepare
                    payment = PaymentIdentity(row.business_number, job.accrual_year, job.accrual_month,
                                              job.payment_month, None, job.kind)
                    prepared = prepare(job.path, payment, db_path=db_path, read_net=read_net)
                    original = prepared.original
                    outcome.notes.extend(prepared.notes)
                    if hasattr(adapter, "register_upload_plan"):
                        adapter.register_upload_plan(original.payment, prepared.plan)
                else:
                    original, rounded = build_original(job, row.business_number, read_net)
                    if rounded:
                        outcome.notes.append(f"원본 실지급액 원 미만 {rounded}명: 엑셀 표시값(반올림)으로 대조")
            except (ValueError, OSError) as exc:
                outcome.notes.append(f"원본 확인 실패: {exc}")
            else:
                result = await run_workflow(original, adapter, journal=journal)
                outcome.result = result
                outcome.status = {WorkflowStatus.CLOSED: "closed", WorkflowStatus.SKIPPED: "skipped",
                                  WorkflowStatus.BLOCKED: "blocked", WorkflowStatus.ERROR: "error"}[result.status]
        outcomes.append(outcome)
        by_company.setdefault(row.business_number, []).append(outcome)

    update = None
    if write_list:
        closures = {}
        for number, items in by_company.items():
            if any(o.status == "skipped" and "이미 O" in " ".join(o.notes) for o in items):
                continue
            # A payment found already closed in WEHAGO counts as closed for the list.
            closures[number] = [ClosureResult(o.job.kind, o.status == "closed" or (
                o.status == "skipped" and o.result is not None
                and any(i.code == "already_closed" for i in o.result.issues)), _note(o)) for o in items]
        if closures:
            update = update_company_list(snapshot, closures)
    return outcomes, update


def outcome_rows(outcomes: list[JobOutcome]) -> list[dict]:
    """Result-table rows: one per source file (person detail lives in the notes)."""
    return [{
        "파일": o.path.name,
        "업체명": o.job.company_name if o.job else "",
        "사업자번호": o.business_number or "",
        "지급구분": o.job.kind if o.job else "",
        "결과": {"closed": "마감", "skipped": "건너뜀", "blocked": "보류", "error": "오류",
                 "unmatched": "연결 안 됨"}[o.status],
        "사유": _note(o),
    } for o in outcomes]
