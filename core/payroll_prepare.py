"""Convert one source ledger into a WEHAGO upload plan and cross-check it.

The existing engine (core.batch.process_one with the registered reference and
learned mappings) produces the upload records. Before anything touches
WEHAGO, every converted person's 차인지급액 must equal the ORIGINAL ledger's
실지급액 read independently (core/payroll_source_check.py), person by person,
with the Excel display rounding the user chose (2026-10-08). Any doubt blocks.

Records the engine could not link to a WEHAGO employee carry the literal code
"사원코드 없음" (core/reconcile.py); only those become temporary registrations.
"""
from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field, replace
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

from core.payroll_workflow import Amounts, OriginalPayroll, PaymentIdentity, PersonAmount, normalize_name

NO_CODE = "사원코드 없음"


class PrepareError(ValueError):
    pass


@dataclass
class UploadPlan:
    """What the browser adapter uploads; codes are filled after registration."""
    profile: object
    records: list
    source_path: Path

    def with_codes(self, codes_by_name: dict[str, str]) -> list:
        """Copy of the records with registered codes put into NO_CODE rows."""
        out = []
        for record in self.records:
            if record.사원코드 == NO_CODE:
                code = codes_by_name.get(normalize_name(record.사원명))
                if not code:
                    raise PrepareError(f"{record.사원명}: 등록된 사원코드가 없어 업로드할 수 없습니다.")
                record = replace(record, 사원코드=code)
            out.append(record)
        return out

    def write(self, path: Path, codes_by_name: dict[str, str]) -> Path:
        from core.export import export_wehago
        export_wehago(self.profile, self.with_codes(codes_by_name), path)
        return path


@dataclass
class Prepared:
    original: OriginalPayroll
    plan: UploadPlan
    notes: list[str] = field(default_factory=list)


def won(value) -> Decimal:
    """Engine floats → whole won exactly as Excel displays (half up)."""
    return Decimal(repr(float(value))).quantize(Decimal(1), rounding=ROUND_HALF_UP)


def cross_check(records: list, source_people: list) -> list[str]:
    """Return problems; empty means every converted net equals a source net.

    Pairs by exact normalized name first (unique on both sides), then requires
    the remaining nets to be identical as a multiset. Counts must agree.
    """
    problems = []
    if len(records) != len(source_people):
        problems.append(f"변환 인원 {len(records)}명 / 원본 인원 {len(source_people)}명")
    by_name: dict[str, list] = {}
    for person in source_people:
        by_name.setdefault(normalize_name(person.name), []).append(person)
    names = Counter(normalize_name(r.사원명) for r in records)
    rest_records, used = [], set()
    for record in records:
        key = normalize_name(record.사원명)
        candidates = by_name.get(key, [])
        if len(candidates) == 1 and names[key] == 1:
            used.add(id(candidates[0]))
            if won(record.차인지급액) != candidates[0].net:
                problems.append(f"{record.사원명}: 변환 차인지급액 {won(record.차인지급액):,} / 원본 실지급액 {candidates[0].net:,}")
        else:
            rest_records.append(record)
    rest_source = Counter(p.net for p in source_people if id(p) not in used)
    rest_converted = Counter(won(r.차인지급액) for r in rest_records)
    if rest_source != rest_converted:
        problems.append("이름으로 짝지을 수 없는 사원들의 실지급액이 원본과 다릅니다.")
    return problems


def net_sheets(path: Path) -> list[str]:
    """Distinct ledger sheets carrying a readable 실지급액 column.

    Sheets with exactly the same people and nets (a calculation sheet and its
    print copy, e.g. 급여계산 / 급여대장-출력만하세요 — observed 2026-10-08) count
    once; nothing is merged. A sheet whose people and nets all appear in another
    sheet is a partial print copy and does not count either (업체006
    급여대장-큰글씨: 5 of 62 people, same nets — 2026-10-12). Sheets with
    different people (24시간 / 36시간) stay separate (→ held)."""
    from collections import Counter
    from core.engine import list_source_sheet_candidates
    from core.payroll_source_check import read_source_net
    sheets: list[tuple[str, Counter]] = []
    for candidate in list_source_sheet_candidates(path):
        try:
            people, _ = read_source_net(path, candidate.name)
        except ValueError:
            continue
        sheets.append((candidate.name, Counter((normalize_name(p.name), p.net) for p in people)))
    names = []
    for i, (name, content) in enumerate(sheets):
        covered = any(j != i and not (content - other) and (other != content or j < i)
                      for j, (_, other) in enumerate(sheets))
        if not covered:
            names.append(name)
    return names


def prepare(path: Path, payment: PaymentIdentity, *, db_path: Path, process=None, read_net=None,
            ledger_sheets=None, reference_path: Path | None = None, client_name: str | None = None,
            company_id: int | None = None) -> Prepared:
    """reference_path: a WEHAGO ledger just exported (Ctrl+G) from the target
    payment; it replaces the stored reference so codes reflect today's roster
    (user's flow, 2026-10-08). Nothing is registered in the local DB.
    company_id: registered company whose saved column mappings apply (without
    it the Ctrl+G path ignored them — 업체022, 2026-10-12)."""
    # User decision 2026-10-08: a file with several ledger sheets (e.g. 24시간/
    # 36시간) is held for a person; sheets are never merged. Helper sheets
    # without a 실지급액 column (장기근속수당, 보수교육자 식비 — observed) do not count.
    if ledger_sheets is None:
        ledger_sheets = net_sheets
    sheets = ledger_sheets(path)
    if len(sheets) > 1:
        raise PrepareError("원본에 급여대장 시트가 여러 개입니다(" + ", ".join(sheets) + "): 사람이 확인 후 처리")
    if process is None and reference_path is not None:
        from core.batch import process_with_reference

        def process(path, db_path):
            return process_with_reference(path, client_name or path.stem, Path(reference_path),
                                          db_path=db_path, auto_register=False, company_id=company_id)
    if process is None:
        from core.batch import process_one as process
    if read_net is None:
        from core.payroll_source_check import read_source_net as read_net
    item = process(path, db_path=db_path)
    if item.status != "matched" or item.result is None:
        raise PrepareError(f"변환 안 됨({item.status}): {item.error or '등록된 사업장/참고파일 확인 필요'}")
    result = item.result
    if result.summary.need_check:
        raise PrepareError(f"이름이 애매해 사원 연결을 확정하지 못한 사원 {result.summary.need_check}명")
    records = list(result.records)
    if not records:
        raise PrepareError("변환된 사원이 없습니다.")
    codes = [r.사원코드 for r in records if r.사원코드 != NO_CODE]
    if any(not str(c).strip() for c in codes) or len(codes) != len(set(codes)):
        raise PrepareError("변환 결과의 사원코드가 비었거나 중복됩니다.")
    source_people, source_total = read_net(path)
    problems = cross_check(records, source_people)
    if problems:
        raise PrepareError("변환 결과가 원본 실지급액과 다릅니다: " + " / ".join(problems))
    # Net only (user rule): WEHAGO's own 지급액계 can exclude some item columns
    # (observed in real exports), so converted gross/deductions are not evidence.
    people = tuple(PersonAmount(r.사원명, Amounts.net_only(won(r.차인지급액)),
                                None if r.사원코드 == NO_CODE else str(r.사원코드)) for r in records)
    original = OriginalPayroll(payment, people, Amounts.net_only(source_total))
    notes = []
    if result.summary.added:
        notes.append(f"사원코드 없음 {result.summary.added}명: 임시 등록 대상")
    if result.summary.removed:
        notes.append(f"위하고 참고명부에만 있는 사원 {result.summary.removed}명은 이번 입력에서 빠짐")
    rounded = sum(1 for p in source_people if getattr(p, "raw", None) is not None)
    if rounded:
        notes.append(f"원본 실지급액 원 미만 {rounded}명: 엑셀 표시값(반올림)으로 대조")
    return Prepared(original, UploadPlan(result.profile, records, Path(path)), notes)
