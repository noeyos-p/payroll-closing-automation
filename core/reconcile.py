"""Main reconciliation engine — port of the VBA macro's
`Sub 급여대장_위하고_현재파일_변환` body (파일/이지케어_Sheet1_현재시트_수정본_v3.bas).
Given an already-loaded 이지케어 원본 worksheet + its detected
`SourceLayout`, and an existing 위하고 Sheet1 worksheet + its detected
`DestLayout`, walks the destination's registered employees, resolves each
to a source row by name (core/name_match.py), updates their amounts from
the source using header-text canonicalization (core/canon.py,
core/classify.py), flags employees the destination has but the source
doesn't (removed from the result) or the source has but the destination
doesn't (added as new rows), and produces three outputs mirroring the
macro's 정리본/검증결과/항목매핑결과 sheets."""

from __future__ import annotations

from dataclasses import dataclass, field

from openpyxl.utils import get_column_letter

import re

from core.canon import _canon_key, canon, is_ambiguous_alias, norm
from core.classify import (
    _EMPLOYER_BURDEN_KEYWORDS,
    canon_of_dst,
    canon_of_src,
    group_of_dst,
    group_of_src,
    is_source_employer_burden_col,
    is_source_info_col,
    is_source_work_time_col,
)
from core.gridutil import cell_text
from core.name_match import (
    SourceNameIndex,
    name_base3,
    person_birth_key,
    resolve_source_row,
    should_add_source_row,
)
from core.profile import ClientProfile, now_iso
from core.sheet_layout import (
    DestLayout,
    SourceLayout,
    dst_full_header,
    dst_item,
    find_total_row,
    last_nonblank_row,
    source_full_header,
    source_item,
)
from core.target_schema import ColumnDef
from core.totals import compute_taxable_split, compute_totals
from core.wehago_schema import WehagoRecord

_VALUE_TOLERANCE = 1
# 정답 파일에서 사업장별 매칭을 거꾸로 알아낼 때(core/mapping_learn.py)만
# 리스트로 바꿔 끼운다 — 그 동안 reconcile()이 원본/위하고 시트와 사원별
# 원본 행 짝, 최종 열 매칭을 여기 남긴다. 평소(None)에는 아무것도 안 함.
_LEARN_SINK: list | None = None
_TOTAL_CANONS = {"지급액계", "공제액계", "차인지급액"}
_TAXABLE_SUBTOTAL_KEYS = {"과세계", "비과세계", "과세", "비과세"}


@dataclass
class VerificationRow:
    name: str
    status: str  # 정상/차이/추가됨/추가됨_차이/삭제됨/확인필요
    src_pay: float | None = None
    dst_pay: float | None = None
    diff_pay: float | None = None
    src_ded: float | None = None
    dst_ded: float | None = None
    diff_ded: float | None = None
    src_net: float | None = None
    dst_net: float | None = None
    diff_net: float | None = None
    memo: str = ""


@dataclass
class MappingRow:
    kind: str  # "자동매칭" | "수동매칭" | "미매칭"
    source_header: str
    dest_column_label: str = ""
    note: str = ""
    source_col: int = 0  # source column index — lets the UI re-run with a manual override for this column
    dest_key: str | None = None  # the raw destination record key (unformatted) — None when unmapped
    ambiguous: bool = False  # matched via core.canon.is_ambiguous_alias — UI shows a yellow "확인 필요" flag
    # `source_item()`(그룹 접두어 없는 최하위 항목명) — 화면 표시에는 그룹까지
    # 포함한 `source_header`를 쓰지만, `core/registry.py`의 사업장별 저장
    # 매칭 키로는 이걸 쓴다. 같은 항목이라도 파일/달마다 그룹 헤더 행이
    # 있다 없다 하면 `source_header`(전체 경로)가 달라져서 저장된 매칭을
    # 다시 못 찾는 버그가 실사에서 확인됨(2026-09-08) — item_label은 그룹
    # 유무와 무관하게 항상 같다.
    item_label: str = ""
    # 이 소스 열의 수당/공제 그룹(group_of_src) — 저장 매칭 키에 항목명과
    # 같이 들어간다. "성과급"이 수당 쪽에도 공제 쪽에도 각각 존재하면서
    # 서로 다른 위하고 열을 가리켜야 하는 실제 사례가 있어서(업체001어르신
    # 주야간보호센터, 2026-09-08), 그룹 없이 항목명만 키로 쓰면 한쪽을
    # 고치면 다른 쪽까지 같이 틀어져 버렸다.
    group: str = ""


def learned_mapping_key(group: str, item_label: str) -> str:
    """`core/registry.py`의 사업장별 저장 매칭 조회/저장 키(정규화 전
    원문) — 그룹까지 포함해서 만든다. 저장(`ui/main_window.py`)과 조회
    (아래 `reconcile()`) 둘 다 이 함수를 거쳐야 서로 어긋나지 않는다.
    구분자는 "|"(core.canon.norm이 지우지 않는 문자) — ":"를 처음 썼다가
    norm()이 그걸 지워버려서 "수당"+"지급>..." 식으로 그냥 붙어버리는
    걸 실사 중 발견, "|"로 교체함(2026-09-08)."""
    return f"{group}|{item_label}"


@dataclass
class ReconcileSummary:
    matched: int = 0
    added: int = 0
    removed: int = 0
    diff: int = 0
    need_check: int = 0


@dataclass
class ReconcileResult:
    profile: ClientProfile
    records: list[WehagoRecord]
    verification_rows: list[VerificationRow] = field(default_factory=list)
    mapping_rows: list[MappingRow] = field(default_factory=list)
    summary: ReconcileSummary = field(default_factory=ReconcileSummary)
    # Destination record keys (WehagoRecord.allowances/deductions keys) whose
    # value came from an ambiguous-alias auto-match — ui/review_table.py
    # highlights these columns yellow so they're visible on the main review
    # screen too, not only inside 검증결과 보기's 항목매핑결과 탭.
    ambiguous_dest_keys: set[str] = field(default_factory=set)
    # Whole-run check, independent of per-employee matching: the source
    # ledger's OWN printed "합계" row for 지급액계/공제액계/차인지급액 (only
    # populated when that row and column were both actually found) versus
    # the sum of every record this run produced. `_apply_source_totals_check`
    # below only ever runs for a row that resolved to a specific source
    # employee — an employee dropped as 삭제됨, or one left 확인필요 because
    # name/생년월일 matching failed, silently skips that check entirely, so a
    # real difference there was previously invisible anywhere in the app.
    # Comparing the ledger's own grand total instead catches exactly that
    # case (사용자 요청, 2026-08-13).
    grand_total_source: dict[str, float] = field(default_factory=dict)
    grand_total_computed: dict[str, float] = field(default_factory=dict)
    grand_total_mismatches: set[str] = field(default_factory=set)
    # 화면에 보여줄 처리 메모(예: "원본을 수정본 양식으로 변환해서 처리했습니다").
    notes: list[str] = field(default_factory=list)


_FOUND_STATUSES = {"정상", "차이", "추가됨", "추가됨_차이"}


def _sum_records(records: list[WehagoRecord], non_taxable_keys: set[str]) -> WehagoRecord:
    """같은 사원(이름 기준)이 여러 시트에 나뉘어 있을 때(예: 추급1/추급2에
    둘 다 근무 기록이 있는 경우) 항목별 금액을 더해서 합친 레코드 하나로
    만든다 — 사용자 요청(2026-09-08): "동명이인에 생년월일 같으면 같은
    사람으로 보고 합쳐서". 각 개별 실행이 이미 이름+생년월일로 목적지
    사원을 정확히 찾아준 뒤라, 여기서는 그 결과(사원명이 항상 목적지
    시트에서 읽은 동일 텍스트)를 이름으로 다시 묶기만 하면 된다."""
    base = records[0]
    allowances: dict[str, float] = dict(base.allowances)
    deductions: dict[str, float] = dict(base.deductions)
    for record in records[1:]:
        for key, value in record.allowances.items():
            allowances[key] = allowances.get(key, 0.0) + value
        for key, value in record.deductions.items():
            deductions[key] = deductions.get(key, 0.0) + value
    지급액계, 공제액계, 차인지급액 = compute_totals(allowances, deductions)
    비과세합계, 과세합계 = compute_taxable_split(allowances, non_taxable_keys)
    return WehagoRecord(
        사원코드=base.사원코드,
        사원명=base.사원명,
        부서=base.부서,
        직급=base.직급,
        직종=base.직종,
        allowances=allowances,
        deductions=deductions,
        비과세합계=비과세합계,
        과세합계=과세합계,
        지급액계=지급액계,
        공제액계=공제액계,
        차인지급액=차인지급액,
        # 여러 시트를 합친 값이라 원본 자체 합계(각 시트 하나만의 값)와
        # 대조할 단일 기준이 없다 — 합친 뒤에는 원본 대조 표시를 비운다.
    )


def _sum_optional(values: list[float | None]) -> float | None:
    present = [v for v in values if v is not None]
    return sum(present) if present else None


def _merge_verification_rows(rows: list[VerificationRow]) -> VerificationRow:
    if len(rows) == 1:
        return rows[0]
    base = rows[0]
    src_pay = _sum_optional([r.src_pay for r in rows])
    dst_pay = _sum_optional([r.dst_pay for r in rows])
    src_ded = _sum_optional([r.src_ded for r in rows])
    dst_ded = _sum_optional([r.dst_ded for r in rows])
    src_net = _sum_optional([r.src_net for r in rows])
    dst_net = _sum_optional([r.dst_net for r in rows])
    diff_pay = (src_pay - dst_pay) if src_pay is not None and dst_pay is not None else None
    diff_ded = (src_ded - dst_ded) if src_ded is not None and dst_ded is not None else None
    diff_net = (src_net - dst_net) if src_net is not None and dst_net is not None else None
    has_diff = any(r.status in ("차이", "추가됨_차이") for r in rows)
    is_added = any(r.status.startswith("추가됨") for r in rows)
    status = ("추가됨_차이" if has_diff else "추가됨") if is_added else ("차이" if has_diff else "정상")
    memo = " / ".join(dict.fromkeys(r.memo for r in rows if r.memo))
    return VerificationRow(
        name=base.name, status=status,
        src_pay=src_pay, dst_pay=dst_pay, diff_pay=diff_pay,
        src_ded=src_ded, dst_ded=dst_ded, diff_ded=diff_ded,
        src_net=src_net, dst_net=dst_net, diff_net=diff_net,
        memo=memo,
    )


def merge_reconcile_results(results: list[ReconcileResult]) -> ReconcileResult:
    """추급1/추급2처럼 같은 종류로 나뉜 여러 시트를 각각 독립적으로
    reconcile()한 결과를 하나로 합친다(사용자 요청, 2026-09-08 — 서로 다른
    사람이 나뉘어 있는 게 보통이지만, 동명이인에 생년월일이 같으면 같은
    사람으로 보고 항목별 금액을 더한다).

    각 개별 실행은 위하고에 등록된 사원 전체를 대상으로 돌아서, 이번
    시트에 없는 사원은 전부 "삭제됨"으로 나온다 — 이건 진짜 삭제가 아니라
    "이 시트엔 없다"는 뜻이므로, 합칠 때는 이름별로 "데이터가 있는(정상/
    차이/추가됨) 결과가 하나라도 있으면 그걸 쓰고, 전부 삭제됨이어야만
    진짜 삭제됨"으로 판단한다. 같은 이름이 여러 실행에서 데이터를 가지고
    있으면(진짜 같은 사람이 두 시트에 다 나온 경우) 항목별로 합산한다."""
    if len(results) == 1:
        return results[0]
    if not results:
        raise ValueError("합칠 결과가 없습니다.")

    base = results[0]
    non_taxable_keys = {c.key for c in base.profile.allowance_columns if not c.taxable}

    records_by_name: dict[str, list[WehagoRecord]] = {}
    name_order: list[str] = []
    for result in results:
        for record in result.records:
            if record.사원명 not in records_by_name:
                name_order.append(record.사원명)
            records_by_name.setdefault(record.사원명, []).append(record)
    merged_records = [
        records_by_name[name][0] if len(records_by_name[name]) == 1 else _sum_records(records_by_name[name], non_taxable_keys)
        for name in name_order
    ]

    verif_by_name: dict[str, list[VerificationRow]] = {}
    verif_order: list[str] = []
    for result in results:
        for row in result.verification_rows:
            if row.name not in verif_by_name:
                verif_order.append(row.name)
            verif_by_name.setdefault(row.name, []).append(row)
    merged_verification_rows: list[VerificationRow] = []
    for name in verif_order:
        rows = verif_by_name[name]
        found = [r for r in rows if r.status in _FOUND_STATUSES]
        if found:
            merged_verification_rows.append(_merge_verification_rows(found))
            continue
        need_check = [r for r in rows if r.status == "확인필요"]
        merged_verification_rows.append(need_check[0] if need_check else rows[0])

    summary = ReconcileSummary()
    for row in merged_verification_rows:
        if row.status == "삭제됨":
            summary.removed += 1
        elif row.status == "확인필요":
            summary.need_check += 1
        elif row.status in ("정상", "차이"):
            summary.matched += 1
            if row.status == "차이":
                summary.diff += 1
        else:  # 추가됨 / 추가됨_차이
            summary.added += 1
            if row.status == "추가됨_차이":
                summary.diff += 1

    mapping_rows: list[MappingRow] = []
    ambiguous_dest_keys: set[str] = set()
    for result in results:
        mapping_rows.extend(result.mapping_rows)
        ambiguous_dest_keys |= result.ambiguous_dest_keys

    grand_total_source: dict[str, float] = {}
    grand_total_computed: dict[str, float] = {}
    for result in results:
        for key, value in result.grand_total_source.items():
            grand_total_source[key] = grand_total_source.get(key, 0.0) + value
        for key, value in result.grand_total_computed.items():
            grand_total_computed[key] = grand_total_computed.get(key, 0.0) + value
    grand_total_mismatches = {
        key for key, stated in grand_total_source.items()
        if abs(grand_total_computed.get(key, 0.0) - stated) > _VALUE_TOLERANCE
    }

    return ReconcileResult(
        profile=base.profile,
        records=merged_records,
        verification_rows=merged_verification_rows,
        mapping_rows=mapping_rows,
        summary=summary,
        ambiguous_dest_keys=ambiguous_dest_keys,
        grand_total_source=grand_total_source,
        grand_total_computed=grand_total_computed,
        grand_total_mismatches=grand_total_mismatches,
    )


def _as_number(value: object) -> float | None:
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        s = value.replace(",", "").replace("원", "").strip()
        if s in ("", "-"):
            return None
        try:
            return float(s)
        except ValueError:
            return None
    return None


def _disambiguate(labels: list[str]) -> list[str]:
    seen: dict[str, int] = {}
    result = []
    for label in labels:
        seen[label] = seen.get(label, 0) + 1
        result.append(label if seen[label] == 1 else f"{label}({seen[label]})")
    return result


# --- destination column classification ------------------------------------


def build_dst_cols(ws, layout: DestLayout) -> dict[str, list[int]]:
    """Port of VBA `BuildDstCols` — canonical item name -> destination
    column(s) currently registered for it (in column order)."""
    cols: dict[str, list[int]] = {}
    for c in range(1, ws.max_column + 1):
        group = group_of_dst(ws, layout, c)
        ca = canon_of_dst(ws, layout, c)
        if group in ("수당", "공제", "차인지급액") and ca:
            cols.setdefault(ca, []).append(c)
    return cols


def _dst_column_schema(
    ws, layout: DestLayout, dst_cols: dict[str, list[int]]
) -> tuple[list[ColumnDef], list[ColumnDef], dict[int, str]]:
    """Builds (allowance_columns, deduction_columns, col->record-key) using
    each destination column's OWN existing label — this app never renames a
    client's registered Wehago columns, it only fills their amounts in.

    Labels are disambiguated across BOTH groups together, not separately —
    a client can register the exact same raw label in an allowance column
    and a deduction column (confirmed real case: 업체013's Sheet1
    has an allowance column literally labeled "식대" and a *deduction*
    column also literally labeled "식대", right next to a third column
    labeled "식대공제"). `WehagoRecord.column_value`/`_apply_amounts_to_record`
    both resolve a key by checking `allowances` before `deductions`, so if
    the two groups produced the same key text independently, the deduction
    column's amount would silently be added to (and only ever displayed as
    part of) the allowance one instead — real money invisibly inflating
    지급액계 and leaving 공제액계 short by the same amount. Disambiguating
    over the combined, column-order-sorted list guarantees every key is
    unique across the whole record, so that never happens."""
    allowance: list[tuple[int, str]] = []
    deduction: list[tuple[int, str]] = []
    for ca, cols in dst_cols.items():
        if ca in _TOTAL_CANONS:
            continue
        for c in cols:
            group = group_of_dst(ws, layout, c)
            label = dst_item(ws, layout, c) or ca
            if group == "수당":
                allowance.append((c, label))
            elif group == "공제":
                deduction.append((c, label))
    allowance.sort(key=lambda pair: pair[0])
    deduction.sort(key=lambda pair: pair[0])

    combined = sorted(allowance + deduction, key=lambda pair: pair[0])
    combined_labels = _disambiguate([label for _, label in combined])
    label_by_col = {c: label for (c, _raw), label in zip(combined, combined_labels)}

    col_to_key: dict[int, str] = {}
    allowance_columns: list[ColumnDef] = []
    for c, _raw in allowance:
        label = label_by_col[c]
        full = dst_full_header(ws, layout, c)
        raw_item = dst_item(ws, layout, c)
        taxable = "비과세" not in norm(full) and "(비)" not in raw_item
        allowance_columns.append(ColumnDef(label, label, "allowance", taxable=taxable))
        col_to_key[c] = label

    deduction_columns: list[ColumnDef] = []
    for c, _raw in deduction:
        label = label_by_col[c]
        deduction_columns.append(ColumnDef(label, label, "deduction"))
        col_to_key[c] = label

    return allowance_columns, deduction_columns, col_to_key


# --- source -> destination column mapping ----------------------------------


def build_dst_text_index(ws, layout: DestLayout) -> list[tuple[int, str, str]]:
    """(dest col, group, normalized item text) for every 수당/공제 destination
    column — `build_map`'s exact-text and fallback matching use it."""
    index: list[tuple[int, str, str]] = []
    for c in range(1, ws.max_column + 1):
        group = group_of_dst(ws, layout, c)
        if group not in ("수당", "공제"):
            continue
        if canon_of_dst(ws, layout, c) in _TOTAL_CANONS:
            continue
        key = _text_key(dst_item(ws, layout, c))
        if key:
            index.append((c, group, key))
    return index


def _text_key(text: object) -> str:
    """Item text normalized for literal comparison — same text pipeline as
    canon() (tax marks/rate/share suffixes stripped) but WITHOUT the alias
    table, so "식대비" stays "식대비" here."""
    return _canon_key(text)


_NUMBER_SUFFIX_RE = re.compile(r"\d+$")


def _fallback_dest_col(
    ws, layout: SourceLayout, c: int, group: str, dst_cols: dict[str, list[int]],
    dst_index: list[tuple[int, str, str]],
) -> int | None:
    """Last-chance matching for a source column the exact-text and canon
    passes both left unmapped. Each rule was derived from the user's own
    hand-corrected 9월지급 uploads (작업파일\\Exel, 2026-10-07 대조):
      1. "기본급 > 기본수당", "연차수당 > 포함" — the leaf label is a generic
         word, but a header ABOVE it names a real item the destination has.
      2. "기본급(식대 미포함)", "기본급및주휴수당", "과세 (기본급여)" 등 —
         the destination's label (or canon) is the start of the source's
         longer label. Longest such label wins.
      3. "회비" → dest "회비공제" (공제 group only).
      4. "추가공제1" → dest "추가공제" (trailing number dropped).
    Matches made here are always surfaced as "확인 필요" (ambiguous) — they
    are reasonable but not certain."""
    same_group = [(col, key) for col, g, key in dst_index if g == group]
    if not same_group:
        return None
    same_cols = {col for col, _ in same_group}
    item_key = _text_key(source_item(ws, layout, c))
    if not item_key:
        return None
    # 2. destination label/canon is a prefix of the source label
    best = None
    for col, key in same_group:
        if len(key) >= 2 and item_key.startswith(key) and key != item_key:
            if best is None or len(key) > len(best[1]):
                best = (col, key)
    for dest_ca, cols in dst_cols.items():
        if len(dest_ca) >= 2 and dest_ca not in _TOTAL_CANONS and item_key.startswith(dest_ca) and dest_ca != item_key:
            for col in cols:
                if col in same_cols and (best is None or len(dest_ca) > len(best[1])):
                    best = (col, dest_ca)
    # "과세(기본급여)"처럼 앞에 과세 구분이 붙은 경우
    for prefix in ("과세", "비과세"):
        if best is None and item_key.startswith(prefix):
            rest = item_key[len(prefix):]
            rest_ca = canon(rest)
            for col in dst_cols.get(rest_ca, []):
                if col in same_cols:
                    best = (col, rest_ca)
                    break
    if best is not None:
        return best[0]
    # 3. "+공제" — "회비"→"회비공제", "본인부담"→"본인부담금공제"
    if group == "공제":
        for col, key in same_group:
            if key == item_key + "공제":
                return col
        hits = [col for col, key in same_group if key.endswith("공제") and key.startswith(item_key) and len(item_key) >= 2]
        if len(hits) == 1:
            return hits[0]
    # 4. trailing number
    stripped = _NUMBER_SUFFIX_RE.sub("", item_key)
    if stripped and stripped != item_key:
        for col, key in same_group:
            if key == stripped:
                return col
    # 5. 위하고 열 이름이 원본 이름을 통째로 품고 있는 경우가 딱 하나면
    #    ("성과격려수당" → "명절성과격려수당", "우수직원" → "근면성실 우수직원")
    if len(item_key) >= 3:
        hits = [col for col, key in same_group if item_key in key and key != item_key]
        if len(hits) == 1:
            return hits[0]
    return None


_SINGLE_SOURCE_CANONS = {"국민연금", "건강보험", "고용보험", "장기요양보험료", "소득세", "지방소득세"}


def _looks_like_money_column(ws, layout: SourceLayout, c: int) -> bool:
    """False for columns whose numbers can't be a month's pay amount: 0/1
    flags (or rates like 0.045) and code-like numbers of 10억 or more
    (2026-10-07 업체041's helper columns: 0/1 가입여부 flags,
    10001000101-style 코드). Blank/text-only columns stay eligible so
    they're still listed as usual."""
    nonzero = []
    last = min(ws.max_row, layout.data_start_row + 300)
    for r in range(layout.data_start_row, last + 1):
        n = _as_number(ws.cell(row=r, column=c).value)
        if n:
            nonzero.append(abs(n))
    if not nonzero:
        return True
    if max(nonzero) >= 1e9:
        return False
    if max(nonzero) <= 1:
        return False
    return True


def build_map(
    ws, layout: SourceLayout, dst_cols: dict[str, list[int]],
    dst_index: list[tuple[int, str, str]] | None = None,
) -> tuple[dict[int, int], dict[str, int], set[int]]:
    """Port of VBA `BuildMap` — which source column's amount accumulates
    into which destination column. Multiple source columns can point at the
    same destination column (they're summed, mirroring `ApplyMap`'s `+=`)
    when a client's Wehago sheet registers fewer numbered variants
    (기타수당/기타수당2/...) than the source ledger has.

    `dst_index` (from `build_dst_text_index`) enables two additions on top
    of the VBA port (2026-10-07, measured against the user's hand-corrected
    9월지급 uploads): a source column canon/group can't classify still maps
    when its text is literally a destination column's ("배상보험"), and
    anything still unmapped after the canon pass gets `_fallback_dest_col`'s
    narrower rules (flagged ambiguous).

    The third return value is the set of source columns whose match relied
    on an `is_ambiguous_alias` label (core/canon.py) — matched with the same
    confidence as any other auto-match, but surfaced to the UI as a yellow
    "확인 필요" flag rather than a plain 자동매칭, since the alias can mean a
    different thing at a different client (사용자 요청, 2026-08-07)."""
    seq: dict[str, int] = {}
    mapping: dict[int, int] = {}
    unmapped: dict[str, int] = {}
    ambiguous: set[int] = set()
    dst_index = dst_index or []

    candidates: list[tuple[int, str, str]] = []  # (src col, group, canon)
    for c in range(1, ws.max_column + 1):
        if is_source_work_time_col(ws, layout, c):
            continue
        if is_source_employer_burden_col(ws, layout, c):
            # 회사 부담분으로 보이는 열이라도 위하고에 글자 그대로 같은 이름의
            # 열이 있으면 그 사업장은 실제로 그 금액을 급여에서 처리하는 것이다
            # (2026-10-07 업체053한사랑/원주한사랑: 원본 "사업자분 공제" → 위하고
            # "사업자분공제" 열 — 사용자가 올린 정답에 그대로 들어가 있었음).
            # 단, 회사부담 표현이 맨 아래 칸 자체에 있을 때만("사업자분 공제") —
            # "기관부담금 > 국민연금"의 '국민연금'이 위하고 국민연금 열과 같다고
            # 붙이면 공제가 두 번 들어간다(골든 표본 sample3로 확인).
            key = _text_key(source_item(ws, layout, c))
            if key and any(kw in norm(source_item(ws, layout, c)) for kw in _EMPLOYER_BURDEN_KEYWORDS) and _looks_like_money_column(ws, layout, c):
                exact = [col for col, _g, k in dst_index if k == key]
                if len(exact) == 1:
                    mapping[c] = exact[0]
            continue
        if is_source_info_col(ws, layout, c):
            continue

        item_key = norm(source_item(ws, layout, c))
        # "과세급여총액" (or similarly named taxable-only subtotal columns) is
        # a subtotal like 계/과세계/비과세계, not a real line item — confirmed
        # real case: 업체013's 원본 has this column right next to the
        # real 지급총액 column; without excluding it, it gets treated as an
        # ordinary (unmapped, harmless but noisy) allowance item.
        if item_key in ("계", "과세계", "비과세계", "과세급여총액", "비과세급여총액"):
            continue

        group = group_of_src(ws, layout, c)
        if group not in ("수당", "공제"):
            # A column canon/group can't classify can still carry a label
            # the destination literally has (2026-10-07: 업체016
            # "배상보험", 업체022 "장기근속장려금소급분") — take the
            # destination's own group when the text is an exact match.
            if group != "차인지급액" and _looks_like_money_column(ws, layout, c):
                key = _text_key(source_item(ws, layout, c))
                exact = [col for col, _g, k in dst_index if k == key]
                if key and len(exact) == 1:
                    mapping[c] = exact[0]
                    continue
            # group_of_src couldn't tell 수당 from 공제 for this column (no
            # group-row text above it and the item name itself isn't in
            # canon.py's dictionary) — rather than silently dropping it with
            # no trace anywhere, surface it as 미매칭 so the accountant can
            # see it exists and connect it by hand (사용자 요청, 2026-08-12:
            # "일단 무조건 뜨게해봐" — a column the app can't classify should
            # still be visible, never invisible). "차인지급액"-group columns
            # are still skipped: those are real-total columns handled
            # separately by `_find_src_pay_total_col`, not per-item mapping.
            if group != "차인지급액":
                full = source_full_header(ws, layout, c)
                if full:
                    unmapped[f"{full} [{get_column_letter(c)}열]"] = c
            continue

        ca = canon_of_src(ws, layout, c)
        if not ca or ca in _TOTAL_CANONS:
            continue
        if not _looks_like_money_column(ws, layout, c):
            continue
        candidates.append((c, group, ca))

    # 같은 canon의 위하고 열이 여러 개(기타수당/기타수당2, 휴일근로수당/
    # 휴일근로수당(요보))면 원본 열 순서대로 배정하는 VBA 규칙을 그대로 쓴다.
    # 2026-10-07에 "글자 완전일치 우선"으로 바꿔봤더니 정답 파일 12곳이
    # 뒤집혀서(참고파일 자체가 이 순서 규칙으로 만들어졌기 때문) 되돌렸다.
    used_exact: set[int] = set()
    seen_single: set[str] = set()
    for c, group, ca in candidates:
        if c in mapping:
            continue
        # No fallback substitution to a conceptually-different canon: a
        # source item only auto-maps to a destination column that shares its
        # OWN exact canon. A previous version redirected e.g. 근로자의날 to
        # a 휴일근로수당 column, or 기타공제-canon items to a 추가공제
        # column, when the exact one was missing — this silently merged
        # items that are "similar but not the same" (confirmed real
        # complaint: 두루누리 items ending up folded into 기타공제). If the
        # destination has no column with this exact canon, the item is left
        # unmapped and surfaces in 항목매핑결과 for the user to pick
        # manually (사용자 요청, 2026-08-06). A 공제-group column always
        # subtracts from 지급액계 (compute_totals never treats a deduction
        # as a credit) — so any recognized 공제 canon with a matching
        # destination column is safe to auto-map, same as 수당.
        if ca in _SINGLE_SOURCE_CANONS and ca in seen_single:
            # 4대보험/세금은 한 사람에게 한 번만 걷는다 — 같은 이름 열이 또
            # 나오면(환급/기준금액/계산용 보조표) 더하지 않고 미매칭으로 남긴다
            # (2026-10-07 업체041: 오른쪽 계산표의 "국민/건강/
            # 장기/고용" 열 수십 개가 전부 공제에 더해지고 있었음).
            full = source_full_header(ws, layout, c)
            if full:
                unmapped[f"{full} [{get_column_letter(c)}열]"] = c
            continue
        if ca in _SINGLE_SOURCE_CANONS:
            seen_single.add(ca)
        all_cols = dst_cols.get(ca) or []
        if dst_index:
            group_cols = {col for col, g, _k in dst_index if g == group}
            all_cols = [col for col in all_cols if col in group_cols] or all_cols
        cols_for_ca = [col for col in all_cols if col not in used_exact] or all_cols
        dc = None
        if cols_for_ca:
            seq[ca] = seq.get(ca, 0) + 1
            idx = seq[ca]
            dc = cols_for_ca[idx - 1] if idx - 1 < len(cols_for_ca) else cols_for_ca[0]

        if dc is None and dst_index:
            dc = _fallback_dest_col(ws, layout, c, group, dst_cols, dst_index)
            if dc is not None:
                single_ca = next((k for k, cols in dst_cols.items() if dc in cols and k in _SINGLE_SOURCE_CANONS), None)
                if single_ca is not None and single_ca in seen_single:
                    dc = None
                elif single_ca is not None:
                    seen_single.add(single_ca)
            if dc is not None:
                ambiguous.add(c)

        if dc is not None:
            mapping[c] = dc
            if is_ambiguous_alias(source_item(ws, layout, c)):
                ambiguous.add(c)
        else:
            full = source_full_header(ws, layout, c)
            if full:
                unmapped[f"{full} [{get_column_letter(c)}열]"] = c

    return mapping, unmapped, ambiguous


def _find_src_pay_total_col(ws, layout: SourceLayout) -> int | None:
    """Port of VBA `FindSrcPayTotalCol` (3-tier fallback)."""

    def skip(c: int) -> bool:
        return (
            is_source_work_time_col(ws, layout, c)
            or is_source_employer_burden_col(ws, layout, c)
            or is_source_info_col(ws, layout, c)
        )

    for c in range(1, ws.max_column + 1):
        if skip(c):
            continue
        top_key = norm(cell_text(ws, layout.header_rows[0], c))
        item_key = norm(source_item(ws, layout, c))
        if top_key == "지급" and item_key in ("지급액", "지급금액"):
            return c

    for c in range(1, ws.max_column + 1):
        if skip(c):
            continue
        item_key = norm(source_item(ws, layout, c))
        full_key = norm(source_full_header(ws, layout, c))
        if item_key in _TAXABLE_SUBTOTAL_KEYS:
            continue
        if "과세계" in full_key or "비과세계" in full_key:
            continue
        if any(kw in full_key for kw in ("실지급액", "실수령액", "차인지급액")):
            continue
        if item_key in ("지급총액", "지급합계", "소득합계", "급여총액", "지급액계"):
            return c

    for c in range(1, ws.max_column + 1):
        item_key = norm(source_item(ws, layout, c))
        full_key = norm(source_full_header(ws, layout, c))
        if item_key in _TAXABLE_SUBTOTAL_KEYS or "과세계" in full_key or "비과세계" in full_key:
            continue
        if canon_of_src(ws, layout, c) == "지급액계":
            return c

    # Last resort: a bare "지급액" column with nothing more specific found —
    # confirmed real case: 업체006's 원본 has a "지급액" column
    # positioned as the gross-pay subtotal right before 공제내역 starts, with
    # no "지급" super-header above it (so tier 1 above doesn't catch it) and
    # no "계"/"총액"/"합계" suffix (so tier 2 doesn't either). "지급액" is
    # also registered as a 차인지급액(net) alias (`canon.py`, ported from
    # VBA) for files where it genuinely means that instead — this only
    # engages once every stronger, gross-specific pattern has already failed.
    for c in range(1, ws.max_column + 1):
        if skip(c):
            continue
        item_key = norm(source_item(ws, layout, c))
        full_key = norm(source_full_header(ws, layout, c))
        if item_key == "지급액" and not any(kw in full_key for kw in ("실지급액", "실수령액", "차인지급액")):
            return c
    return None


# A column literally titled "실지급액"/"실수령액"/etc. means net pay
# unambiguously; a bare "지급액"/"지급급여"/"지급금액" (also aliased to
# 차인지급액 in canon.py, ported from VBA) is genuinely ambiguous — some
# files use it for net pay, but 업체006's 원본 uses it as the
# GROSS subtotal positioned right before deductions, with the real net pay
# in a separate "실지급액" column further right. When both exist, prefer the
# unambiguous one.
_STRONG_NET_PAY_TERMS = {"실지급액", "실수령액", "차인지급액", "차감지급액"}


def _find_src_canon_col(ws, layout: SourceLayout, target_canon: str) -> int | None:
    if target_canon == "차인지급액":
        matches = [c for c in range(1, ws.max_column + 1) if canon_of_src(ws, layout, c) == target_canon]
        strong = [c for c in matches if norm(source_item(ws, layout, c)) in _STRONG_NET_PAY_TERMS]
        if strong:
            return strong[0]
        return matches[0] if matches else None

    for c in range(1, ws.max_column + 1):
        if canon_of_src(ws, layout, c) == target_canon:
            return c
    return None


# --- per-row amount computation --------------------------------------------


def _block_number(
    ws, start_row: int, block_height: int, col: int | None, extra_rows: tuple[int, ...] | list[int] = ()
) -> float | None:
    """Sums a column across one employee's whole block of physical rows
    (SourceLayout.block_height) instead of just its first row — a
    continuation row's own numeric cells (e.g. this app's own added
    columns for a client whose ledger splits identity fields onto a second
    row) belong to the same employee, not a separate one."""
    if col is None:
        return None
    total = None
    for r in [*range(start_row, start_row + block_height), *extra_rows]:
        n = _as_number(ws.cell(row=r, column=col).value)
        if n is not None:
            total = (total or 0.0) + n
    return total


def _row_amounts(
    ws, row: int, block_height: int, column_map: dict[int, int], extra_rows: tuple[int, ...] | list[int] = ()
) -> dict[int, float]:
    values: dict[int, float] = {}
    for src_col, dst_col in column_map.items():
        n = _block_number(ws, row, block_height, src_col, extra_rows)
        if n:
            values[dst_col] = values.get(dst_col, 0.0) + n
    return values


def _row_has_total_label(ws, r: int) -> bool:
    for c in range(1, min(ws.max_column, 8) + 1):
        v = ws.cell(row=r, column=c).value
        if isinstance(v, str) and ("계" in v or "합" in v or "총" in v):
            return True
    return False


def _row_has_text(ws, r: int, first_col: int, last_col: int) -> bool:
    """이름 없는 합계 줄(숫자만 있음)과 수급자별 보조 줄(수급자명/급여내용
    같은 글자가 있음)을 구분한다."""
    for c in range(first_col, min(ws.max_column, last_col) + 1):
        v = ws.cell(row=r, column=c).value
        if isinstance(v, str) and v.strip() and _as_number(v) is None:
            return True
    return False


def _row_has_number(ws, r: int, start_col: int) -> bool:
    for c in range(start_col, min(ws.max_column, 80) + 1):
        n = _as_number(ws.cell(row=r, column=c).value)
        if n:
            return True
    return False


def _continuation_rows(ws, layout: SourceLayout, total_row: int) -> dict[int, list[int]]:
    """직원 한 명의 금액이 보조 줄에 이어지는 경우 — 이름 칸이 "(목)"처럼
    괄호로 시작하고 순번 칸이 비어 있는 줄은 바로 위 직원의 금액에 더한다
    (2026-10-07 업체054: 사원009 아래 "(목)" 줄의 기본급 51,600이
    사용자가 올린 정답에 사원009 기본급으로 합산돼 있었음)."""
    extra: dict[int, list[int]] = {}
    if layout.block_height != 1:
        return extra
    current = None
    for r in range(layout.data_start_row, total_row):
        name = ws.cell(row=r, column=layout.name_col).value
        text = str(name).strip() if name is not None else ""
        marker_blank = (
            layout.marker_col == layout.name_col
            or ws.cell(row=r, column=layout.marker_col).value in (None, "")
        )
        if text.startswith(("(", "（")) and marker_blank and current is not None:
            extra.setdefault(current, []).append(r)
            continue
        # 순번 열이 따로 있는 표에서 순번·이름이 모두 빈데 금액이 있는 줄 —
        # 한 직원이 수급자별로 여러 줄을 쓰는 양식(2026-10-07 업체055
        # 센터: 사원010 아래 김성섭 수급자 줄 742,000원이 정답에선 사원010
        # 급여에 합산돼 있었음). "소계"/"합계" 같은 글자가 있는 줄은 제외.
        if (
            not text
            and current is not None
            and layout.has_marker
            and layout.marker_col != layout.name_col
            and ws.cell(row=r, column=layout.marker_col).value in (None, "")
            and not _row_has_total_label(ws, r)
            and _row_has_text(ws, r, layout.name_col + 1, layout.name_col + 8)
            and _row_has_number(ws, r, layout.name_col + 1)
        ):
            extra.setdefault(current, []).append(r)
            continue
        if name_base3(name):
            current = r
    return extra


def _apply_amounts_to_record(
    record: WehagoRecord, amounts: dict[int, float], col_to_key: dict[int, str]
) -> None:
    for dst_col, value in amounts.items():
        key = col_to_key.get(dst_col)
        if key is None:
            continue
        if key in record.allowances:
            record.allowances[key] = record.allowances.get(key, 0.0) + value
        elif key in record.deductions:
            record.deductions[key] = record.deductions.get(key, 0.0) + value


def _finalize(record: WehagoRecord, non_taxable_keys: set[str]) -> None:
    record.지급액계, record.공제액계, record.차인지급액 = compute_totals(record.allowances, record.deductions)
    record.비과세합계, record.과세합계 = compute_taxable_split(record.allowances, non_taxable_keys)


def _apply_source_totals_check(
    record: WehagoRecord, src_pay: float | None, src_ded: float | None, src_net: float | None
) -> None:
    for total_field, stated in (("지급액계", src_pay), ("공제액계", src_ded), ("차인지급액", src_net)):
        if stated is None:
            continue
        record.source_totals[total_field] = stated
        computed = getattr(record, total_field)
        diff = computed - stated
        if abs(diff) > _VALUE_TOLERANCE:
            sign = "+" if diff > 0 else "-"
            record.mismatches.append(
                f"{total_field}: 사업장 {stated:,.0f}원 → 위하고 {computed:,.0f}원 (차이 {sign}{abs(diff):,.0f}원)"
            )
            record.mismatched_fields.add(total_field)


# --- main entry point --------------------------------------------------


def reconcile(
    client_name: str,
    source_ws,
    source_layout: SourceLayout,
    dest_ws,
    dest_layout: DestLayout,
    manual_overrides: dict[int, str | None] | None = None,
    learned_mappings: dict[str, str] | None = None,
) -> ReconcileResult:
    """`manual_overrides` maps a source column index -> a destination
    column's record key (from `ClientProfile.allowance_columns`/
    `deduction_columns`, i.e. what `ui/verification_dialog.py`'s picker
    shows), or `None` to force that column to stay unmapped even if
    `build_map` would otherwise have resolved it automatically. This is
    the accountant's manual fix for a header text canon.py doesn't (yet, or
    can't safely) recognize on its own — see CLAUDE.md: rather than a code
    change per new synonym forever, 항목매핑결과's "미매칭" rows are directly
    pickable and the whole run re-executes with the override applied.
    `manual_overrides` itself is never persisted — scoped to one run.

    `learned_mappings` (normalized source header text -> destination
    column's record key, from `core/registry.py`'s per-company
    `column_mappings` table) is a persisted, lower-priority version of the
    same idea — 사용자 요청(2026-09-08): 한 번 손으로 맞춰준 매칭을 그
    사업장에 한해 다음 달에도 기억해서 자동 적용해달라는 것. Applied to any
    column `build_map` left unmapped, BEFORE `manual_overrides` — an
    explicit pick in the current run always wins over a remembered one, and
    can still force a learned column back to "매칭 안 함" via
    `manual_overrides[col] = None`."""
    dst_cols = build_dst_cols(dest_ws, dest_layout)
    allowance_columns, deduction_columns, col_to_key = _dst_column_schema(dest_ws, dest_layout, dst_cols)
    non_taxable_keys = {c.key for c in allowance_columns if not c.taxable}

    column_map, unmapped, ambiguous_src_cols = build_map(
        source_ws, source_layout, dst_cols, build_dst_text_index(dest_ws, dest_layout)
    )
    key_to_col = {key: col for col, key in col_to_key.items()}

    learned_mapped_cols: set[int] = set()
    # 사업장별로 "이 항목은 항상 매칭하지 마라"고 저장해둔(빈 문자열
    # dest_key — registry.save_column_mapping/get_column_mappings 참고)
    # 열들. build_map이 뭘로 자동매칭했든 강제로 미매칭 처리한다(사용자
    # 요청, 2026-09-14 — "선택 안 함"으로 바꾼 것도 다음 실행/빌드에
    # 그대로 반영돼야 한다는 것. 예전엔 "일회성 판단"이라며 이 경우만
    # 저장하지 않았는데, 그러면 매번 다시 손으로 꺼줘야 했다).
    learned_forced_unmapped_cols: set[int] = set()
    if learned_mappings:
        first_statutory: dict[str, int] = {}
        for c in range(1, source_ws.max_column + 1):
            if group_of_src(source_ws, source_layout, c) == "공제":
                ca = canon(source_item(source_ws, source_layout, c))
                if ca in {"국민연금", "건강보험", "장기요양보험료", "고용보험", "소득세", "지방소득세"}:
                    first_statutory.setdefault(ca, c)
        # 미매칭 열뿐 아니라 **이미 build_map이 (잘못) 자동매칭해버린 열도**
        # 대상으로 본다 — 실사 확인(2026-09-08, 업체001어르신주야간보호센터):
        # 공제 그룹의 "성과급" 항목이 canon.py에서 수당 쪽 "성과급"과 같은
        # canon으로 떨어져서, 사용자가 "성과수당"으로 직접 고쳐 저장까지
        # 해놔도 이 열이 이미 (틀리게) 자동매칭된 상태라 미매칭 목록에 없어
        # 저장된 매칭이 평생 반영될 기회가 없었다. 그래서 수당/공제 그룹인
        # 열은 전부 훑어서, 저장된 매칭이 있으면 build_map의 판단(맞든
        # 틀리든)을 덮어쓴다 — 우선순위는 여전히 manual_overrides(이번 실행
        # 중 선택) > learned_mappings(저장된 과거 선택) > build_map 순.
        for src_col in range(1, source_ws.max_column + 1):
            src_group = group_of_src(source_ws, source_layout, src_col)
            # 그룹을 못 정한 열("")도 이 사업장에 그 이름으로 저장된 매칭이
            # 있을 때만 적용한다 — 정답 파일에서 학습한 매칭(2026-10-08,
            # core/mapping_learn.py)에 "추가근로수당"처럼 상위 머리글이 없어
            # 수당/공제를 못 정한 금액 열이 실제로 있었다. 저장된 게 없으면
            # 아래 `key not in learned_mappings`에서 그대로 건너뛴다.
            if src_group not in ("수당", "공제", ""):
                continue
            # A saved "|국민연금" must not revive an employer-share column:
            # 업체042 (2026-10-12) "기관부담금 > 국민연금" was mapped
            # onto 국민연금 and every employee's pension was deducted twice.
            if is_source_employer_burden_col(source_ws, source_layout, src_col):
                continue
            # 그룹 접두어 없는 최하위 항목명으로 조회한다(source_item) —
            # source_full_header(그룹까지 포함)를 쓰면 그룹 헤더 행이 있는
            # 달과 없는 달에 같은 항목이 다른 텍스트로 보여서 저장된
            # 매칭을 다시 못 찾는 버그가 있었다(사용자 실사, 2026-09-08).
            # 키에 그룹(수당/공제)도 같이 넣는다 — "성과급"처럼 같은
            # 텍스트가 수당/공제 양쪽에 다 있으면서 서로 다른 위하고 열을
            # 가리켜야 하는 실제 사례가 있어서(같은 실사) 그룹 없이는
            # 한쪽 수정이 다른 쪽까지 잘못 덮어써 버렸다.
            item_label = source_item(source_ws, source_layout, src_col)
            if not item_label:
                continue
            key = norm(learned_mapping_key(src_group, item_label))
            if key not in learned_mappings:
                continue  # 저장된 적 없는 항목 — build_map 판단 그대로 둔다
            dest_key = learned_mappings[key]
            if not dest_key:
                # "선택 안 함"으로 저장된 항목 — build_map이 자동매칭했어도
                # 강제로 미매칭 처리해서 항목매핑결과에 계속 보이게 한다
                # (manual_overrides의 dest_key=None과 같은 원칙). 아래
                # learned_mapped_cols 필터에서 다시 걸러지지 않도록 그
                # 집합에는 넣지 않는다 — manual_overrides 쪽의 같은 패턴
                # 참고(바로 아래 블록).
                if column_map.pop(src_col, None) is not None:
                    full = source_full_header(source_ws, source_layout, src_col)
                    if full:
                        unmapped[f"{full} [{get_column_letter(src_col)}열]"] = src_col
                learned_forced_unmapped_cols.add(src_col)
                continue
            # 저장 매칭도 보험/세금 반복 열을 다시 살려 두 번 공제하지 않는다.
            ca = canon(item_label)
            if src_group == "공제" and ca in first_statutory and first_statutory[ca] != src_col:
                continue
            dest_col = key_to_col.get(dest_key)
            if dest_col is not None:
                column_map[src_col] = dest_col
                learned_mapped_cols.add(src_col)
        unmapped = {label: col for label, col in unmapped.items() if col not in learned_mapped_cols}

    manually_mapped_cols: set[int] = set()
    if manual_overrides:
        for src_col, dest_key in manual_overrides.items():
            # An explicit choice in this run always wins over — and
            # supersedes the "이 사업장에 저장된 매칭을 자동 적용했습니다"
            # labeling of — a remembered mapping for the same column.
            learned_mapped_cols.discard(src_col)
            if dest_key is None:
                # Forcing "매칭 안 함" on a column build_map WOULD have
                # auto-mapped (it's not already in `unmapped`) — add it there
                # explicitly so it stays visible in 항목매핑결과 instead of
                # silently disappearing from both lists.
                if column_map.pop(src_col, None) is not None:
                    full = source_full_header(source_ws, source_layout, src_col)
                    if full:
                        unmapped[f"{full} [{get_column_letter(src_col)}열]"] = src_col
                continue
            dest_col = key_to_col.get(dest_key)
            if dest_col is not None:
                column_map[src_col] = dest_col
                manually_mapped_cols.add(src_col)
        unmapped = {label: col for label, col in unmapped.items() if col not in manually_mapped_cols}

    src_pay_col = _find_src_pay_total_col(source_ws, source_layout)
    src_ded_col = _find_src_canon_col(source_ws, source_layout, "공제액계")
    src_net_col = _find_src_canon_col(source_ws, source_layout, "차인지급액")

    # --- index source rows by name --------------------------------------
    src_last_row = last_nonblank_row(source_ws)
    real_total_row = find_total_row(source_ws, source_layout.data_start_row, src_last_row)
    src_total_row = real_total_row or (src_last_row + 1)

    continuation = _continuation_rows(source_ws, source_layout, src_total_row)
    continuation_rows = {row for rows in continuation.values() for row in rows}

    source_index = SourceNameIndex()
    for r in range(source_layout.data_start_row, src_total_row, source_layout.block_height):
        if r in continuation_rows:
            continue
        name = source_ws.cell(row=r, column=source_layout.name_col).value
        if name_base3(name):
            source_index.add(r, name)

    def source_birth(row: int) -> str:
        name = source_ws.cell(row=row, column=source_layout.name_col).value
        birth_cell = (
            source_ws.cell(row=row, column=source_layout.birth_col).value if source_layout.birth_col else None
        )
        return person_birth_key(name, birth_cell)

    # --- pass 1: walk destination rows -----------------------------------
    dst_last_row = last_nonblank_row(dest_ws)
    dst_total_row = find_total_row(dest_ws, dest_layout.data_start_row, dst_last_row) or (dst_last_row + 1)

    records: list[WehagoRecord] = []
    verification_rows: list[VerificationRow] = []
    used_source_rows: set[int] = set()
    kept_dest_rows: list[tuple[int, str]] = []  # (row, name) for rows that survive pass 1
    learn_pairs: list[tuple[int, int, tuple[int, ...]]] = []  # (dest_row, src_row, continuation) — _LEARN_SINK용
    summary = ReconcileSummary()

    for r in range(dest_layout.data_start_row, dst_total_row):
        name = dest_ws.cell(row=r, column=dest_layout.name_col).value
        if norm(name) == "합계" or not name_base3(name):
            continue

        birth_cell = dest_ws.cell(row=r, column=dest_layout.birth_col).value if dest_layout.birth_col else None
        dest_birth = person_birth_key(name, birth_cell)

        src_row, memo = resolve_source_row(name, dest_birth, source_index, used_source_rows, source_birth)

        if src_row is None:
            if memo == "":
                summary.removed += 1
                verification_rows.append(
                    VerificationRow(
                        name=str(name),
                        status="삭제됨",
                        memo="위하고에는 있으나 사업장 원본에는 없는 사원이라 정리본에서 제외했습니다.",
                    )
                )
                continue
            summary.need_check += 1
            verification_rows.append(
                VerificationRow(name=str(name), status="확인필요", memo=f"{memo} 자동 수정하지 않았습니다.")
            )
            kept_dest_rows.append((r, str(name)))
            record = WehagoRecord(
                사원코드=str(dest_ws.cell(row=r, column=dest_layout.code_col).value or "")
                if dest_layout.code_col
                else "",
                사원명=str(name),
                allowances={c.key: 0.0 for c in allowance_columns},
                deductions={c.key: 0.0 for c in deduction_columns},
            )
            for c, key in col_to_key.items():
                existing = _as_number(dest_ws.cell(row=r, column=c).value) or 0.0
                if key in record.allowances:
                    record.allowances[key] = existing
                elif key in record.deductions:
                    record.deductions[key] = existing
            _finalize(record, non_taxable_keys)
            records.append(record)
            continue

        used_source_rows.add(src_row)
        kept_dest_rows.append((r, str(name)))
        learn_pairs.append((r, src_row, tuple(continuation.get(src_row, ()))))

        amounts = _row_amounts(
            source_ws, src_row, source_layout.block_height, column_map, continuation.get(src_row, ())
        )
        record = WehagoRecord(
            사원코드=str(dest_ws.cell(row=r, column=dest_layout.code_col).value or "") if dest_layout.code_col else "",
            사원명=str(name),
            allowances={c.key: 0.0 for c in allowance_columns},
            deductions={c.key: 0.0 for c in deduction_columns},
        )
        _apply_amounts_to_record(record, amounts, col_to_key)
        _finalize(record, non_taxable_keys)

        src_pay = _block_number(
            source_ws, src_row, source_layout.block_height, src_pay_col, continuation.get(src_row, ())
        )
        src_ded = _block_number(
            source_ws, src_row, source_layout.block_height, src_ded_col, continuation.get(src_row, ())
        )
        src_net = _block_number(
            source_ws, src_row, source_layout.block_height, src_net_col, continuation.get(src_row, ())
        )
        _apply_source_totals_check(record, src_pay, src_ded, src_net)

        status = "차이" if record.mismatched_fields else "정상"
        if status == "차이":
            summary.diff += 1
        summary.matched += 1

        verification_rows.append(
            VerificationRow(
                name=str(name),
                status=status,
                src_pay=src_pay,
                dst_pay=record.지급액계,
                diff_pay=(src_pay - record.지급액계) if src_pay is not None else None,
                src_ded=src_ded,
                dst_ded=record.공제액계,
                diff_ded=(src_ded - record.공제액계) if src_ded is not None else None,
                src_net=src_net,
                dst_net=record.차인지급액,
                diff_net=(src_net - record.차인지급액) if src_net is not None else None,
                memo=memo,
            )
        )
        records.append(record)

    # --- pass 2: leftover source rows -> add or flag ---------------------
    dest_base_index: dict[str, list[int]] = {}
    for row, name in kept_dest_rows:
        base_key = name_base3(name)
        if base_key:
            dest_base_index.setdefault(base_key, []).append(row)

    def dest_birth(row: int) -> str:
        birth_cell = dest_ws.cell(row=row, column=dest_layout.birth_col).value if dest_layout.birth_col else None
        name = dest_ws.cell(row=row, column=dest_layout.name_col).value
        return person_birth_key(name, birth_cell)

    for r in range(source_layout.data_start_row, src_total_row, source_layout.block_height):
        if r in used_source_rows or r in continuation_rows:
            continue
        name = source_ws.cell(row=r, column=source_layout.name_col).value
        if not name_base3(name):
            continue

        src_birth = source_birth(r)
        if should_add_source_row(name, src_birth, dest_base_index, dest_birth):
            amounts = _row_amounts(source_ws, r, source_layout.block_height, column_map, continuation.get(r, ()))
            record = WehagoRecord(
                사원코드="사원코드 없음",
                사원명=str(name),
                allowances={c.key: 0.0 for c in allowance_columns},
                deductions={c.key: 0.0 for c in deduction_columns},
            )
            _apply_amounts_to_record(record, amounts, col_to_key)
            _finalize(record, non_taxable_keys)

            src_pay = _block_number(source_ws, r, source_layout.block_height, src_pay_col, continuation.get(r, ()))
            src_ded = _block_number(source_ws, r, source_layout.block_height, src_ded_col, continuation.get(r, ()))
            src_net = _block_number(source_ws, r, source_layout.block_height, src_net_col, continuation.get(r, ()))
            _apply_source_totals_check(record, src_pay, src_ded, src_net)

            status = "추가됨_차이" if record.mismatched_fields else "추가됨"
            if status == "추가됨_차이":
                summary.diff += 1
            summary.added += 1

            verification_rows.append(
                VerificationRow(
                    name=str(name),
                    status=status,
                    src_pay=src_pay,
                    dst_pay=record.지급액계,
                    diff_pay=(src_pay - record.지급액계) if src_pay is not None else None,
                    src_ded=src_ded,
                    dst_ded=record.공제액계,
                    diff_ded=(src_ded - record.공제액계) if src_ded is not None else None,
                    src_net=src_net,
                    dst_net=record.차인지급액,
                    diff_net=(src_net - record.차인지급액) if src_net is not None else None,
                    memo="위하고에 해당 직원이 없어 정리본에 신규 행을 추가했습니다. 사원코드는 '사원코드 없음'으로 입력했습니다.",
                )
            )
            records.append(record)
        else:
            summary.need_check += 1
            verification_rows.append(
                VerificationRow(
                    name=str(name),
                    status="확인필요",
                    memo="위하고에 앞 3글자가 같은 이름이 있으나 자동 매칭되지 않았습니다. 생년월일 또는 사원정보 확인이 필요합니다.",
                )
            )

    mapping_rows: list[MappingRow] = []
    ambiguous_dest_keys: set[str] = set()
    for src_col, dst_col in column_map.items():
        is_learned = src_col in learned_mapped_cols
        is_manual = src_col in manually_mapped_cols or is_learned
        # A manual (or remembered) pick means the accountant already looked
        # at it and resolved the ambiguity themselves — only flag matches
        # build_map made on its own.
        is_ambiguous = not is_manual and src_col in ambiguous_src_cols
        dest_label = col_to_key.get(dst_col, get_column_letter(dst_col))
        if is_learned:
            note = "이 사업장에 저장된 매칭을 자동 적용했습니다."
        elif is_manual:
            note = "사용자가 직접 지정했습니다."
        else:
            note = "해당 위하고 열에 금액 반영"
        if is_ambiguous:
            note += " — 같은 이름의 다른 항목일 수 있어 확인이 필요합니다."
            ambiguous_dest_keys.add(dest_label)
        mapping_rows.append(
            MappingRow(
                kind="수동매칭" if is_manual else "자동매칭",
                source_header=source_full_header(source_ws, source_layout, src_col),
                dest_column_label=f"{dest_label} ({get_column_letter(dst_col)}열)",
                note=note,
                source_col=src_col,
                dest_key=dest_label,
                ambiguous=is_ambiguous,
                item_label=source_item(source_ws, source_layout, src_col),
                group=group_of_src(source_ws, source_layout, src_col),
            )
        )
    for label, src_col in unmapped.items():
        # `source_header`는 화면 표시용(그룹까지 포함, 열 문자는 note로
        # 옮김) — 저장 매칭 키는 `item_label`(그룹 없는 최하위 항목명) 쪽을
        # 쓴다(사용자 요청, 2026-09-08 — source_header를 그대로 키로 쓰면
        # 그룹 헤더 행이 있는 달과 없는 달에 텍스트가 달라져 저장된 매칭을
        # 다시 못 찾는 버그가 있었다).
        full = source_full_header(source_ws, source_layout, src_col) or label
        if src_col in learned_forced_unmapped_cols:
            note = f'이 사업장에서 "매칭 안 함"으로 저장되어 있습니다({get_column_letter(src_col)}열).'
        else:
            note = f"위하고 항목과 자동 매칭하지 못했습니다({get_column_letter(src_col)}열). 금액이 있으면 수동 확인 필요"
        mapping_rows.append(
            MappingRow(
                kind="미매칭",
                source_header=full,
                note=note,
                source_col=src_col,
                dest_key=None,
                item_label=source_item(source_ws, source_layout, src_col) or full,
                group=group_of_src(source_ws, source_layout, src_col),
            )
        )

    profile = ClientProfile(
        id=f"reconcile-{client_name}",
        client_name=client_name,
        version=1,
        created_at=now_iso(),
        updated_at=now_iso(),
        allowance_columns=allowance_columns,
        deduction_columns=deduction_columns,
    )

    grand_total_source: dict[str, float] = {}
    if real_total_row is not None:
        for total_field, col in (
            ("지급액계", src_pay_col),
            ("공제액계", src_ded_col),
            ("차인지급액", src_net_col),
        ):
            if col is None:
                continue
            stated = _as_number(source_ws.cell(row=real_total_row, column=col).value)
            if stated is not None:
                grand_total_source[total_field] = stated

    grand_total_computed = {
        "지급액계": sum(r.지급액계 for r in records),
        "공제액계": sum(r.공제액계 for r in records),
        "차인지급액": sum(r.차인지급액 for r in records),
    }
    grand_total_mismatches = {
        field_name
        for field_name, stated in grand_total_source.items()
        if abs(grand_total_computed[field_name] - stated) > _VALUE_TOLERANCE
    }

    if _LEARN_SINK is not None:
        _LEARN_SINK.append(
            dict(
                source_ws=source_ws,
                source_layout=source_layout,
                dest_ws=dest_ws,
                col_to_key=dict(col_to_key),
                column_map=dict(column_map),
                pairs=learn_pairs,
            )
        )

    return ReconcileResult(
        profile=profile,
        records=records,
        verification_rows=verification_rows,
        mapping_rows=mapping_rows,
        summary=summary,
        ambiguous_dest_keys=ambiguous_dest_keys,
        grand_total_source=grand_total_source,
        grand_total_computed=grand_total_computed,
        grand_total_mismatches=grand_total_mismatches,
    )
