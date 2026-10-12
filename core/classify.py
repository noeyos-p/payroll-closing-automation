"""Column classification — which group (수당/공제/차인지급액) a source or
destination column belongs to, and its canonical item name. Ported from
파일/이지케어_Sheet1_현재시트_수정본_v3.bas's `GroupOfSrc`/`GroupOfDst`/
`CanonOfSrc`/`CanonOfDst`/`IsSourceInfoCol`/`IsSourceEmployerBurdenCol`/
`IsSourceWorkTimeCol`.

The VBA read exactly two fixed header rows (SRC_H1/SRC_H2, DST_H1/DST_H2)
per column. core/sheet_layout.py's detectors instead find however many
header rows a real file actually has, so "h1" here means the detected
anchor/group row and "h2" means the detected leaf/item row — same roles,
computed positions instead of hardcoded ones."""

from __future__ import annotations

import weakref

from core.canon import canon, norm
from core.gridutil import cell_text
from core.sheet_layout import DestLayout, SourceLayout, dst_group_label, dst_item, source_full_header, source_item

_EMPLOYER_BURDEN_KEYWORDS = (
    "인건비신고", "사업자부담", "기관부담", "사회보험료", "산재보험",
    "퇴직적립", "공단신고", "기준소득", "보수월액", "지원율",
    # 4대보험을 "국민연금(본인부담)"/"국민연금(회사분)"처럼 부담 주체별로 두
    # 열로 쪼개는 원본이 실사로 확인됨(업체014) — 직원
    # 본인 부담분은 canon.py의 strip_share_suffix로 걸러내 정상 공제 항목으로
    # 인식하고, 회사(사업장) 부담분은 여기 걸려서 제외되는 게 맞다(직원 급여
    # 명세에는 안 들어가는 사업장 비용).
    "회사분", "사업자분", "사업장부담",
    # 2026-10-07 (업체036): "공단보수신고액 > 연금",
    # "실과세액 > 건강/고용"은 보험료 산정 기준금액(수백만 원)이지 공제액이
    # 아니다 — 국민연금/건강보험 열로 잡혀 공제가 수백만 원씩 부풀었다.
    "보수신고", "실과세액", "과세표준",
)

_INFO_ITEM_KEYS = {
    "no", "번호", "순번", "성명", "사원명", "직원명", "이름",
    "주민번호", "주민등록번호", "주민등록", "직책", "직종",
    "입사일", "퇴사일", "통상임금", "생년월일", "주소",
    # "시급"은 그 달 실제 지급액이 아니라 시간당 단가(참고용 요율)다 —
    # canon.py엔 여전히 "기본급"의 표기 변형(시급제로 급여 자체를 "시급"이라
    # 부르는 위하고 등록 컬럼과의 매칭용)으로 남아있지만, 원본(이지케어)
    # 쪽에서 "시급" 열은 항상 단가이지 금액이 아니므로 여기서 걸러 금액으로
    # 합산되지 않게 한다(사용자 확인: 시급 값이 기본급에 잘못 더해지고
    # 있었음, 2026-08-06).
    "시급",
    # 2026-10-07: 펼친(flatten) 양식에서 금액 열 옆에 따로 서는 인적사항 열들
    # (업체034/업체035의 담당업무/부양가족/호봉/보수월액/두루누리 여부, 업체037의
    # 구분/업체038) — 금액이 아니므로 미매칭 목록에도 안 띄운다.
    "담당업무", "부양가족", "호봉", "보수월액", "두루누리여부", "업체038", "퇴사사유", "나이", "서명",
}
_INFO_FULL_KEYWORDS = ("주민번호", "주민등록번호", "입사일", "퇴사일", "통상임금", "기본사항")

_ALLOWANCE_CANON = {
    "기본급", "주휴수당", "시간외수당", "연장근로수당", "야간수당",
    "휴일근로수당", "공휴일근로수당", "근로자의날", "연차수당", "직책수당",
    "보전수당", "목욕수당", "요양수당", "중증수당", "식대", "차량유지비",
    "사무보조비", "주유비", "식비", "장기근속수당", "개별추가수당", "기타수당",
    "기타수당(과)", "선임요양보호사", "상여", "명절수당", "지급액계",
    # 그룹행(상위 헤더) 텍스트가 아예 없는 원본에서, 항목명 자체도
    # canon.py 사전에 등록 안 돼 있으면 group_of_src가 판단할 근거가
    # 전혀 없어 "" 반환 → build_map에서 미매칭 표시도 없이 통째로
    # 누락되는 문제가 있었다(사용자 확인, 2026-08-12). "식대"/"식비"
    # 등과 자동으로 합치기엔 다른 항목일 수 있어(두루누리 사례와 같은
    # 원칙), canon엔 자기 자신 그대로 남기고 그룹만 여기서 인식시켜
    # 항목매핑결과에 뜨게 한다 — 실제 연결은 사용자가 직접 매칭 UI로.
    "식대비", "교육비",
}
_DEDUCTION_CANON = {
    "국민연금", "건강보험", "고용보험", "장기요양보험료",
    "소득세", "지방소득세", "학자금상환액", "기타공제",
    "선지급", "식대공제", "연말정산", "두루누리지원액", "공제액계",
    # 위 "식대비"/"교육비"와 같은 이유(2026-08-12) — "고용보험"과 자동
    # 병합하지 않고 그룹만 인식시켜 미매칭으로 뜨게 한다.
    "고용보험정산",
}


def is_allowance_canon(ca: str) -> bool:
    return ca in _ALLOWANCE_CANON


def is_deduction_canon(ca: str) -> bool:
    return ca in _DEDUCTION_CANON


def is_source_employer_burden_col(ws, layout: SourceLayout, col: int) -> bool:
    full = norm(source_full_header(ws, layout, col))
    if "장기근속" in full:
        # "장기근속퇴직적립금"처럼 장기근속 관련 항목은 일반적인 "퇴직적립"
        # (사업장이 부담하는, 직원 급여명세에 안 들어가는 비용)과 달리 실제
        # 직원에게 지급되는 항목이라 예외로 둔다 — 이것마저 걸러지면 목록에
        # 아예 안 보여서 손으로 매칭할 기회조차 없었다(사용자 요청,
        # 2026-09-08).
        return False
    return any(kw in full for kw in _EMPLOYER_BURDEN_KEYWORDS)


def is_source_info_col(ws, layout: SourceLayout, col: int) -> bool:
    raw_item = source_item(ws, layout, col)
    item_key = norm(raw_item)
    if item_key in _INFO_ITEM_KEYS:
        return True
    # "기본급(시)"/"시간당 단가" — 시간당 단가는 금액이 아니다(2026-10-07
    # 업체016: "기본사항 > 기본급(시)"가 기본급에 더해질 뻔함).
    compact = raw_item.replace(" ", "")
    if "(시)" in compact or "단가" in compact or "시간당" in compact:
        return True
    full_key = norm(source_full_header(ws, layout, col))
    return any(kw in full_key for kw in _INFO_FULL_KEYWORDS)


def is_source_work_time_col(ws, layout: SourceLayout, col: int) -> bool:
    full_key = norm(source_full_header(ws, layout, col))
    if "근로시간" in full_key or "근무시간" in full_key:
        return True
    return norm(source_item(ws, layout, col)) in {"총근로", "야간", "휴일", "연장"}


def group_of_src(ws, layout: SourceLayout, col: int) -> str:
    if is_source_employer_burden_col(ws, layout, col):
        return ""
    if is_source_info_col(ws, layout, col):
        return ""
    if is_source_work_time_col(ws, layout, col):
        return ""

    top_key = norm(cell_text(ws, layout.header_rows[0], col))
    if top_key in ("실지급액", "실수령액", "차인지급액", "지급급여"):
        return "차인지급액"
    if top_key in ("공제액계", "사회보험", "소득세", "기타공제"):
        return "공제"
    if top_key in ("임금상세내역", "지급", "급여", "수당"):
        return "수당"

    full_key = norm(source_full_header(ws, layout, col))
    ca = canon(source_item(ws, layout, col))

    if ca == "차인지급액" or any(kw in full_key for kw in ("실수령액", "실지급액", "차인지급액", "지급급여")):
        return "차인지급액"
    if ca == "공제액계" or is_deduction_canon(ca) or any(
        kw in full_key for kw in ("공제", "소득세", "지방세", "주민세", "식대공제")
    ):
        return "공제"
    if ca == "지급액계" or is_allowance_canon(ca) or any(
        kw in full_key for kw in ("급여", "수당", "지급", "임금상세내역")
    ):
        return "수당"
    # "기본급(식대 미포함)"처럼 사전 항목명 뒤에 설명이 붙은 단독 머리글
    # (2026-10-07 업체039류) — 앞부분이 수당/공제 항목명이면 그 그룹.
    if any(len(a) >= 3 and ca.startswith(a) for a in _PREFIX_ALLOWANCES):
        return "수당"
    if any(len(d) >= 4 and ca.startswith(d) for d in _PREFIX_DEDUCTIONS):
        return "공제"
    return ""


_PREFIX_ALLOWANCES = ("기본급", "주휴수당", "연차수당", "보전수당", "장기근속수당", "직책수당")
_PREFIX_DEDUCTIONS = ("국민연금", "건강보험", "고용보험", "장기요양", "지방소득세")


def canon_of_src(ws, layout: SourceLayout, col: int) -> str:
    top_key = norm(cell_text(ws, layout.header_rows[0], col))
    item_key = norm(source_item(ws, layout, col))
    full_key = norm(source_full_header(ws, layout, col))

    if is_source_employer_burden_col(ws, layout, col):
        return ""
    if is_source_info_col(ws, layout, col):
        return ""

    if any(kw in full_key for kw in ("실지급액", "실수령액", "차인지급액")) or top_key == "지급급여":
        return "차인지급액"
    if top_key == "지급" and item_key in ("지급액", "지급금액", "지급총액", "지급합계", "소득합계", "급여총액"):
        return "지급액계"
    # "지급총액" as a substring, not just an exact match — real headers wrap
    # it with more text (confirmed real case: 업체013's
    # "지급총액(과세+비과세)", which after paren-stripping is
    # "지급총액과세+비과세", not the bare word).
    if "지급총액" in item_key or item_key in ("지급합계", "소득합계", "급여총액", "지급액계"):
        return "지급액계"
    if item_key in ("공제계", "공제합계", "공제액계", "공제총액", "공제액") or top_key == "공제액계":
        return "공제액계"
    if top_key == "임금상세내역" and item_key == "계":
        return "지급액계"

    group = group_of_src(ws, layout, col)
    ca = canon(source_item(ws, layout, col))
    if group == "공제" and ca == "식대":
        return "식대공제"
    if group == "공제" and item_key in ("합계", "계", "소계"):
        return "공제액계"
    return ca


def group_of_dst(ws, layout: DestLayout, col: int) -> str:
    # 위치(지급액계 앞=수당, 뒤=공제)가 잡히면 그게 글자 추측보다 정확하다 —
    # 공제 영역에 "성과수당"처럼 '수당'이 들어간 이름의 공제 열이 실제로 있다
    # (업체001어르신주야간보호센터).
    positional = _dst_positional_groups(ws, layout).get(col)
    if positional:
        return positional
    return _group_of_dst_by_text(ws, layout, col)


# 위하고 Sheet1에서 금액 항목이 아닌 정보 열 — 위치 기반 분류에서 제외한다.
_DST_INFO_KEYS = {
    "사원코드", "사원번호", "사번", "사원명", "성명", "이름", "부서", "직급", "직종", "직책", "직위",
    "입사일", "퇴사일", "생년월일", "주민번호", "주민등록번호", "비고",
}
_DST_SUBTOTAL_KEYS = {"과세합계", "비과세합계", "과세계", "비과세계", "과세", "비과세"}

_DST_POSITIONAL_CACHE: "weakref.WeakKeyDictionary[object, dict[int, str]]" = weakref.WeakKeyDictionary()


def _dst_positional_groups(ws, layout: DestLayout) -> dict[int, str]:
    """위하고 업로드 양식(Sheet1)은 항목 순서가 정해져 있다 — 사원명 뒤부터
    지급액계 앞까지가 수당, 지급액계 뒤부터 공제액계 앞까지가 공제다. 이
    앱이 내보낸 파일(그리고 그걸 참고파일로 다시 등록한 것)은 헤더가 한
    줄뿐이라 "처우개선비"/"두루누리 국민연금"/"가족휴가제"처럼 사전에 없는
    항목명은 글자만으로는 수당/공제를 알 수 없어서, 원본에 글자 그대로
    같은 항목이 있어도 매칭 대상에서 아예 빠지고 있었다(2026-10-07 9월지급
    전체 급여대장을 정답 파일과 대조해서 확인 — 가장 흔한 누락 원인).
    글자로 판단이 안 되는 열만 이 위치 규칙으로 채운다."""
    try:
        cached = _DST_POSITIONAL_CACHE.get(ws)
    except TypeError:
        cached = None
    if cached is not None and cached.get(-1) == id(layout):
        return cached
    pay_col = ded_col = None
    for c in range(1, ws.max_column + 1):
        ca = canon(dst_item(ws, layout, c))
        if ca == "지급액계" and pay_col is None:
            pay_col = c
        elif ca == "공제액계" and ded_col is None:
            ded_col = c
    groups: dict[int, str] = {-1: id(layout)}  # type: ignore[dict-item]
    if pay_col is not None:
        start = max(layout.name_col, layout.code_col or 0) + 1
        end = ded_col if ded_col is not None and ded_col > pay_col else None
        for c in range(start, ws.max_column + 1):
            if c == pay_col or c == ded_col:
                continue
            if end is not None and c >= end:
                break
            key = norm(dst_item(ws, layout, c))
            if not key or key in _DST_INFO_KEYS or key in _DST_SUBTOTAL_KEYS:
                continue
            if "차인지급" in key or "실수령" in key:
                continue
            groups[c] = "수당" if c < pay_col else "공제"
    try:
        _DST_POSITIONAL_CACHE[ws] = groups
    except TypeError:
        pass
    return groups


def _group_of_dst_by_text(ws, layout: DestLayout, col: int) -> str:
    h1 = dst_group_label(ws, layout, col)
    h2 = dst_item(ws, layout, col)
    k1 = norm(h1)
    kall = norm(h1 + h2)
    ca = canon(h2) or canon(h1)

    if ca == "차인지급액" or "차인지급액" in kall or "실수령액" in kall:
        return "차인지급액"
    if ca == "공제액계" or is_deduction_canon(ca) or k1 == "공제" or "공제" in kall:
        return "공제"
    if (
        ca == "지급액계"
        or is_allowance_canon(ca)
        or k1 in ("급여", "수당", "지급", "과세", "비과세", "과세수당", "비과세수당")
        or "수당" in kall
        or "급여" in kall
        or "지급" in kall
    ):
        return "수당"
    return ""


def canon_of_dst(ws, layout: DestLayout, col: int) -> str:
    group = group_of_dst(ws, layout, col)
    item = dst_item(ws, layout, col)
    ca = canon(item)
    if group == "공제" and ca == "식대":
        return "식대공제"
    return ca
