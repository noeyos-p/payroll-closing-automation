"""Korean payroll term normalization + canonicalization, ported near-verbatim
from 파일/이지케어_Sheet1_현재시트_수정본_v3.bas (`Norm`/`HeaderKey`/`Canon`/
`StripTaxSuffix`/`StripFixedSuffix`/`IsVehicleAllowanceKey`/`IsNumberedBase`/
`IsKnownBiEndingItem`). This dictionary is the core of the header-text-based
matching engine that replaced this app's old value-cross-reference matching
(core/heuristic_match.py, removed) — see CLAUDE.md for why.

`norm()` plays both of the VBA's `Norm`/`HeaderKey` roles (they were
byte-identical functions in the macro, just duplicated under two names for
two call sites)."""

from __future__ import annotations

import re

_STRIP_CHARS = (
    "\r", "\n", "\t", " ", "　", "(", ")", "（", "）", "-", "_", "/", ".", ",", "·", "ㆍ", ":", "：",
)

_BI_ENDING_ITEMS = {
    "식비", "주유비", "교통비", "유류비", "차량유지비",
    "사무보조비", "업무보조비", "보수교육비", "직무교육비",
    "식대보조비", "차량운행비", "차량보조비",
}


def norm(value: object) -> str:
    s = str(value if value is not None else "").strip().lower()
    for ch in _STRIP_CHARS:
        s = s.replace(ch, "")
    return s


# Alias kept for readability at call sites that mirror the VBA's HeaderKey
# usage (header-cell text) versus Norm usage (general text) — same function.
header_key = norm


def is_known_bi_ending_item(k: str) -> bool:
    return k in _BI_ENDING_ITEMS


def strip_tax_suffix(k: str) -> str:
    if len(k) >= len("비과세") and k.endswith("비과세"):
        base = k[: -len("비과세")]
        candidate = base + "비"
        return candidate if is_known_bi_ending_item(candidate) else base
    if len(k) >= len("과세") and k.endswith("과세"):
        return k[: -len("과세")]
    return k


_RATE_SUFFIX_RE = re.compile(r"\d+%$")


def strip_rate_suffix(k: str) -> str:
    """Strips a trailing 4대보험 rate annotation like "(4.5%)" — by the time
    this runs, `norm()` has already dropped the parens and the decimal
    point, so "국민연금(4.5%)" has become "국민연금45%"; without this the
    rate digits stay glued onto the label and it never matches the plain
    "국민연금" alias (confirmed real case: 업체013's 원본 labels
    every 4대보험 deduction column this way — 국민연금(4.5%), 건강보험
    (3.43%), 장기요양(11.52%), 고용보험(0.8%) — and every one of them was
    silently unmapped as a result, understating 차인지급액 for every
    employee by the sum of their own 4대보험)."""
    return _RATE_SUFFIX_RE.sub("", k)


def strip_fixed_suffix(k: str) -> str:
    for tok in ("고정값", "고정", "확정값", "확정"):
        k = k.replace(tok, "")
    return k


_SHARE_SUFFIXES = ("본인부담", "근로자부담", "개인부담", "자기부담")


def strip_share_suffix(k: str) -> str:
    """Strips a trailing "(본인부담)"-style annotation — some 원본 files
    split each 4대보험 item into the employee's own share and the employer's
    (업체014: "국민연금(본인부담)"/"국민연금(회사분)"),
    which without this leaves "국민연금본인부담" not matching the plain
    "국민연금" alias. The employer-share column is a separate, correct
    exclusion (core/classify.py's `_EMPLOYER_BURDEN_KEYWORDS` — it's the
    business's own cost, not something deducted from the employee's pay)."""
    for suffix in _SHARE_SUFFIXES:
        if k.endswith(suffix) and len(k) > len(suffix):
            # "본인부담" 자체가 항목명인 경우(업체015 "공제 > 본인부담")는
            # 지우면 빈 이름이 되므로 그대로 둔다.
            return k[: -len(suffix)]
    return k


_FORMULA_PAREN_RE = re.compile(r"\(\s*[a-zA-Z](\s*=\s*[a-zA-Z0-9+\-\s]*)?\s*\)\s*$")


def strip_formula_paren_suffix(text: str) -> str:
    """Strips a trailing single-letter or spreadsheet-formula annotation
    like "(a)"/"(c=a-b)" — confirmed real case: 업체016's 원본
    labels its 3 grand-total columns "급여총액(a)"/"공제총액(b)"/"지급액
    (c=a-b)", spelling out the arithmetic for a human reader. `norm()` only
    strips the parens themselves, leaving "급여총액a"/"지급액c=ab" behind,
    which no longer matches the plain "급여총액"/"지급액" aliases already
    registered below — this has to run on the RAW text (before norm() drops
    the parens) to tell "a real annotation" apart from a label that
    legitimately ends in a letter."""
    return _FORMULA_PAREN_RE.sub("", text).strip()


def is_vehicle_allowance_key(k: str) -> bool:
    x = strip_tax_suffix(k)
    # VBA only did an exact-string check for "교통비" itself (missing real
    # variants like "교통비지원"/"교통비(방문)" — confirmed real gap against
    # 업체017's Sheet1); substring, like the other three
    # vehicle-allowance keywords already get, catches those too.
    # "가산"(가산액/가산금 등)이 붙은 복합어는 제외한다 — "원거리교통비가산액"은
    # 실제 차량/유류비 보전이 아니라 "기타수당" 그룹의 별개 가산 항목인데(실사:
    # 사업장 급여대장에 "차량유지비(비과세)" 열이 따로 있는데도 이 substring
    # 규칙 때문에 둘 다 같은 위하고 "차량유지비" 열로 몰려서 충돌했다,
    # 2026-08-07), "교통비" 부분일치만으로 차량유지비에 합쳐버리면 서로 다른
    # 두 항목이 같은 열을 두고 경합하게 된다.
    if "가산" in x:
        return x in ("차량운행비", "주유비", "유류비", "유류대")
    if "자가운전" in x or "차량유지" in x or "차량보조" in x or "교통비" in x:
        return True
    return x in ("차량운행비", "주유비", "유류비", "유류대")


def is_numbered_base(s: str, base_text: str) -> bool:
    """True if `s` is exactly `base_text`, or `base_text` followed only by
    digits (e.g. "기타수당2" for base "기타수당")."""
    if s == base_text:
        return True
    if not s.startswith(base_text):
        return False
    tail = s[len(base_text):]
    if tail == "":
        return True
    return tail.isdigit()


_EXACT_CANON: dict[str, str] = {}


def _register(canon_name: str, *aliases: str) -> None:
    for alias in aliases:
        _EXACT_CANON[alias] = canon_name


_register("기본급", "기본급", "기본급여", "월급", "시급", "근로금액", "급여", "기본수당")
_register("주휴수당", "주휴수당", "주휴")
_register("시간외수당", "시간외수당", "시간외", "시간외근로수당", "시간외수당금액", "시간외근로수당금액")
_register(
    "연장근로수당",
    "연장수당", "연장근로수당", "연장근무수당", "연장근무", "연장", "연장수당금액", "연장근로수당금액",
    "초과근무수당", "초과수당", "초과연장근로수당", "초과연장수당", "연장추가근무수당",
)
_register("야간수당", "야간수당", "야간근로수당", "야간", "야간수당금액", "야간근로수당금액", "야간수당05배")
_register(
    "휴일근로수당",
    "휴일수당", "휴일근로수당", "휴일", "휴일수당금액", "휴일근로수당금액", "유급휴일수당", "유급휴일",
)
_register(
    "공휴일근로수당",
    "공휴일수당", "공휴일근로수당", "공휴일", "공휴수당", "공휴일수당금액", "공휴일근로수당금액",
)
_register("연차수당", "연차수당", "연차", "연차유급수당", "연차사용수당", "연차수당금액", "년월차수당", "연월차수당", "년차수당")
_register("근로자의날", "근로자의날", "근로자의날수당", "노동절", "노동절수당", "근로자날", "근로자의날수당금액")
_register("직책수당", "직책수당", "직무수당", "직급수당", "기타직책수당", "직책", "직책급")
_register("보전수당", "보전수당", "보전", "보전급여", "조정보전수당")
_register("목욕수당", "목욕수당", "방문목욕수당", "목욕근로수당", "목욕근로", "목욕")
_register("요양수당", "요양수당", "방문요양수당", "요양근로수당", "요양근로", "요양")
_register("중증수당", "중증수당", "중증수급자가산액", "중증수급자가산수당", "중증가산", "중증가산액", "중증수당금액",
          "중증가산수당", "중증가산금", "중증수급자가산")
_register("식대", "식대", "식비", "식대보조", "식대보조비", "식사대", "식사비", "식비수당",
          # 2026-10-07 정답 파일 대조: "식대비"(수당/공제 모두)와 "(비)식대"가 실제로
          # 위하고 식대/식대공제 열로 들어가 있었다(업체018/업체019/업체020/업체021 등 8곳).
          # 사업장마다 뜻이 다를 수 있어 "확인 필요"(노란색) 별칭으로도 등록해 둔다.
          "식대비", "비식대")
_register(
    "차량유지비",
    "차량유지비", "차량유지", "자가운전보조금", "자가운전보조", "자가운전", "교통비", "차량보조금",
    "차량운행비", "주유비", "유류비", "유류대", "주유수당", "운전보조", "운전보조금",
    "챠량유지비", "차량지원", "차량비", "자가차보조",
)
_register("명절수당", "명절수당", "명절보너스", "명절상여", "명절상여금")
_register("사무보조비", "사무보조비", "사무보조", "업무보조비", "업무보조", "사무수당")
_register(
    "장기근속수당",
    "장기근속수당", "장기근속", "장기근속장려금", "장기근속장여금", "장기근속장려수당", "장기근속장여수당", "공단장기근속수당",
    "장기근속수당과세", "장기근속장려금과세", "장기근속장여금과세", "근속장려금", "근속장려수당", "근속수당",
)
_register("상여", "상여", "상여금", "보너스")
_register("기타수당", "기타수당과", "기타수당과세", "과세기타수당", "기타수당비과", "비과세기타수당")
_register(
    "개별추가수당",
    "개별추가수당", "개별추가수당과세", "개별추가수당비과세", "추가수당", "추가수당과세", "추가수당비과세",
)
# 위 "개별추가수당"에 예전엔 아래 항목들이 전부 동의어로 뭉쳐 있었다 — 텍스트
# 표기 차이가 아니라 서로 다른 실제 급여 항목(보수교육비/농어촌지원금/육아수당/
# 간호가산액/인지활동형가산액)인데 같은 canon으로 묶여서, 위하고 쪽에 그중
# 아무 열이나 하나 있으면(예: "보수교육비") 나머지가 전부 자동으로 거기 매칭돼
# 버리는 버그가 있었다(2026-08-07, 실사: 농어촌지원금이 보수교육비 열로 조용히
# 합쳐짐). 두루누리/근로자의날에서 이미 고친 것과 같은 패턴이라 여기도 각자
# 독립된 canon으로 분리한다 — 위하고에 정확히 같은 이름의 열이 없으면 자동
# 매칭되지 않고 "미매칭"으로 남아 직접 매칭 UI로 고르게 된다.
_register("보수교육비", "보수교육비", "요양보호사보수교육비", "직무교육비", "직무교육")
_register("농어촌지원금", "농어촌수당", "농어촌수당과세", "농어촌지원금", "농어촌지원금과세")
_register("육아수당", "육아수당", "육아휴직")
_register("간호가산액", "간호가산액")
_register("인지활동형가산액", "인지활동형가산액")
# "원거리교통비가산액"은 차량유지비와 별개의 항목이다 — is_vehicle_allowance_key()가
# "가산"이 붙은 복합어는 더 이상 substring으로 차량유지비에 묶지 않으므로(위 함수
# 참고, 2026-08-07 수정) 여기 등록한 canon이 실제로 쓰인다.
_register("원거리교통비가산액", "원거리교통비가산액")
_register("기타수당", "기타", "기타급여", "기타제수당", "제수당")
_register("지급액계", "소득합계", "지급합계", "지급액계", "급여합계", "지급총액", "급여총액")
_register("국민연금", "국민연금", "연금보험", "연금", "국민", "국민연금수기입력")
_register("건강보험", "건강보험", "건강보험료", "건강", "의료보험", "의료보험료", "건강장기요양", "건강보험수기입력")
_register("고용보험", "고용보험", "고용보험료", "고용", "실업급여", "고용보험수기입력")
_register(
    "장기요양보험료",
    "장기요양보험료", "장기요양보험", "장기요양", "장기",
    # "요양보험"("장기" 없이) — 실사 확인(업체022): 4대보험을
    # 국민연금/건강보험/요양보험/고용보험 순서로 나열하는 원본이 실제로 있고,
    # 이 문맥에서 "요양보험"은 예외 없이 장기요양보험료를 줄여 부르는 말이다
    # (건강보험 바로 옆, 4대보험 항목 사이에 있음). "요양수당" 그룹이 이미
    # "요양"이라는 짧은 별칭을 쓰고 있지만 그건 별개의 문자열("요양" vs
    # "요양보험")이라 겹치지 않는다.
    "요양보험", "요양보험료",
)
_register("소득세", "소득세", "갑근세", "근로소득세", "갑종근로소득세", "사업소득세", "소민세", "갑근세소득세")
_register("지방소득세", "지방소득세", "지방세", "주민세", "지소세", "주민세지방소득세")
_register("학자금상환액", "학자금상환액", "학자금", "학자금상환")
_register("식대공제", "식대공제", "식대차감", "식비공제", "식사공제", "식사비공제", "직원식대공제", "식대비공제", "중식공제", "급식비공제")
# 장기근속장려금 사업주 몫을 직원 급여에서 떼는 공제 — 사업장마다 이름이 제각각이다
# (2026-10-07 정답 파일 대조: 업체023/업체024/업체025/업체026/업체015/업체027/
# 업체028/업체029 전부 위하고 "장기근속수당공제" 열에 들어가 있었음).
_register(
    "장기근속수당공제",
    "장기근속수당공제", "장기근속공제", "장기수당공제", "근속장려금공제", "장기근속공제액",
    "장기근속장려금공제", "근속수당공제", "장기근속수당공제액",
)
_register("선지급", "선지급", "선급금", "가지급금")
_register(
    "기타공제",
    "기타공제액", "공제기타",
    "가족요양본인부담금", "본인부담금",
)
# 두리누리/두루누리 지원액은 destination이 자체 열(두루누리지원액)을 따로
# 두는 경우가 많아 위 "기타공제"에 뭉뚱그리지 않고 별도 canon으로 뺀다 —
# 실사 확인(업체013): 원본엔 "두리누리"(오타/변형 표기)만 있고
# canon 사전에 전혀 등록이 안 돼 있어서 GroupOfSrc가 공제로도 분류를 못 해
# 항목매핑결과에 뜨지도 않고 통째로 누락되고 있었다.
_register("두루누리지원액", "두루누리지원액", "두리누리", "두루누리", "두루누리지원금", "두루누리사회보험료지원금")
# "두루누리국민연금"/"두루누리연금"/"두루누리고용보험"/"두루누리고용"처럼
# 4대보험 종류별로 나뉜 두루누리 감면분은 일부러 canon을 등록하지 않는다
# (예전엔 "기타공제"의 별칭이었음) — 두루누리지원액(총액)과도, 기타공제와도
# 완전히 같은 개념이 아닌데 자동으로 그 중 하나에 합쳐지면 회계사가 실제로
# 뭐가 반영됐는지 알 수 없다. canon 미등록 상태로 두면 group_of_src가
# 공제로는 분류하되 canon_of_src는 원문 그대로를 반환해서 destination에
# 똑같은 이름의 열이 없는 한 자동 매칭되지 않고, 항목매핑결과에 "미매칭"으로
# 남아 사용자가 직접 매칭 UI로 정확히 고를 수 있다(사용자 요청, 2026-08-06).
_register("공제액계", "공제합계", "공제액계", "공제총액")
_register("차인지급액", "차인지급액", "실수령액", "실지급액", "지급액", "차감지급액", "지급급여", "지급급여액", "지급금액")


_TAX_MARK_PAREN_RE = re.compile(r"\s*[(（]\s*(비|과|비과세|과세)\s*[)）]\s*$")
_LEADING_TAX_MARK_RE = re.compile(r"^\s*[(（]\s*(비|과)\s*[)）]\s*")
_INSURANCE_BASES = ("국민연금", "건강보험", "고용보험", "장기요양보험료", "장기요양보험", "장기요양")
_TRAILING_NUMBER_RE = re.compile(r"\d+$")
_CIRCLED_RE = re.compile("[①-⑳㉠-㉻Ⓐ-ⓩ]")


def _canon_key(text: object) -> str:
    raw = strip_formula_paren_suffix(str(text if text is not None else ""))
    # "식대(비)"/"연장수당(과)"/"(비)식대"처럼 과세 구분을 한 글자 괄호로 붙인
    # 위하고 열 이름 — norm()이 괄호를 지우면 "식대비"처럼 다른 단어가 돼
    # 버려서 원문 단계에서 먼저 떼어낸다(2026-10-07 정답 파일 대조:
    # 업체030/업체031/업체032/업체033 등).
    raw = _TAX_MARK_PAREN_RE.sub("", raw)
    raw = _LEADING_TAX_MARK_RE.sub("", raw)
    # "기 본 급 ①", "급여합계 (①+②+③)" — 양식이 계산 순서를 동그라미 숫자로
    # 붙여둔 경우(업체034/업체035, 2026-10-07).
    raw = _CIRCLED_RE.sub("", raw)
    k = strip_share_suffix(strip_fixed_suffix(strip_tax_suffix(strip_rate_suffix(norm(raw)))))
    k = k.rstrip("+")
    # "고용보험(0.9/100)"처럼 4대보험 이름 뒤에 요율 숫자만 남은 경우.
    for base in _INSURANCE_BASES:
        if k.startswith(base) and k[len(base):].isdigit():
            return base
    return k


def canon(text: object) -> str:
    """Maps a raw source/destination item-name string to its canonical
    Wehago item name. Order of checks mirrors the VBA `Canon` function
    exactly — several rules (근로자의날 vs 노동절, 자가운전 vs 차량유지비,
    numbered 기타수당N/기타공제N) must run before the plain alias table."""
    k = _canon_key(text)

    if k in ("국민연금", "국민연금보험료", "국민연금료"):
        return "국민연금"
    if k in ("건강보험", "건강보험료"):
        return "건강보험"
    if k in ("장기요양", "장기요양보험", "장기요양보험료"):
        return "장기요양보험료"
    if k in ("고용보험", "고용보험료"):
        return "고용보험"
    if "근로자의날" in k or "노동절" in k:
        return "근로자의날"
    if "휴일근로" in k or "휴일수당" in k:
        return "휴일근로수당"
    if "선임" in k and "요양보호사" in k:
        return "선임요양보호사"
    if "연말정산" in k:
        return "연말정산"
    if is_vehicle_allowance_key(k):
        return "차량유지비"
    if is_numbered_base(k, "기타수당"):
        return "기타수당"
    if is_numbered_base(k, "기타공제"):
        return "기타공제"

    return _EXACT_CANON.get(k, k)


# Some aliases above are grouped under a broader canon by convention, but
# aren't guaranteed to mean the same thing at every client — "교통비" is
# usually the same as 차량유지비 for this app's care-worker payroll domain,
# but could be a separate commute stipend elsewhere; "직무수당"/"직급수당"
# might be a genuinely different allowance from 직책수당 depending on the
# organization; "사업소득세" is a different withholding regime from
# 근로소득세/갑근세 even though both end up in the same "소득세" deduction
# slot. Auto-matching still applies (refusing to match at all would be
# worse — see the "미매칭" philosophy elsewhere in this module), but the UI
# surfaces these with a yellow "확인 필요" flag instead of treating them as
# certain as an exact/near-exact synonym match (사용자 요청, 2026-08-07).
_AMBIGUOUS_ALIASES = {
    "교통비",
    "직무수당", "직급수당",
    "사업소득세",
    "가족요양본인부담금", "본인부담금",
    "식대비", "비식대", "건강장기요양", "기본수당",
}


def is_ambiguous_alias(text: object) -> bool:
    """True if `text` normalizes to one of `_AMBIGUOUS_ALIASES` — same
    normalization pipeline as `canon()` so this reflects what actually got
    matched, not the raw cell text."""
    return _canon_key(text) in _AMBIGUOUS_ALIASES
