"""ClientProfile — a pure output-schema container for ONE run: which
allowance/deduction columns this client's Wehago sheet has, in what order.
Nothing here is ever saved to disk (see CLAUDE.md — no per-client
registration in this app at all).

Trimmed down from an earlier value-cross-reference-matching design (which
also carried a `row_pattern`/`field_mappings`/`employee_roster` describing
how to scan the source file) once core/reconcile.py replaced that matching
approach with header-text canonicalization ported from the VBA macro —
core/sheet_layout.py's `SourceLayout`/`DestLayout` now carry the "how to
read this file" information instead, and `WehagoRecord` objects already
carry each employee's own data, so no separate roster is needed here."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime, timezone

from core.target_schema import ColumnDef


@dataclass
class ClientProfile:
    id: str
    client_name: str
    version: int
    created_at: str
    updated_at: str
    allowance_columns: list[ColumnDef] = field(default_factory=list)
    deduction_columns: list[ColumnDef] = field(default_factory=list)


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


_ACCRUAL_MONTH_RE = re.compile(r"(\d{1,2})\s*월\s*귀속")


def derive_client_label(filename_stem: str) -> str:
    """Best-effort 거래처 표시용 label guessed from a 급여대장 원본 filename,
    e.g. "업체022 7월귀속 8월지급 급여대장" -> "업체02207".
    This app's real client files follow "{거래처명} N월귀속 M월지급
    급여대장" (confirmed pattern across real files), and this client's
    위하고 reference file for the same month is separately named
    "{거래처명}-YYYYMM", using the same 귀속월(accrual month) N. Mirrors that
    convention without the dash/year so the accountant doesn't have to
    retype it by hand every month (사용자 요청, 2026-08-07) — falls back to
    the first whitespace-separated token (a coarser guess) when the
    "N월귀속" pattern isn't present in the filename at all."""
    match = _ACCRUAL_MONTH_RE.search(filename_stem)
    if match:
        prefix = filename_stem[: match.start()].strip()
        if prefix:
            return f"{prefix}{int(match.group(1)):02d}"
    return filename_stem.split(" ")[0]


def guess_company_name_from_filename(filename_stem: str) -> str:
    """`derive_client_label`과 같은 "{거래처명} N월귀속..." 관례에서 거래처명
    부분만 뽑는다(월 접미사 없이) — 등록된 사업장 이름과 정확히 대조하는 용도.
    사용자 피드백: 실제 파일들이 이 관례를 정확히 따르는데(예: "홍길동업체
    7월귀속 8월지급 급여대장.xlsx") 사업장 자동 인식이 사원 명부 겹침만
    보고 파일명은 전혀 안 써서 놓치는 경우가 있었다 — core/company_match.py의
    사원 명부 매칭보다 먼저, 파일명이 등록된 사업장명과 정확히 일치하는지부터
    확인하는 데 쓴다(core/batch.py)."""
    match = _ACCRUAL_MONTH_RE.search(filename_stem)
    if match:
        prefix = filename_stem[: match.start()].strip()
        if prefix:
            return prefix
    return filename_stem.split(" ")[0]


# 같은 귀속월 급여를 "추급"(추가지급)/"급여"(정기)/"급+상"(급여+상여) 세
# 종류로 나눠 처리하는 실사 관례가 확인됨(사용자 요청, 2026-09-08) — 한
# 워크북에 세 시트가 다 있는 경우도 있고, 파일 자체가 셋으로 나뉘어 각
# 파일명 끝에 "...급여대장(급여)"처럼 종류가 붙어있는 경우도 있다. 순서는
# 중요하지 않다("급+상"이 "급여"의 부분문자열이 아니라 서로 겹치지 않음).
PAYROLL_TYPE_TOKENS = ("추급", "급+상", "급여")

# 시트 이름에는 "+"가 빠진 "급상"으로 쓰이는 실사 사례가 확인됨(업체050
# 복지센터, 2026-09-08 — 파일명은 "...(급+상)(추급)"인데 실제 시트 이름은
# "급상"/"추급"). 파일명 쪽 관례("급+상")와 표시를 통일하려고, 인식은
# "급상"도 받아주되 반환값은 항상 정식 표기 "급+상"으로 정규화한다.
_SHEET_NAME_ALIASES = {"급상": "급+상"}


def _sheet_name_matches_token(name: str, token: str) -> bool:
    """`name`이 `token` 자체이거나, 뒤에 순번 숫자만 붙은 형태("추급1"/
    "추급2")인지 — 딱 이 형태만 인정한다. `token in name`(부분문자열)
    이었다면 "급여계산"/"급여대장"처럼 종류 구분과 무관한 실제 표 이름도
    전부 "급여"로 오인식된다(업체051 실사, 2026-09-10 — 이
    워크북의 구조가 다른 두 후보 시트 "급여계산"/"급여대장"이 이 오탐으로
    "종류가 다른 시트 2개"로 잘못 합쳐져서, 서로 다른 두 표의 항목이 한
    항목매핑결과 화면에 전부 뒤섞여 뜨고 수동 매칭 적용도 열 번호 충돌로
    안 먹혔다)."""
    if name == token:
        return True
    return name.startswith(token) and name[len(token):].isdigit()


def payroll_type_from_sheet_name(sheet_name: str) -> str | None:
    """워크북 하나에 추급/급여/급+상 시트가 나뉘어 있는 경우, 시트 이름만
    보고 어떤 종류인지 판단한다(core/batch.py의 `sheets_for_batch_add`가
    이걸로 한 파일을 몇 개의 큐 항목으로 나눌지 결정하고, 다운로드 파일명
    라벨에도 그대로 쓴다). 시트 이름 전체가 종류 표시 그 자체(+순번 숫자)일
    때만 인정 — 부분문자열 포함으로는 판단하지 않는다(`_sheet_name_matches_token`
    참고)."""
    name = sheet_name.strip()
    token = parse_template_token(name)
    if token is not None:
        return token or None
    for alias, canonical in _SHEET_NAME_ALIASES.items():
        if _sheet_name_matches_token(name, alias):
            return canonical
    for token in PAYROLL_TYPE_TOKENS:
        if _sheet_name_matches_token(name, token):
            return token
    return None


_PAYROLL_TYPE_FILENAME_RE = re.compile(r"\((추급|급\+상|급상|급여)\d*\)")


def payroll_type_from_filename(filename_stem: str) -> str | None:
    """"업체052 7월귀속 8월지급 급여대장(급여)"처럼 파일명 끝에 붙은 종류
    표시를 그대로 뽑는다 — 추급/급여/급+상이 각각 별도 파일로 나뉘어 있는
    경우, 원본 파일명의 이 괄호를 다운로드 파일명에도 그대로 옮겨 붙이는
    데 쓴다. 괄호가 여러 개면(드묾) 마지막 것을 쓴다. "(추급1)"/"(추급2)"
    처럼 뒤에 번호가 붙어도(파일 자체가 추급1/추급2로 나뉜 실사 사례,
    2026-09-08) 종류만 뽑는다 — 번호까지 남으면 "추급1"≠"추급2"로 서로
    다른 종류 취급돼서 합쳐져야 할 두 파일이 안 합쳐진다. "급상"(플러스
    없이)도 시트 이름과 동일하게 받아 "급+상"으로 정규화한다."""
    matches = _PAYROLL_TYPE_FILENAME_RE.findall(filename_stem)
    if not matches:
        return None
    token = matches[-1]
    return _SHEET_NAME_ALIASES.get(token, token)


def payroll_types_in_filename(filename_stem: str) -> list[str]:
    """파일명 괄호 속 종류 전부(순서 유지, 중복 제거) — "(급여)(급+상)"이면
    ["급여", "급+상"]."""
    out: list[str] = []
    for token in _PAYROLL_TYPE_FILENAME_RE.findall(filename_stem):
        t = _SHEET_NAME_ALIASES.get(token, token)
        if t not in out:
            out.append(t)
    return out


# 원본 파일 하나에 급여/급+상이 같이 들어 있고, 그 사업장에 종류별 수정본
# 양식 변환 규칙(core/source_template.py)이 있으면 큐에서 종류별 항목으로
# 나눈다. 이때 시트 이름 자리에 이 표시를 넣어 넘긴다(core/batch.py).
_TEMPLATE_TOKEN_RE = re.compile(r"^<양식:(.*)>$")


def template_token(payroll_type: str | None) -> str:
    return f"<양식:{payroll_type or ''}>"


def parse_template_token(sheet_name: str | None) -> str | None:
    """`template_token`이 만든 표시면 그 종류(""일 수 있음), 아니면 None."""
    if not sheet_name:
        return None
    m = _TEMPLATE_TOKEN_RE.match(sheet_name.strip())
    return m.group(1) if m else None
