"""급여대장(원본)만으로 어느 사업장인지 자동 판별한다.

등록된 사업장들의 사원 이름 지문(core/registry.py)과 급여대장에서 뽑은 사원
이름을 대조한다 — 엄격 모드: 1위 일치율이 임계값을 넘고, 2위와 확실히 격차가
나야 자동 인식한다. 애매하면 절대 추측하지 않고 "확인 필요"로 멈춘다(이 앱
전체의 원칙과 동일 — core/reconcile.py의 "미매칭" 철학과 같은 결).

이름 정규화는 매칭 엔진 본체(core/reconcile.py)가 이미 쓰는
core/name_match.py의 `name_base3`를 그대로 재사용한다 — 별도 정규화 규칙을
새로 만들지 않아서, 사업장 인식과 사원 매칭이 항상 같은 기준으로 이름을 본다.
"""
from __future__ import annotations

from dataclasses import dataclass, field

from core.name_match import name_base3

_DEFAULT_MIN_RATIO = 0.6
_DEFAULT_MIN_GAP = 0.15


@dataclass
class CompanyRoster:
    company_id: int
    company_name: str
    base_names: set[str]


@dataclass
class CompanyCandidate:
    company_name: str
    ratio: float


@dataclass
class CompanyMatchResult:
    status: str  # "matched" | "ambiguous" | "no_match"
    company_id: int | None
    company_name: str | None
    candidates: list[CompanyCandidate] = field(default_factory=list)


def match_company(
    ledger_names: list[str],
    rosters: list[CompanyRoster],
    min_ratio: float = _DEFAULT_MIN_RATIO,
    min_gap: float = _DEFAULT_MIN_GAP,
) -> CompanyMatchResult:
    """`ledger_names`(급여대장에서 뽑은 이름들) 중 몇 %가 각 등록 사업장
    로스터에 있는지로 순위를 매긴다. 1위가 `min_ratio` 미만이면 no_match,
    1위와 2위 격차가 `min_gap` 미만이면(둘 다 비슷하게 그럴듯함) ambiguous."""
    ledger_bases = {b for n in ledger_names if (b := name_base3(n))}

    scored: list[tuple[CompanyRoster, float]] = []
    for roster in rosters:
        if not ledger_bases or not roster.base_names:
            ratio = 0.0
        else:
            ratio = len(ledger_bases & roster.base_names) / len(ledger_bases)
        scored.append((roster, ratio))
    scored.sort(key=lambda pair: pair[1], reverse=True)

    candidates = [CompanyCandidate(r.company_name, ratio) for r, ratio in scored[:5]]

    if not scored or scored[0][1] < min_ratio:
        return CompanyMatchResult("no_match", None, None, candidates)
    if len(scored) >= 2 and scored[0][1] - scored[1][1] < min_gap:
        return CompanyMatchResult("ambiguous", None, None, candidates)

    best = scored[0][0]
    return CompanyMatchResult("matched", best.company_id, best.company_name, candidates)
