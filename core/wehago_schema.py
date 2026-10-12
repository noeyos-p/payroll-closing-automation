"""Generic Wehago output record. allowances/deductions are plain dicts keyed
by whatever a given ClientProfile declares in its allowance_columns /
deduction_columns (core/target_schema.py) — the specific set of allowance
and deduction items differs per client, so no fixed key list lives here."""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class WehagoRecord:
    사원코드: str = ""
    사원명: str = ""
    부서: str = ""
    직급: str = ""
    직종: str = ""
    allowances: dict[str, float] = field(default_factory=dict)
    deductions: dict[str, float] = field(default_factory=dict)
    비과세합계: float = 0
    과세합계: float = 0
    지급액계: float = 0
    공제액계: float = 0
    차인지급액: float = 0
    # Populated when the source ledger's own stated total for 지급액계/
    # 공제액계/차인지급액 disagrees with our recomputed total — the source
    # file's own totals are not always correct, so we always recompute from
    # the raw line items, but still flag when they diverge for the
    # accountant to double-check.
    mismatches: list[str] = field(default_factory=list)
    # Which specific total field(s) ("지급액계"/"공제액계"/"차인지급액") the
    # mismatch above is about — lets the review grid highlight only the
    # exact cell(s) that disagree with the source instead of the whole row.
    mismatched_fields: set[str] = field(default_factory=set)
    # The source ledger's own stated value for each total field
    # (core/reconcile.py's `_apply_source_totals_check`, comparing against
    # the source's own 지급/공제/차인지급액 columns) — populated whenever the
    # source has a usable number there, whether or not it agrees with our
    # recomputed total, so the review grid can show "원본 X계" right next to
    # our "X계" for a direct side-by-side comparison.
    source_totals: dict[str, float] = field(default_factory=dict)

    def column_value(self, key: str) -> str | float:
        if key in (
            "사원코드", "사원명", "부서", "직급", "직종",
            "비과세합계", "과세합계", "지급액계", "공제액계", "차인지급액",
        ):
            return getattr(self, key)
        if key in self.allowances:
            return self.allowances[key]
        if key in self.deductions:
            return self.deductions[key]
        return 0
