"""Total computation shared by the review grid (live preview) and export —
지급액계/공제액계/차인지급액/비과세합계/과세합계 are always computed by the
app, never trusted from the source file or left as live Excel formulas in
the output."""

from __future__ import annotations


def compute_totals(allowances: dict[str, float], deductions: dict[str, float]) -> tuple[float, float, float]:
    지급액계 = sum(allowances.values())
    공제액계 = sum(deductions.values())
    차인지급액 = 지급액계 - 공제액계
    return 지급액계, 공제액계, 차인지급액


def compute_taxable_split(allowances: dict[str, float], non_taxable_keys: set[str]) -> tuple[float, float]:
    """Returns (비과세합계, 과세합계) — which allowance items are non-taxable
    differs per client, so the caller (core/mapping.py) supplies the set of
    keys marked non-taxable in that client's profile."""
    비과세합계 = sum(v for k, v in allowances.items() if k in non_taxable_keys)
    과세합계 = sum(v for k, v in allowances.items() if k not in non_taxable_keys)
    return 비과세합계, 과세합계
