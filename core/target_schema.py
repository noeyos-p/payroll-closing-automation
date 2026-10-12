"""Wehago upload column layout. Unlike our first assumption, this is NOT a
single universal format — each client has their own registered set of
allowance/deduction items in Wehago (confirmed by comparing two different
clients' real upload files, which had almost no overlapping column names).
So a ClientProfile carries its own `allowance_columns`/`deduction_columns`;
this module only fixes the parts that are truly constant across every
client: the 2 identity columns (사원코드/사원명 — 부서/직급/직종 aren't
needed) and the computed subtotal columns placed right after their group:
비과세합계/과세합계/지급액계 after allowances, 공제액계 after deductions,
and 차인지급액 last."""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, Literal

if TYPE_CHECKING:
    from core.wehago_schema import WehagoRecord

Group = Literal["identity", "allowance", "deduction", "total"]


@dataclass(frozen=True)
class ColumnDef:
    key: str
    label: str
    group: Group
    # Only meaningful for group="allowance": whether this item is taxable
    # (기본급, most 수당) or non-taxable (비과세 — e.g. 식대, 자가운전 for
    # some clients). Differs per client, so it's set explicitly per column
    # rather than inferred from the label.
    taxable: bool = True


IDENTITY_COLUMNS: list[ColumnDef] = [
    ColumnDef("사원코드", "사원코드", "identity"),
    ColumnDef("사원명", "사원명", "identity"),
]


def build_target_column_order(
    allowance_columns: list[ColumnDef], deduction_columns: list[ColumnDef]
) -> list[ColumnDef]:
    return (
        IDENTITY_COLUMNS
        + allowance_columns
        + [
            ColumnDef("비과세합계", "비과세합계", "total"),
            ColumnDef("과세합계", "과세합계", "total"),
            ColumnDef("지급액계", "지급액계", "total"),
        ]
        + deduction_columns
        + [ColumnDef("공제액계", "공제액계", "total")]
        + [ColumnDef("차인지급액", "차인지급액", "total")]
    )


def drop_all_zero_columns(
    columns: list[ColumnDef], records: list["WehagoRecord"], group: Literal["allowance", "deduction"]
) -> list[ColumnDef]:
    """A client's Wehago profile registers a fixed set of allowance/deduction
    items, but not every item has an amount every month. Columns with a zero
    value for every employee this run carry no information, so they're
    dropped from both the review grid and the exported file — not just left
    as a column of zeros."""

    def value_of(record: "WehagoRecord", key: str) -> float:
        bucket = record.allowances if group == "allowance" else record.deductions
        return bucket.get(key, 0)

    return [c for c in columns if any(value_of(r, c.key) for r in records)]


GRAND_TOTAL_LABEL = "합계"


def compute_grand_totals(columns: list[ColumnDef], records: list["WehagoRecord"]) -> dict[str, object]:
    """Sums every amount column (수당/공제/합계) across all employees for a
    single grand-total row appended at the bottom of the review grid and the
    exported file. Identity columns get a "합계" label instead of a sum."""
    totals: dict[str, object] = {}
    for i, col in enumerate(columns):
        if col.group == "identity":
            totals[col.key] = GRAND_TOTAL_LABEL if i == 0 else ""
        else:
            totals[col.key] = sum(r.column_value(col.key) for r in records)
    return totals
