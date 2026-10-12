"""core/hwp_reader.py: extracting tables out of a real .hwp payroll ledger
into the same Grid shape an Excel worksheet already satisfies. This only
locks in the STRUCTURAL extraction — column alignment through merged cells
(rowspan/colspan), continuation rows, block detection. The .hwp path isn't
wired into core/engine.py's reconciliation pipeline yet (this round of work
is Excel-only — see CLAUDE.md), so there's no value-matched output test
here; this file just guards the extraction step itself from regressing.
"""

from pathlib import Path

from core.hwp_reader import list_hwp_tables, pick_largest_table

FIXTURES = Path(__file__).parent / "fixtures"
SAMPLE_PATH = FIXTURES / "업체045 5월귀속 6월지급 급여대장.hwp"


def test_list_hwp_tables_finds_the_payroll_table_and_a_small_signoff_box():
    tables = list_hwp_tables(SAMPLE_PATH)
    assert len(tables) == 2
    sizes = sorted((t.max_row, t.max_column) for t in tables)
    assert sizes[0] == (2, 3)  # 결재/담당자/센터장 signoff box
    assert sizes[1] == (25, 14)  # the real payroll table


def test_pick_largest_table_returns_the_payroll_table():
    tables = list_hwp_tables(SAMPLE_PATH)
    largest = pick_largest_table(tables)
    assert largest is not None
    assert largest.max_row == 25
    assert largest.max_column == 14


def test_merged_header_cells_land_in_their_real_columns():
    # The header spans 3 physical rows via rowspan/colspan (a "세금공제"
    # group header over several sub-item columns) — a naive "one row = one
    # list of <td> texts" read would misalign every column from the merge
    # onward. Confirms core/hwp_reader.py's rowspan/colspan-aware grid
    # construction keeps every column aligned to its true position.
    tables = list_hwp_tables(SAMPLE_PATH)
    payroll = pick_largest_table(tables)
    assert payroll.cell(1, 1).value == "연 번"
    assert payroll.cell(1, 2).value == "성 명"
    assert payroll.cell(4, 1).value == "1"
    assert payroll.cell(4, 2).value == "사원017"
    # row 6 is 사원018's main payroll row; row 7 is his continuation line
    # (국민연금/건강보험/고용보험 detail) — same column positions as the
    # header's sub-item labels above.
    assert payroll.cell(6, 2).value == "사원018"
    assert payroll.cell(6, 4).value == "2,440,000"  # 기본금
