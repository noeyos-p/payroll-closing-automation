"""Extracts tables out of a .hwp (Hangul Word Processor 5.x binary format)
file into the same Grid abstraction an Excel worksheet already satisfies.

`hwp5txt` (pyhwp's plain-text converter) replaces every table with a bare
"<표>" placeholder and drops the cell contents entirely — confirmed on a
real payroll ledger (업체045). `hwp5html`'s HTMLTransform converts
to XHTML with real <table>/<tr>/<td> elements instead, so that's what we
use — via pyhwp's own Python API directly (`Hwp5File` + `HTMLTransform`),
NOT by shelling out to the `hwp5html` command-line script. A subprocess call
to a separate console-script executable is fragile once this app is frozen
with PyInstaller for Windows distribution: that script only exists as a
standalone .exe if pip installed it onto PATH on the *build* machine, and
PyInstaller does not automatically bundle a sibling package's console-script
entry point just because the package itself is imported — calling the
library function in-process sidesteps all of that.

HWP tables commonly merge cells (a 2-row sub-header under one payroll
category, e.g. "세금공제" spanning "국민연금"/"건강보험"/"소득세" columns
below it) — an HTML <td rowspan="2"> or colspan="2"> — so a naive "one row
= one list of <td> texts" read would misalign every column from the merge
onward. _table_to_grid tracks which (row, col) positions a rowspan/colspan
cell occupies so every later cell still lands in its true column.
"""

from __future__ import annotations

import tempfile
from contextlib import closing
from pathlib import Path

from hwp5.hwp5html import HTMLTransform
from hwp5.xmlmodel import Hwp5File
from lxml import html

from core.grid import Grid


def _cell_text(cell) -> str:
    return " ".join("".join(cell.itertext()).split())


def _table_to_grid(table_element) -> Grid:
    occupied: set[tuple[int, int]] = set()
    row_values: list[dict[int, str]] = []

    for row_idx, tr in enumerate(table_element.xpath(".//tr")):
        values: dict[int, str] = {}
        col_idx = 0
        for cell in tr.xpath("./td|./th"):
            while (row_idx, col_idx) in occupied:
                col_idx += 1
            rowspan = int(cell.get("rowspan") or 1)
            colspan = int(cell.get("colspan") or 1)
            values[col_idx] = _cell_text(cell)
            for dr in range(rowspan):
                for dc in range(colspan):
                    occupied.add((row_idx + dr, col_idx + dc))
            col_idx += colspan
        row_values.append(values)

    width = max((max(v.keys()) + 1 if v else 0 for v in row_values), default=0)
    rows = [[values.get(c, "") for c in range(width)] for values in row_values]
    return Grid(rows)


def list_hwp_tables(path: Path) -> list[Grid]:
    """Every table in the document, in document order, each as a Grid."""
    with tempfile.TemporaryDirectory() as tmp_dir:
        with closing(Hwp5File(str(path))) as hwp5file:
            HTMLTransform().transform_hwp5_to_dir(hwp5file, tmp_dir)
        tree = html.parse(str(Path(tmp_dir) / "index.xhtml"))
        return [_table_to_grid(t) for t in tree.xpath("//table")]


def pick_largest_table(tables: list[Grid]) -> Grid | None:
    """Best-effort default when nothing else has picked a table yet: a real
    payroll table is virtually always the largest one in the document (a
    signoff/approval box like "결재/담당자/센터장" is small and always
    present alongside it — confirmed on a real sample). Callers that can
    verify against reference values should still try every table and keep
    whichever one actually reconciles, the same way core/heuristic_match.py
    never trusts a single guess without checking it — this is only a
    fallback ordering hint, not a decision."""
    if not tables:
        return None
    return max(tables, key=lambda g: g.max_row * g.max_column)
