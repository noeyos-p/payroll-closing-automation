"""Coverage smoke test for core/canon.py's synonym dictionary against the
broad set of real Wehago upload samples collected in 세무관련/파일/ (~140
real client files, gathered specifically to stress-test header-text
recognition across many different registered item names — see the root
CLAUDE.md). Those files live outside this repo on purpose (they're real
clients' payroll data — 주민등록번호 and all — so they're never copied into
tests/fixtures/); this test reads them in place and is skipped entirely if
that folder isn't present (e.g. a fresh checkout without the sibling
data folder).

A column with real money in it that `group_of_dst` can't classify into
수당/공제/차인지급액 means core/reconcile.py would silently ignore that
column when reconciling that client — exactly the kind of "이 파일은 파일마다
서식이 다르다" gap the user asked this port to be precise about."""

from __future__ import annotations

import unicodedata
import warnings
from pathlib import Path

import pytest

openpyxl = pytest.importorskip("openpyxl")

from core.classify import group_of_dst  # noqa: E402
from core.sheet_layout import detect_dest_layout  # noqa: E402

SAMPLE_DIR = Path(__file__).resolve().parents[2] / "파일"

pytestmark = pytest.mark.skipif(not SAMPLE_DIR.is_dir(), reason="파일/ 실사 표본 폴더가 없음(별도 보관, 이 저장소에 포함 안 함)")


def _sample_files() -> list[Path]:
    if not SAMPLE_DIR.is_dir():
        return []
    # macOS/APFS returns filenames NFD-normalized (Hangul syllables stored
    # decomposed) — a plain "in" check against an NFC literal never matches,
    # so every 사원등록(employee-roster, not payroll) file was silently
    # included until this was normalized. Those files legitimately have
    # 사원코드/사원명 headers (they pass detect_dest_layout) but their other
    # columns are demographic/insurance-base fields, not payroll items — real
    # money there being "unclassified" is expected and not a canon.py gap.
    return sorted(p for p in SAMPLE_DIR.glob("*.xlsx") if "사원등록" not in unicodedata.normalize("NFC", p.name))


def test_every_column_with_real_money_classifies_into_a_known_group() -> None:
    warnings.filterwarnings("ignore")
    unclassified: dict[str, set[str]] = {}
    files_checked = 0

    for path in _sample_files():
        try:
            workbook = openpyxl.load_workbook(path, data_only=True)
        except Exception:
            continue
        names = workbook.sheetnames
        candidate_names = (["Sheet1"] if "Sheet1" in names else []) + [n for n in names if n != "Sheet1"]
        for name in candidate_names:
            worksheet = workbook[name]
            layout = detect_dest_layout(worksheet)
            if layout is None:
                continue
            files_checked += 1
            last_row = min(worksheet.max_row, layout.data_start_row + 400)
            for c in range(1, worksheet.max_column + 1):
                if group_of_dst(worksheet, layout, c):
                    continue
                has_money = False
                for r in range(layout.data_start_row, last_row + 1):
                    v = worksheet.cell(row=r, column=c).value
                    if isinstance(v, (int, float)) and abs(v) >= 1:
                        has_money = True
                        break
                if has_money:
                    from core.sheet_layout import dst_full_header

                    label = dst_full_header(worksheet, layout, c)
                    unclassified.setdefault(str(path.name), set()).add(label)
            break  # only the first recognized sheet per workbook

    assert files_checked > 50, f"실사 표본을 충분히 못 찾음 (찾은 시트 수: {files_checked})"
    assert not unclassified, f"금액이 있는데 그룹 분류 안 된 열 발견: {unclassified}"
