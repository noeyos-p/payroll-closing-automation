"""Broad smoke test against real files: this month's 이지케어 원본 급여대장
(사원002(6월지급 급여입력)/, one real file per client) paired with that same
client's existing 위하고 Sheet1 (세무관련/파일/) by matching filenames. Both
folders live outside this repo (real client data) and this test is skipped
if either is missing.

This is what actually caught three real detection bugs while validating
this port against genuinely separate two-file inputs (as opposed to the
golden fixtures in tests/fixtures/golden/, which were split out of a single
combined workbook) — see CLAUDE.md for the full story of each:
  1. A 4대보험 rate row (raw floats like 0.0475) sitting between the source
     header and the first data row got read as an item label, shadowing the
     real one above it (업체063).
  2. A per-employee pay-stub sheet ("급여명세서") in the same workbook as
     the real table satisfied the name-only fallback anchor just as well as
     the real table, and out-ranked it (업체043).
  3. A "원본" file that was itself already Wehago-shaped (사원코드/사원명
     header) passed the same fallback and got misread as a source ledger
     (업체060 — this one is still an open case: the file
     genuinely isn't a supported source shape, so `process()` correctly
     returns None rather than guessing).
  4. No continuation-row support at all — 업체043's 급여대장 sheet
     splits each employee across 2 physical rows (name/생년월일/기본급 on
     one row, 담당업무/입사일 on the next, no marker on the second) with no
     "No" header text, just a real 1..N sequence at a strict 2-row gap.
     core/sheet_layout.py's `_detect_value_sequence_marker` now finds that
     kind of unlabeled-but-real sequence and core/reconcile.py sums each
     block's rows together instead of misreading job titles as employees.
  5. A destination sheet with the SAME raw label ("식대") used for both an
     allowance column and a deduction column (업체013's Sheet1 —
     an allowance "식대" and a deduction "식대", right next to a third
     deduction column literally labeled "식대공제"). `WehagoRecord` resolves
     a key by checking `allowances` before `deductions`, so with identical
     keys in both dicts the deduction column's amount was silently folded
     into the allowance total instead — 지급액계 quietly inflated and
     공제액계 understated by the same amount, for every employee who had
     that deduction. `core/reconcile.py`'s `_dst_column_schema` now
     disambiguates labels across BOTH groups together so this can't happen.
     Two more real gaps surfaced by the same file: "연장근무"/"운전보조"
     weren't recognized as 연장근로수당/차량유지비 synonyms, and 4대보험
     columns labeled with an embedded rate ("국민연금(4.5%)") weren't
     recognized at all (core/canon.py's `strip_rate_suffix`).
  6. Two more label gaps (업체056): "연번" instead of "No"/"순번"
     for the marker column, "종사자명" instead of "성명"/"이름" for the name
     column — neither was in `core/sheet_layout.py`'s label sets at all, so
     `detect_source_layout` failed outright (not a false match, just no
     candidate). Separately, that same workbook has TWO other sheets that
     are legitimately not payroll ledgers (a bank-transfer account list with
     no amount columns, and a "확인용" summary sheet with only one lump-sum
     total per employee, no line-item breakdown) — `process()` correctly
     returns None for those; that's not a bug to route around.
"""

from __future__ import annotations

import re
import unicodedata
import warnings
from pathlib import Path

import pytest

openpyxl = pytest.importorskip("openpyxl")

from core.engine import process  # noqa: E402
from core.sheet_layout import detect_source_layout  # noqa: E402

SOURCE_DIR = Path("/Users/soyeon/Desktop/세무관련/사원002(6월지급 급여입력)")
WEHAGO_DIR = Path("/Users/soyeon/Desktop/세무관련/파일")
# 업체016의 위하고 참고파일은 (업체006처럼) 파일/이
# 아니라 다운로드 폴더에 있었음 — 실사 확인.
DOWNLOADS_DIR = Path("/Users/soyeon/Downloads")

pytestmark = pytest.mark.skipif(
    not (SOURCE_DIR.is_dir() and WEHAGO_DIR.is_dir()),
    reason="실사 표본 폴더가 없음(별도 보관, 이 저장소에 포함 안 함)",
)

# Known, explained non-matches — see module docstring point 3 and the
# multi-sheet ambiguity note in the session summary (a workbook with
# several equally-plausible payroll sheets, e.g. per income-category splits,
# can pick a smaller sub-sheet instead of the main one; not a crash, just
# not always the intended candidate without a user picking the right file).
EXPECTED_NO_MATCH = {"업체060"}


def _nfc(s: str) -> str:
    return unicodedata.normalize("NFC", s)


def _client_key_src(name: str) -> str:
    return re.sub(r"\s*\d+월귀속\s*\d+월지급\s*급여대장.*\.xlsx$", "", _nfc(name)).strip()


def _client_key_dst(name: str) -> str:
    s = re.sub(r"_사원등록_\d+\.xlsx$", "", _nfc(name))
    return re.sub(r"-\d{6}\.xlsx$", "", s).strip()


def _matched_client_pairs() -> list[tuple[str, Path, Path]]:
    src_by_key: dict[str, list[Path]] = {}
    for f in SOURCE_DIR.glob("*.xlsx"):
        src_by_key.setdefault(_client_key_src(f.name), []).append(f)
    dst_by_key: dict[str, list[Path]] = {}
    for f in WEHAGO_DIR.glob("*.xlsx"):
        if "사원등록" in _nfc(f.name):
            continue
        dst_by_key.setdefault(_client_key_dst(f.name), []).append(f)

    def pick_src(paths: list[Path]) -> Path:
        def rank(p: Path) -> int:
            n = _nfc(p.name)
            if "추급" in n or "사업" in n:
                return 2
            if "급+상" in n:
                return 0
            return 1

        return sorted(paths, key=rank)[0]

    common = sorted(set(src_by_key) & set(dst_by_key))
    return [(key, pick_src(src_by_key[key]), dst_by_key[key][0]) for key in common]


def test_every_matched_client_pair_processes_without_crashing() -> None:
    warnings.filterwarnings("ignore")
    pairs = _matched_client_pairs()
    assert len(pairs) > 15, f"실사 표본에서 짝지어진 거래처를 충분히 못 찾음 ({len(pairs)}개)"

    failures = []
    for client, source_path, reference_path in pairs:
        try:
            result = process(client, source_path, reference_path)
        except Exception as exc:  # pragma: no cover - failure path
            failures.append((client, f"EXCEPTION: {exc}"))
            continue
        if result is None and client not in EXPECTED_NO_MATCH:
            failures.append((client, "예상치 못한 자동 매칭 실패"))
        if result is not None and client in EXPECTED_NO_MATCH:
            failures.append((client, "이제 매칭되므로 EXPECTED_NO_MATCH에서 빼야 함"))
    assert not failures, failures


def test_업체043_continuation_rows_are_merged_not_read_as_employees() -> None:
    warnings.filterwarnings("ignore")
    source_path = SOURCE_DIR / "업체043 6월귀속 6월지급 급여대장.xlsx"
    reference_path = WEHAGO_DIR / "업체043-202606.xlsx"
    if not source_path.exists() or not reference_path.exists():
        pytest.skip("업체043 표본 파일이 없음")

    workbook = openpyxl.load_workbook(source_path, data_only=True)
    layout = detect_source_layout(workbook["급여대장"])
    assert layout is not None
    assert layout.block_height == 2  # 순번/성명 row + 담당업무/입사일 continuation row
    assert layout.has_marker is True  # found via the real (unlabeled) 1..N sequence, not text-guessing

    result = process("업체043", source_path, reference_path)
    assert result is not None
    names = {row.name for row in result.verification_rows}
    # Job titles from the continuation row (row 2 of each block) must never
    # show up as if they were employees.
    assert not names & {"담당업무", "센터장", "사회복지사", "요양보호사", "운전원", "사무원", "간호조무사"}
    assert result.summary.matched == 13


def test_업체013_allowance_and_deduction_dont_share_a_label() -> None:
    warnings.filterwarnings("ignore")
    source_path = SOURCE_DIR / "업체013 6월귀속 6월지급 급여대장.xlsx"
    reference_path = WEHAGO_DIR / "업체078 (2026)-202606.xlsx"
    if not source_path.exists() or not reference_path.exists():
        pytest.skip("업체013 표본 파일이 없음")

    result = process("업체013", source_path, reference_path, source_sheet="24.09월급여대장")
    assert result is not None
    assert result.summary.matched == 9

    by_name = {r.사원명: r for r in result.records}
    # 김현호's source row: 기본급 2,900,000 + 차량보조금(비과세) 300,000 +
    # 식대(비과세) 300,000 + 연장근무 500,000 = 4,000,000 지급총액 exactly
    # (confirmed against the source file by hand). Before the label-collision
    # fix, the deduction-side "식대공제" (30,000) silently landed inside the
    # allowance "식대" instead of 공제액계, inflating this to 4,030,000.
    kim = by_name["김현호"]
    assert kim.지급액계 == 4_000_000
    assert kim.allowances["식대"] == 300_000
    assert kim.공제액계 == 130_500 + 102_800 + 13_310 + 75_860 + 7_580 + 30_000  # 국민연금+건강보험+장기요양+소득세+지방세+식대공제


def test_업체014_계좌리스트_시트가_아니라_진짜_급여대장이_뽑힘() -> None:
    """실사 확인: "(급+상)2.xlsx" 워크북엔 진짜 급여대장("종사자급여대장",
    직원 1명뿐) 말고 "계좌 리스트"(은행 계좌 목록) 시트도 있는데, 이 시트도
    순번+이름 컬럼 패턴이 있어 `detect_source_layout`을 통과하고, 빈 행까지
    딸린 순번 열이 길게 이어져서 행 수 기준 랭킹에서 진짜 급여대장을 이겨버
    렸다 — 수당/공제로 분류되는 열이 하나도 없는 시트(순번/은행명/계좌번호/
    이름/비고뿐)는 애초에 급여대장 후보에서 빼도록 고쳤다. 같은 파일에서
    "국민연금(본인부담)"/"건강보험(본인부담)"처럼 부담 주체별로 4대보험을
    나눈 헤더가 canon.py에 없어 전혀 인식이 안 되던 것도 같이 발견해 고쳤다
    (strip_share_suffix)."""
    warnings.filterwarnings("ignore")
    source_path = SOURCE_DIR / "업체014 5월귀속 6월지급 급여대장(급+상)2.xlsx"
    reference_path = WEHAGO_DIR / "업체014-202605.xlsx"
    if not source_path.exists() or not reference_path.exists():
        pytest.skip("업체014 표본 파일이 없음")

    workbook = openpyxl.load_workbook(source_path, data_only=True)
    assert "계좌 리스트" in workbook.sheetnames  # confirms the false-positive sheet is really in this file

    # No source_sheet= override — the auto-pick itself must land on the real
    # payroll sheet now, not just work when told which one to use.
    result = process("업체014", source_path, reference_path)
    assert result is not None

    mapped_headers = {m.source_header.replace("\n", "") for m in result.mapping_rows if m.kind == "자동매칭"}
    for expected in ("국민연금(본인부담)", "건강보험(본인부담)", "고용보험(본인부담)", "장기요양보험(본인부담)"):
        assert expected in mapped_headers, (expected, mapped_headers)


def test_업체014_근로금액과_소민세_인식() -> None:
    """"인 건 비 > 근로금액"(기본급 동의어 누락)과 "소민세"(소득세 오타)가
    canon.py에 없어서 통째로 미분류됐던 실사 케이스 — 두 항목 다 금액이 커서
    (기본급은 급여의 대부분, 소득세도 몇만 원 단위) 이것만으로 전 직원의
    지급액계/공제액계가 어긋났었다. 국민연금/건강보험/장기요양/고용보험
    자체는 이 파일에서도 원래부터 정상 인식됐음(canon 사전에 이미 등록돼
    있던 "(4.75%)" 식 요율 표기, strip_rate_suffix로 처리됨) — 사용자가
    본 "인식이 안 되는 문제"의 실체는 근로금액/소민세였다."""
    warnings.filterwarnings("ignore")
    source_path = SOURCE_DIR / "업체014 5월귀속 6월지급 급여대장(급+상).xlsx"
    reference_path = WEHAGO_DIR / "업체014-202605.xlsx"
    if not source_path.exists() or not reference_path.exists():
        pytest.skip("업체014 표본 파일이 없음")

    result = process("업체014", source_path, reference_path)
    assert result is not None

    mapped_headers = {m.source_header for m in result.mapping_rows if m.kind == "자동매칭"}
    assert "인 건 비 > 근로금액" in mapped_headers
    assert "소민세" in mapped_headers

    # Every matched employee's 지급액계/공제액계 now agrees with the source's
    # own stated totals exactly (both items carry real money for everyone).
    assert result.summary.matched > 0
    assert all(row.status == "정상" for row in result.verification_rows if row.status not in ("삭제됨",))


def test_업체056_연번과_종사자명_라벨_인식() -> None:
    """"No"/"순번" 대신 "연번", "성명"/"이름" 대신 "종사자명"을 쓰는 원본이라
    라벨 세트에 없어서 아예 구조 탐지 자체가 실패하고 있었다(오탐이 아니라
    후보가 0개). 같은 폴더의 "(급여)" 파일은 별개 문제(은행 계좌이체 목록일
    뿐 실제 급여 항목이 하나도 없음)라 여전히 인식 안 되는 게 맞다 — 위하고
    참고파일이 파일/에 없어서 process() 전체는 못 돌려보고 구조 탐지만
    검증."""
    warnings.filterwarnings("ignore")
    surcharge_path = SOURCE_DIR / "업체056 5월귀속 6월지급 급여대장(추급).xlsx"
    payroll_only_path = SOURCE_DIR / "업체056 5월귀속 6월지급 급여대장(급여).xlsx"
    if not surcharge_path.exists() or not payroll_only_path.exists():
        pytest.skip("업체056 표본 파일이 없음")

    workbook = openpyxl.load_workbook(surcharge_path, data_only=True)
    layout = detect_source_layout(workbook["2024-8"])
    assert layout is not None
    assert layout.name_col == 2  # "종사자명"
    assert layout.marker_col == 1  # "연번"

    # The bank-transfer-only file has no itemized payroll columns at all —
    # correctly produces no source candidate, not something to "recognize"
    # harder (see core.engine._has_any_amount_column).
    from core.engine import _source_candidates  # noqa: PLC0415

    assert _source_candidates(payroll_only_path) == []


def test_업체056_급여_열이_두_행에_걸친_그룹헤더_때문에_비어보임() -> None:
    """"연번"/"종사자명"을 고친 뒤에도 F열("급여", 실제 기본급 금액)이 여전히
    안 잡혔다 — 이 원본은 "급여"/"급여총계" 같은 금액 열 라벨이 마커+이름
    행(4행) 위의 별도 행(3행)에 세로 병합으로 걸쳐 있는데, 헤더 탐지가 앵커
    행(4행)만 헤더로 잡아서 3행의 라벨을 통째로 놓치고 있었다(F열의 4행
    쪽은 완전히 빈칸). 게다가 "급여"라는 텍스트 자체가 보통은 "급여 > 기본급"
    처럼 상위 그룹 라벨일 뿐이라 core/sheet_layout.py의 SrcItem 포팅이 일부러
    건너뛰게 되어 있어서, 병합 때문에 값을 억지로 읽어와도 그 열의 유일한
    라벨이 "급여"뿐이면 그마저 버려지고 있었다. 두 가지 다 고쳤다: (1) 앵커
    행 위에 "이미 알려진 헤더 행들이 비어있는 열을 채워주는" 행이 있으면
    헤더 블록을 위로 확장하고, (2) 더 구체적인 라벨이 전혀 없을 때는 "급여"
    같은 그룹 단어라도 마지막 수단으로 채택한다."""
    warnings.filterwarnings("ignore")
    source_path = SOURCE_DIR / "업체056 5월귀속 6월지급 급여대장(추급).xlsx"
    if not source_path.exists():
        pytest.skip("업체056 표본 파일이 없음")

    from core.classify import canon_of_src, group_of_src  # noqa: PLC0415
    from core.sheet_layout import source_item  # noqa: PLC0415

    workbook = openpyxl.load_workbook(source_path, data_only=True)
    worksheet = workbook["2024-8"]
    layout = detect_source_layout(worksheet)
    assert layout is not None
    # "급여"/"급여총계" (row 3) is vertically merged into row 4, so
    # merge-aware cell_text already resolves it without needing row 3 in
    # header_rows itself — the header-block-extension logic (still exercised
    # by other files, e.g. non-merged group headers) just isn't the
    # mechanism that fixes this particular file.

    salary_col = 6  # "급여" — 사원020/사원021 모두 2,156,880원, 실제 기본급 금액
    assert source_item(worksheet, layout, salary_col) == "급여"
    assert group_of_src(worksheet, layout, salary_col) == "수당"
    assert canon_of_src(worksheet, layout, salary_col) == "기본급"


def test_업체006_선임수당_공단_시트가_아니라_진짜_급여대장이_뽑힘() -> None:
    """이 워크북엔 진짜 급여대장("급여대장") 말고 "선임수당-공단"이라는
    시트도 있는데, 선임요양보호사 지원금 하나만 계산하는 부수적인 리포트다
    (건강보험/장기요양/고용보험 등 항목은 있지만 연장수당/야간수당/휴일수당/
    연차수당 등 실제 급여 항목은 전혀 없음). 그런데 이 시트가 빈 서식 행까지
    포함해서 실제 데이터 행 수가 진짜 급여대장(67행)보다 많아(94행) 행 수
    기준 랭킹에서 이겨버렸다 — "연장수당이 화면엔 있는데 인식이 안 된다"는
    증상은 사실 애초에 연장수당이 아예 없는 엉뚱한 시트를 대상으로 돌고
    있었기 때문이었다. 행 수 대신 서로 다른 급여 항목 canon 개수(다양성)를
    우선 기준으로 삼도록 후보 정렬을 고쳤다.

    같은 파일에서 진짜 급여대장의 "지급액"(총 지급액, 공제 시작 직전에 있는
    소계) 컬럼이 canon.py에서 "차인지급액"(실수령액)의 동의어로 등록돼 있어
    (VBA 원본 그대로 포팅한 값 — 대개는 "지급액"이 실수령액을 뜻하는 파일이
    많음) 실제 실수령액 칸인 "실지급액"보다 앞서 있다 보니 검증 시 실수령액
    자리에 총 지급액이 잘못 들어가고 있었다 — "지급액"/"지급액계" 둘 다 못
    찾던 상태였음. 여러 열이 "차인지급액"으로 인식되면 "실지급액"처럼 더
    명확한 표현을 우선하도록, "지급액"만 있고 더 나은 후보가 없을 때만
    지급액계의 최후 수단으로도 쓰도록 고쳤다."""
    warnings.filterwarnings("ignore")
    source_path = SOURCE_DIR / "업체006 5월귀속 6월지급 급여대장.xlsx"
    if not source_path.exists():
        pytest.skip("업체006 원본 표본 파일이 없음")

    from core.engine import _source_candidates  # noqa: PLC0415

    candidates = _source_candidates(source_path)
    assert candidates
    assert candidates[0][0].title == "급여대장"

    workbook = openpyxl.load_workbook(source_path, data_only=True)
    worksheet = workbook["급여대장"]
    layout = detect_source_layout(worksheet)
    assert layout is not None

    from core.reconcile import _find_src_canon_col, _find_src_pay_total_col  # noqa: PLC0415
    from core.sheet_layout import source_item  # noqa: PLC0415

    assert source_item(worksheet, layout, _find_src_pay_total_col(worksheet, layout)) == "지급액"
    assert source_item(worksheet, layout, _find_src_canon_col(worksheet, layout, "차인지급액")) == "실지급액"


def test_업체016_3행짜리_직원블록을_1행으로_펼침() -> None:
    """이 원본은 직원 1명이 물리적으로 3행을 차지하는데(연번/성명이 3행에
    걸쳐 세로 병합됨), 업체043의 연속행과는 근본적으로 다르다 —
    업체034은 연속행의 금액 칸이 대부분 비어있거나 텍스트라 합산해도 안전했지만,
    이 파일은 같은 열이 3개 행에 걸쳐 서로 다른 실제 숫자를 담고 있다(예:
    M열이 블록의 1행째엔 기본급, 2행째엔 주휴수당, 3행째엔 사대보험 환급).
    처음엔 블록 높이를 감지하고 그냥 합산했더니 지급액계 등 총계 칸이
    우연히 블록 첫 행에만 있어서 검증은 "정상"으로 통과했지만, 실제로는
    기본급이 0원, 요양수당 자리에 근무시간(54)이 들어가는 등 항목별 금액이
    전부 뒤섞여 있었다 — 세금/급여 데이터에서 가장 위험한 유형의 버그(겉보기엔
    멀쩡한데 숫자가 조용히 틀림)라, 한동안은 이런 구조를 아예 처리 거절하도록
    막아뒀었다.

    하지만 이 파일은 3개의 헤더 행(5/6/7행)이 블록의 3개 오프셋(8/9/10행,
    11/12/13행, ...)과 정확히 1:1로 대응한다는 걸 실사로 확인했다
    (`SourceLayout.header_rows`가 딱 `block_height`개일 때만) — 그래서 물리
    컬럼 하나를 오프셋 수만큼의 논리 컬럼으로 "펼쳐서"(`flatten_source_layout`)
    각 논리 컬럼이 자기 오프셋 행의 값만 읽도록 고쳤다. 인적사항 그룹 헤더
    행(4행, "인 적 사 항"/"기 본 사 항"/...)은 병합 때문에 헤더 블록 상단
    추출 로직에 딸려 들어와 `header_rows`가 4개가 되는데, 오프셋과 무관한
    상위 그룹 라벨일 뿐이라 뒤쪽 `block_height`개만 쓰면 된다. 연번/성명처럼
    3행에 걸쳐 병합된 라벨은 병합 인식 없는 원시 셀 값으로 읽어야
    중복 논리 컬럼이 생기지 않는다(병합 셀은 앵커가 아닌 칸이 진짜로
    비어 있으므로)."""
    warnings.filterwarnings("ignore")
    source_path = SOURCE_DIR / "업체016 5월귀속 6월지급 급여대장.xlsx"
    reference_path = DOWNLOADS_DIR / "업체016 (2026)-202605.xlsx"
    if not source_path.exists():
        pytest.skip("업체016 원본 표본 파일이 없음")
    if not reference_path.exists():
        pytest.skip("업체016 위하고 참고파일이 없음")

    workbook = openpyxl.load_workbook(source_path, data_only=True)
    worksheet = workbook["급여대장"]
    layout = detect_source_layout(worksheet)
    assert layout is not None
    assert layout.needs_flatten is True
    assert layout.block_height == 3

    from core.sheet_layout import flatten_source_layout  # noqa: PLC0415

    grid, flat_layout = flatten_source_layout(worksheet, layout)
    assert flat_layout.needs_flatten is False
    assert flat_layout.block_height == 1
    header = grid._rows[0]
    사원022 = dict(zip(header, grid._rows[1], strict=True))
    assert 사원022["성명"] == "사원022"
    assert 사원022["기본급"] == 601800
    assert 사원022["주휴수당"] == 125400
    assert 사원022["사대보험 환급"] is None

    from core.engine import _source_candidates  # noqa: PLC0415

    candidates = _source_candidates(source_path)
    assert candidates
    assert candidates[0][0].title == "급여대장"

    result = process("업체016", source_path, reference_path)
    assert result is not None
    assert result.summary.matched > 0
