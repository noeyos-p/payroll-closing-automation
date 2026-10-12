from pathlib import Path

from core import registry
from core.batch import process_batch, process_one, process_with_reference

FIXTURES = Path(__file__).parent / "fixtures" / "golden"


def _paths(tmp_path: Path) -> tuple[Path, Path]:
    return tmp_path / "registry.db", tmp_path / "reference_files"


def test_process_one_needs_reference_when_no_company_registered(tmp_path):
    db_path, _ = _paths(tmp_path)
    item = process_one(FIXTURES / "sample1-source.xlsx", db_path=db_path)
    assert item.status == "needs_reference"
    assert item.result is None


def test_process_with_reference_succeeds_and_auto_registers(tmp_path):
    db_path, ref_dir = _paths(tmp_path)
    item = process_with_reference(
        FIXTURES / "sample1-source.xlsx",
        "업체075",
        FIXTURES / "sample1-wehago-before.xlsx",
        db_path=db_path,
        reference_dir=ref_dir,
    )
    assert item.status == "matched"
    assert item.result is not None
    assert item.company_id is not None

    registered = registry.list_companies(db_path=db_path)
    assert len(registered) == 1
    assert registered[0].name == "업체075"


def test_process_one_auto_detects_after_registration(tmp_path):
    db_path, ref_dir = _paths(tmp_path)
    process_with_reference(
        FIXTURES / "sample1-source.xlsx",
        "업체075",
        FIXTURES / "sample1-wehago-before.xlsx",
        db_path=db_path,
        reference_dir=ref_dir,
    )
    # 다음부터는 급여대장 원본만으로 자동 인식되어야 한다.
    item = process_one(FIXTURES / "sample1-source.xlsx", db_path=db_path)
    assert item.status == "matched"
    assert item.company_name == "업체075"
    assert item.result is not None


def test_batch_one_unregistered_file_does_not_block_others(tmp_path):
    db_path, ref_dir = _paths(tmp_path)
    process_with_reference(
        FIXTURES / "sample1-source.xlsx",
        "업체075",
        FIXTURES / "sample1-wehago-before.xlsx",
        db_path=db_path,
        reference_dir=ref_dir,
    )
    results = process_batch(
        [FIXTURES / "sample1-source.xlsx", FIXTURES / "sample2-source.xlsx"], db_path=db_path
    )
    statuses = {r.path.name: r.status for r in results}
    assert statuses["sample1-source.xlsx"] == "matched"
    assert statuses["sample2-source.xlsx"] == "needs_reference"


def test_process_one_recognizes_company_from_filename(tmp_path):
    """실사 피드백: 파일명이 "{거래처명} N월귀속 M월지급 급여대장" 관례를
    따르면(사용자의 실제 파일 전부가 그럼) 사원 명부가 안 겹쳐도 파일명만으로
    인식돼야 한다."""
    db_path, ref_dir = _paths(tmp_path)
    registry.register_company(
        "업체075", FIXTURES / "sample1-wehago-before.xlsx", db_path=db_path, reference_dir=ref_dir
    )

    renamed = tmp_path / "업체075 6월귀속 7월지급 급여대장.xlsx"
    renamed.write_bytes((FIXTURES / "sample1-source.xlsx").read_bytes())

    item = process_one(renamed, db_path=db_path)
    assert item.status == "matched"
    assert item.company_name == "업체075"


def test_filename_match_wins_even_when_employee_roster_would_not_match(tmp_path):
    """파일명 매칭이 사원 명부 매칭보다 우선이어야 한다 — 실제 급여대장의
    사원 구성이(이직/신규입사 등으로) 등록 당시 위하고 명부와 많이 달라져도,
    파일명에 회사 이름이 정확히 적혀 있으면 그걸로 인식돼야 한다."""
    db_path, ref_dir = _paths(tmp_path)
    registry.register_company(
        "업체076", FIXTURES / "sample2-wehago-before.xlsx", db_path=db_path, reference_dir=ref_dir
    )

    # 내용은 완전히 다른 회사(sample1)의 사원 명부지만, 파일명은 "업체076"로 되어 있음.
    renamed = tmp_path / "업체076 6월귀속 7월지급 급여대장.xlsx"
    renamed.write_bytes((FIXTURES / "sample1-source.xlsx").read_bytes())

    item = process_one(renamed, db_path=db_path)
    assert item.status == "matched"
    assert item.company_name == "업체076"


def test_process_with_reference_error_when_reference_missing(tmp_path):
    db_path, ref_dir = _paths(tmp_path)
    item = process_with_reference(
        FIXTURES / "sample1-source.xlsx",
        "존재안함",
        tmp_path / "없는파일.xlsx",
        db_path=db_path,
        reference_dir=ref_dir,
    )
    assert item.status == "error"
    assert item.error
