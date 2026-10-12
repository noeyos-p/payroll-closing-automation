from pathlib import Path

from core.registry import delete_company, get_company, get_company_by_name, list_companies, register_company

FIXTURES = Path(__file__).parent / "fixtures" / "golden"


def _paths(tmp_path: Path) -> tuple[Path, Path]:
    return tmp_path / "registry.db", tmp_path / "reference_files"


def test_register_company_creates_entry_with_employee_roster(tmp_path):
    db_path, ref_dir = _paths(tmp_path)
    company = register_company(
        "업체075", FIXTURES / "sample1-wehago-before.xlsx", db_path=db_path, reference_dir=ref_dir
    )
    assert company.name == "업체075"
    assert company.employee_count > 0
    assert Path(company.reference_file_path).exists()
    # 원본 파일이 지워져도 안전하도록 앱 데이터 폴더로 복사된 사본을 가리켜야 함
    assert Path(company.reference_file_path) != FIXTURES / "sample1-wehago-before.xlsx"


def test_register_company_is_upsert_by_name(tmp_path):
    db_path, ref_dir = _paths(tmp_path)
    first = register_company("업체076", FIXTURES / "sample1-wehago-before.xlsx", db_path=db_path, reference_dir=ref_dir)
    second = register_company(
        "업체076", FIXTURES / "sample2-wehago-before.xlsx", db_path=db_path, reference_dir=ref_dir
    )
    assert first.id == second.id
    assert second.reference_file_name == "sample2-wehago-before.xlsx"
    assert len(list_companies(db_path=db_path)) == 1


def test_list_companies_returns_all_registered(tmp_path):
    db_path, ref_dir = _paths(tmp_path)
    register_company("업체075", FIXTURES / "sample1-wehago-before.xlsx", db_path=db_path, reference_dir=ref_dir)
    register_company("업체076", FIXTURES / "sample2-wehago-before.xlsx", db_path=db_path, reference_dir=ref_dir)
    companies = list_companies(db_path=db_path)
    assert {c.name for c in companies} == {"업체075", "업체076"}


def test_get_company_by_name(tmp_path):
    db_path, ref_dir = _paths(tmp_path)
    register_company("업체075", FIXTURES / "sample1-wehago-before.xlsx", db_path=db_path, reference_dir=ref_dir)
    found = get_company_by_name("업체075", db_path=db_path)
    assert found is not None
    missing = get_company_by_name("존재하지않음", db_path=db_path)
    assert missing is None


def test_delete_company(tmp_path):
    db_path, ref_dir = _paths(tmp_path)
    company = register_company(
        "업체075", FIXTURES / "sample1-wehago-before.xlsx", db_path=db_path, reference_dir=ref_dir
    )
    delete_company(company.id, db_path=db_path)
    assert get_company(company.id, db_path=db_path) is None
    assert list_companies(db_path=db_path) == []
