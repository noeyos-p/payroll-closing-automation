from pathlib import Path

from core.engine import list_dest_employee_names, list_source_employee_names

FIXTURES = Path(__file__).parent / "fixtures" / "golden"


def test_list_source_employee_names_reads_sample1():
    names = list_source_employee_names(FIXTURES / "sample1-source.xlsx")
    assert len(names) > 0
    assert all(isinstance(n, str) and n.strip() for n in names)


def test_list_dest_employee_names_reads_sample1():
    names = list_dest_employee_names(FIXTURES / "sample1-wehago-before.xlsx")
    assert len(names) > 0
    assert all(isinstance(n, str) and n.strip() for n in names)


def test_source_and_dest_names_overlap_substantially_for_same_client():
    # 같은 거래처의 원본/위하고 파일이면 이름이 거의 다 겹쳐야 한다 —
    # 자동 인식(core/company_match.py)이 실제로 동작할 수 있는 전제조건 확인.
    source_names = set(list_source_employee_names(FIXTURES / "sample1-source.xlsx"))
    dest_names = set(list_dest_employee_names(FIXTURES / "sample1-wehago-before.xlsx"))
    overlap = source_names & dest_names
    assert len(overlap) / max(len(source_names), 1) > 0.5
