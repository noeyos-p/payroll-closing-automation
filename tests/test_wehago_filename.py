from pathlib import Path

from core.wehago_filename import parse_upload_filename


def test_parses_company_month_and_paytype():
    info = parse_upload_filename(Path("260921_업체00708(급여).xlsx"))
    assert info is not None
    assert info.company == "업체007"
    assert info.month == 8
    assert info.pay_type == "급여"


def test_company_name_with_leading_digits_is_not_confused_with_month():
    # 실사 확인 사례: 파일명에 회사명 자체가 숫자로 시작해도, 정규식이
    # "마지막 2자리"를 귀속월로 고정해서 뽑기 때문에 문제없다.
    info = parse_upload_filename(Path("260827_업체07107(추급).xlsx"))
    assert info is not None
    assert info.company == "업체071"
    assert info.month == 7
    assert info.pay_type == "추급"


def test_pay_type_is_none_when_only_one_pay_run_that_month():
    info = parse_upload_filename(Path("260921_업체00708.xlsx"))
    assert info is not None
    assert info.pay_type is None


def test_unrecognized_filename_returns_none():
    assert parse_upload_filename(Path("아무개파일.xlsx")) is None
    assert parse_upload_filename(Path("260921_회사명.xlsx")) is None  # 귀속월 2자리 없음
