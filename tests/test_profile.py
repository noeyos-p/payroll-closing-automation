from core.profile import derive_client_label, guess_company_name_from_filename, payroll_type_from_sheet_name


def test_derive_client_label_appends_accrual_month():
    assert derive_client_label("업체022 7월귀속 8월지급 급여대장") == "업체02207"


def test_derive_client_label_falls_back_to_first_token():
    assert derive_client_label("아무개파일") == "아무개파일"


def test_guess_company_name_strips_month_suffix():
    assert guess_company_name_from_filename("홍길동업체 7월귀속 8월지급 급여대장") == "홍길동업체"


def test_guess_company_name_ignores_trailing_suffix_after_accrual():
    # 실사 확인: 업체077 7월귀속 8월지급 급여대장(급여)(급+상).xlsx
    assert guess_company_name_from_filename("업체077 7월귀속 8월지급 급여대장(급여)(급+상)") == "업체077"


def test_guess_company_name_falls_back_to_first_token():
    assert guess_company_name_from_filename("아무개파일") == "아무개파일"


def test_payroll_type_from_sheet_name_exact_and_numbered():
    assert payroll_type_from_sheet_name("급여") == "급여"
    assert payroll_type_from_sheet_name("추급1") == "추급"
    assert payroll_type_from_sheet_name("추급2") == "추급"
    assert payroll_type_from_sheet_name("급상") == "급+상"
    assert payroll_type_from_sheet_name("급+상") == "급+상"


def test_payroll_type_from_sheet_name_ignores_unrelated_names_containing_token():
    # 실사 확인: 업체051 8월귀속 9월지급 급여대장.xlsx — 서로
    # 다른 두 후보 시트("급여계산"/"급여대장")가 종류 구분과 무관하게
    # "급여"라는 글자를 포함할 뿐인데, 예전엔 부분문자열 포함으로 판정해서
    # "종류가 다른 시트 2개"로 오인해 잘못 합쳐졌다(항목매핑결과가 두
    # 배로 뜨고 수동 매칭 적용도 안 먹히는 버그로 이어짐).
    assert payroll_type_from_sheet_name("급여계산") is None
    assert payroll_type_from_sheet_name("급여대장") is None
