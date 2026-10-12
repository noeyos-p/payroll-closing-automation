from core.company_match import CompanyRoster, match_company


def test_matches_clear_winner():
    ledger_names = ["김철수", "박영희", "최민수"]
    rosters = [
        CompanyRoster(1, "테스트요양원", {"김철수", "박영희", "최민수"}),
        CompanyRoster(2, "다른요양원", {"최아무개"}),
    ]
    result = match_company(ledger_names, rosters)
    assert result.status == "matched"
    assert result.company_id == 1
    assert result.company_name == "테스트요양원"


def test_no_match_when_ratio_too_low():
    ledger_names = ["김철수", "박영희", "최민수"]
    rosters = [CompanyRoster(1, "테스트요양원", {"최아무개", "정아무개"})]
    result = match_company(ledger_names, rosters)
    assert result.status == "no_match"
    assert result.company_id is None


def test_ambiguous_when_two_companies_score_similarly():
    ledger_names = ["김철수", "박영희", "최민수", "정수진"]
    rosters = [
        CompanyRoster(1, "요양원A", {"김철수", "박영희", "최민수"}),
        CompanyRoster(2, "요양원B", {"김철수", "박영희", "정수진"}),
    ]
    result = match_company(ledger_names, rosters)
    assert result.status == "ambiguous"
    assert result.company_id is None
    assert len(result.candidates) == 2


def test_no_rosters_registered_yields_no_match():
    result = match_company(["김철수"], [])
    assert result.status == "no_match"


def test_candidates_list_sorted_by_ratio_descending():
    ledger_names = ["김철수", "박영희", "최민수", "정수진"]
    rosters = [
        CompanyRoster(1, "낮은일치", {"김철수"}),
        CompanyRoster(2, "높은일치", {"김철수", "박영희", "최민수", "정수진"}),
    ]
    result = match_company(ledger_names, rosters)
    assert result.status == "matched"
    assert result.company_name == "높은일치"
    assert result.candidates[0].company_name == "높은일치"
    assert result.candidates[0].ratio == 1.0
