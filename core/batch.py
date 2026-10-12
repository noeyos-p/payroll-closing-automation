"""여러 급여대장을 한 번에 처리하는 배치 로직 — Qt와 무관한 순수 함수/데이터라
GUI 없이도 테스트할 수 있다(ui/batch_worker.py가 QThread에서 이 함수들을
호출한다).

파일마다 사업장을 두 단계로 인식한다:
1. **파일명 우선** — "{거래처명} N월귀속 M월지급 급여대장..." 관례(사용자의
   실제 파일 전부가 이 형식을 따름)에서 거래처명만 뽑아(core/profile.py의
   `guess_company_name_from_filename`) 등록된 사업장명과 정확히 같은 게 있으면
   그걸로 확정한다. 사원 명부를 아예 안 봐도 되는 가장 직접적인 신호라 — 실사
   피드백: 사원 명부 겹침만 보던 이전 버전은 신규입사/이름 파싱 차이로 인식을
   놓치는 경우가 있었는데, 파일명에 회사명이 그대로 있는데도 못 찾는 게 부자연
   스럽다는 지적을 받고 추가함.
2. **사원 명부 겹침(폴백)** — 파일명이 이 관례를 안 따르거나(등록된 이름과 정확히
   안 맞음) 등록된 사업장이 아니면, core/company_match.py로 사원 이름 겹침을 본다.

둘 다 인식 안 되면(no_match/ambiguous) 그 자리에서 처리하지 않고 "확인 필요"
상태로만 남긴다 — 배치 전체를 막지 않고, 사용자가 그 파일에 한해 위하고 파일을
직접 골라 `process_with_reference`를 다시 호출하면 된다(한 파일이 애매해도
나머지는 정상 처리된 채로 유지됨).

위하고 참고파일을 사람이 직접 골라서 처리에 성공하면(`process_with_reference`)
그 사업장명으로 조용히 등록/갱신한다(core/registry.py) — 다음부터는 원본 파일만
올려도 자동 인식된다(사용자 지시 — 명시적 등록 버튼 없이 자동 저장)."""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from core import registry
from core.company_match import CompanyCandidate, CompanyRoster, match_company
from core.engine import list_source_employee_names, list_source_sheet_candidates, process
from core.profile import (
    derive_client_label,
    guess_company_name_from_filename,
    parse_template_token,
    payroll_type_from_filename,
    payroll_type_from_sheet_name,
    payroll_types_in_filename,
    template_token,
)
from core.reconcile import ReconcileResult, merge_reconcile_results


@dataclass
class BatchItem:
    path: Path
    status: str  # "matched" | "needs_reference" | "error"
    company_name: str | None = None
    company_id: int | None = None
    match_candidates: list[CompanyCandidate] = field(default_factory=list)
    result: ReconcileResult | None = None
    error: str | None = None
    source_sheet: str | None = None
    reference_path: Path | None = None
    reference_sheet: str | None = None
    # True면 이 항목이 `plan_batch_add`가 추급/급여/급+상 시트(또는 파일)별로
    # 쪼갠 결과 중 하나라는 뜻 — `ui/main_window.py`의 라벨링이 이 경우
    # 파일명보다 시트 이름을 우선하도록 판단 근거로 쓴다(파일명 자체가
    # "...(급+상)(추급)"처럼 두 종류를 다 담고 있을 수 있어서, 이땐 파일명만
    # 으로는 이 항목이 정확히 어느 쪽인지 구분이 안 된다 — 사용자 실사,
    # 2026-09-08).
    was_split: bool = False
    # 같은 종류(추급1/추급2처럼)로 나뉜 원본을 여러 개 합쳐서 이 항목
    # 하나로 처리했으면 그 (파일, 시트) 조합 전부(2개 이상) — 한 워크북
    # 안에서 시트로만 나뉜 경우("추급1"/"추급2" 두 시트, 같은 path)도,
    # 아예 별도 파일로 나뉜 경우(다른 path)도 똑같이 이걸로 표현한다.
    # 재실행(수동 매칭 적용/위하고 파일 다시 지정) 때 같은 묶음으로 다시
    # 합치기 위해 들고 있는다. 조합이 하나뿐이면 None(merge_reconcile_results를
    # 안 탐 — `source_sheet` 하나만 쓰는 기존 경로와 동일).
    merge_sources: list[tuple[Path, str | None]] | None = None


def _template_for(
    company_id: int | None,
    path: Path,
    source_sheet: str | None,
    merge_sources: list[tuple[Path, str | None]] | None,
    db_path: Path,
) -> tuple[object, Path] | None:
    """이 항목에 쓸 "원본 -> 수정본 양식" 변환 규칙(core/source_template.py).
    종류(급여/급+상/추급)는 큐 분리 표시(`<양식:…>`) -> 파일명 괄호 하나 순으로
    정하고, 그 종류의 규칙이 저장돼 있으면 쓴다. 종류 표시가 없는 파일은
    종류 구분 없이 저장된 규칙("")만 쓴다. 사용자가 시트를 직접 고른 경우나
    여러 원본을 합치는 항목에는 쓰지 않는다."""
    if company_id is None or merge_sources:
        return None
    token = parse_template_token(source_sheet)
    if token is not None:
        ptype = token
    elif source_sheet:
        return None
    else:
        types = payroll_types_in_filename(path.stem)
        if len(types) > 1:
            return None
        ptype = types[0] if types else ""
    try:
        templates = registry.get_source_templates(company_id, db_path=db_path)
    except Exception:
        return None
    if ptype not in templates:
        return None
    from core.source_template import Recipe

    recipe_json, template_path = templates[ptype]
    try:
        recipe = Recipe.from_json(recipe_json)
    except Exception:
        return None
    return (recipe, template_path) if recipe.complete else None


def _rosters_from_companies(companies: list[registry.Company]) -> list[CompanyRoster]:
    return [CompanyRoster(c.id, c.name, c.base_names) for c in companies]


def _run(
    path: Path,
    client_name: str,
    reference_path: Path,
    *,
    company_id: int | None = None,
    source_sheet: str | None = None,
    reference_sheet: str | None = None,
    manual_overrides: dict[int, str | None] | None = None,
    learned_mappings: dict[str, str] | None = None,
    was_split: bool = False,
    merge_sources: list[tuple[Path, str | None]] | None = None,
    source_template: tuple[object, Path] | None = None,
) -> BatchItem:
    # 큐 분리 표시(`<양식:급+상>`)는 실제 시트가 아니다 — 처리할 때는 시트를
    # 고정하지 않는다(변환 규칙이 자기 시트를 스스로 찾는다).
    engine_sheet = None if parse_template_token(source_sheet) is not None else source_sheet
    try:
        if merge_sources and len(merge_sources) > 1:
            # 추급1/추급2처럼 같은 종류로 나뉜 원본 여러 개(같은 파일 안
            # 시트로 나뉘었든, 아예 별도 파일로 나뉘었든) — 하나씩 따로
            # reconcile()한 뒤 하나로 합친다(core/reconcile.py의
            # merge_reconcile_results, 사용자 요청 2026-09-08). 하나가
            # 구조를 못 읽어도(None) 나머지로는 계속 진행 — 완전히 실패한
            # 것만 조용히 빠진다.
            sub_results = [
                r for (p, sheet) in merge_sources
                if (r := process(
                    client_name, p, reference_path,
                    source_sheet=sheet, reference_sheet=reference_sheet,
                    manual_overrides=manual_overrides, learned_mappings=learned_mappings,
                )) is not None
            ]
            result = merge_reconcile_results(sub_results) if sub_results else None
        else:
            result = process(
                client_name, path, reference_path,
                source_sheet=engine_sheet, reference_sheet=reference_sheet,
                manual_overrides=manual_overrides, learned_mappings=learned_mappings,
                source_template=source_template,
            )
            if result is not None and source_template is not None and not result.notes:
                from core import source_template as _st

                why = _st.LAST_APPLY_FAILURE
                result.notes.append(
                    "저장된 수정본 양식으로 변환하지 못해 원본을 그대로 읽어 처리했습니다"
                    + (f"({why})" if why else "")
                    + ". 결과를 꼭 확인하세요."
                )
    except Exception as exc:
        return BatchItem(path, "error", company_name=client_name, error=str(exc))
    if result is None:
        return BatchItem(
            path, "error", company_name=client_name,
            error="원본 또는 위하고 파일에서 급여 구조를 인식하지 못했습니다.",
        )
    # 합쳐진 항목은 개별 sub-call마다 다른 시트/파일을 썼으니 대표로 첫
    # 것을 남겨둔다 — `ui/main_window.py`의 라벨링(_type_tagged_label)이
    # `item.source_sheet`로 종류(추급/급여/급+상)를 판단하는데, "추급1"도
    # `payroll_type_from_sheet_name`이 "추급"으로 정규화해주므로 그대로
    # 대표값으로 써도 정확하다(별도 파일로 나뉜 경우엔 파일명 쪽에서
    # 종류를 뽑으므로 이 값은 안 쓰인다).
    display_sheet = source_sheet or (merge_sources[0][1] if merge_sources else None)
    return BatchItem(
        path, "matched", company_name=client_name, company_id=company_id, result=result,
        source_sheet=display_sheet, reference_path=reference_path, reference_sheet=reference_sheet,
        was_split=was_split, merge_sources=merge_sources,
    )


def sheets_for_batch_add(path: Path) -> list[list[str] | None]:
    """파일을 큐에 추가할 때 몇 개의 항목으로 나눌지, 그리고 항목마다 어느
    시트(들)를 처리할지 결정한다. 한 워크북 안에 추급/급여/급+상 중 서로
    다른 두 종류 이상의 시트가 있으면(사용자 실사 확인, 2026-09-08 —
    8월지급 급여를 이렇게 세 시트로 나눠서 한 파일에 넣어두는 관례가 있음)
    **종류별로** 묶어서 그 그룹들을 돌려준다 — 같은 종류가 시트 여러 개로
    또 나뉜 경우(예: "추급1"/"추급2", 다른 종류가 하나도 안 섞여 있어도)는
    한 그룹(한 큐 항목)으로 묶여서 나중에 합쳐진다(`core/reconcile.py`의
    `merge_reconcile_results`, 사용자 요청 2026-09-08: "동명이인에
    생년월일 같으면 합쳐서 다운로드"). 트리거 조건은 "종류가 붙은 시트가
    총 2개 이상"이다(종류가 서로 달라야 하는 게 아니라 — "추급1"/"추급2"
    딱 둘뿐이어도 갈라야 하므로). 그 외(종류 붙은 시트가 1개 이하)에는
    `[None]`을 돌려줘서 기존처럼 최선 후보 시트를 자동으로 쓰게 한다.
    워크북을 못 읽는 등의 이유로 후보 목록 자체를 못 구하면 `[None]`으로
    안전하게 폴백 — 실제 처리 단계에서 에러로 다시 표면화된다."""
    try:
        candidates = list_source_sheet_candidates(path)
    except Exception:
        return [None]
    tagged = [(c.name, payroll_type_from_sheet_name(c.name)) for c in candidates]
    typed = [(name, t) for name, t in tagged if t is not None]
    if len(typed) < 2:
        return [None]
    groups: dict[str, list[str]] = {}
    for name, t in typed:
        groups.setdefault(t, []).append(name)
    return list(groups.values())


def _cross_file_group_key(path: Path) -> tuple[str, str] | None:
    """"거래처명+귀속월" 라벨(derive_client_label)과 파일명 끝의 종류
    표시(payroll_type_from_filename)를 합친 키 — 이 둘이 같은 파일 여러
    개는 같은 사업장의 같은 달, 같은 종류(추급 등)가 아예 별도 파일로
    나뉜 것으로 본다(사용자 요청, 2026-09-08: "시트 말고도 엑셀파일이
    나눠져있는 경우도 있어"). 파일명에 종류 표시가 없으면(추급/급여/
    급+상 중 아무것도 안 붙어있으면) 이 그룹핑 대상이 아니다 — `None`."""
    type_token = payroll_type_from_filename(path.stem)
    if type_token is None:
        return None
    return (derive_client_label(path.stem), type_token)


def _template_split(path: Path, db_path: Path) -> list[str] | None:
    """파일명에 종류가 둘 이상("(급여)(급+상)")이고, 그 사업장에 그 종류들
    전부의 수정본 양식 변환 규칙이 있으면 종류 목록 — 큐에서 종류별 항목으로
    나눈다(사람이 실제로 그렇게 파일을 나눠 고쳐 저장해 둔 사업장: 업체012 등)."""
    types = payroll_types_in_filename(path.stem)
    if len(types) < 2:
        return None
    try:
        company = registry.get_company_by_name(guess_company_name_from_filename(path.stem), db_path=db_path)
        if company is None:
            return None
        templates = registry.get_source_templates(company.id, db_path=db_path)
    except Exception:
        return None
    if all(t in templates and templates[t][0].get("complete") for t in types):
        return types
    return None


def plan_batch_add(
    paths: list[Path], db_path: Path = registry.DEFAULT_DB_PATH
) -> list[list[tuple[Path, str | None]]]:
    """여러 파일을 큐에 한꺼번에 추가할 때, 결과적으로 큐 항목이 몇 개가
    되고 각 항목이 어느 (파일, 시트) 조합을 처리할지 계획한다. 두 가지
    분리/합치기 규칙을 함께 본다:
    1. 파일 하나 안에 추급/급여/급+상 시트가 여러 개 있으면(`sheets_for_batch_add`)
       종류별로 나누고, 같은 종류가 시트로 또 나뉘었으면(추급1/추급2) 그
       시트들을 합쳐서 한 항목으로 처리한다 — 이런 파일은 아래 2번(파일
       간 합치기) 대상에서는 제외한다(이미 자기 안에서 다 처리됨).
    2. 내부적으로 안 나뉜 파일들끼리는, "거래처명+귀속월+종류"가 같은
       파일이 2개 이상이면(예: "...추급1.xlsx"/"...추급2.xlsx"처럼 아예
       별도 파일로 나뉜 경우) 그 파일들을 합쳐서 한 항목으로 처리한다.
    반환값의 각 원소는 한 큐 항목이 처리할 (파일, 시트) 목록 — 보통
    1개(그냥 그 파일 하나, 시트는 자동 판단)지만 위 두 규칙에 걸리면
    2개 이상이 된다."""
    groups: list[list[tuple[Path, str | None]]] = []
    cross_file: dict[tuple[str, str], list[tuple[Path, str | None]]] = {}
    cross_file_order: list[tuple[str, str]] = []

    for path in paths:
        split_types = _template_split(path, db_path)
        if split_types:
            for t in split_types:
                groups.append([(path, template_token(t))])
            continue
        sheet_groups = sheets_for_batch_add(path)
        if len(sheet_groups) == 1 and sheet_groups[0] is None:
            # 이 파일 안에서는 안 나뉨 — 다른 파일과 합칠 후보로 본다.
            key = _cross_file_group_key(path)
            if key is None:
                groups.append([(path, None)])
                continue
            if key not in cross_file:
                cross_file_order.append(key)
                cross_file[key] = []
            cross_file[key].append((path, None))
        else:
            for sheet_group in sheet_groups:
                if sheet_group is None:
                    groups.append([(path, None)])
                else:
                    groups.append([(path, sheet) for sheet in sheet_group])

    # 짝이 있든(2개 이상) 없든(혼자뿐) 그대로 큐 항목 하나씩으로 낸다 —
    # 혼자뿐이면 `_run`이 merge_sources 길이 1은 그냥 일반 단일 처리로
    # 취급하므로(BatchItem 필드 주석 참고) 별도 분기가 필요 없다.
    for key in cross_file_order:
        groups.append(cross_file[key])
    return groups


def _registered_reference(company: registry.Company, path: Path, source_sheet: str | None,
                          merge_sources: list[tuple[Path, str | None]] | None, db_path: Path) -> Path:
    token = parse_template_token(source_sheet)
    if token is not None:
        ptype = token
    elif source_sheet:
        ptype = payroll_type_from_sheet_name(source_sheet)
    else:
        types = payroll_types_in_filename(path.stem)
        ptype = types[0] if len(types) == 1 else '' if not types else None
    if merge_sources:
        categories = {parse_template_token(s) or payroll_type_from_sheet_name(s or '') or payroll_type_from_filename(p.stem)
                      for p, s in merge_sources}
        ptype = next(iter(categories)) if len(categories) == 1 else None
    reference = registry.get_payroll_reference(company.id, ptype, db_path=db_path) if ptype is not None else None
    return reference or Path(company.reference_file_path)


def process_one(
    path: Path, db_path: Path = registry.DEFAULT_DB_PATH, source_sheet: str | None = None,
    was_split: bool = False, merge_sources: list[tuple[Path, str | None]] | None = None,
) -> BatchItem:
    """급여대장 파일 하나를 등록된 사업장과 대조해 자동 인식 + 처리한다.
    인식 안 되면 처리하지 않고 상태만 "needs_reference"로 돌려준다.
    `source_sheet`를 지정하면(한 워크북에 추급/급여/급+상 시트가 나뉘어
    있어 `plan_batch_add`가 이 파일을 여러 큐 항목으로 쪼갠 경우) 사업장
    인식용 사원 명부도 그 시트에서 읽고, 처리도 그 시트로 고정한다 —
    시트마다 사원 구성이 다를 수 있어(예: 추급 시트는 일부 인원만) 최선
    후보 시트만 보면 엉뚱한 명부로 사업장을 잘못 인식할 수 있다.
    `merge_sources`(2개 이상)를 지정하면 그 (파일, 시트) 조합들을 각각
    처리한 뒤 하나로 합친다(`_run` 참고) — 이땐 사업장 인식용 명부를
    `merge_sources[0]`에서 읽는다(같은 사업장의 같은 달이라 어느 걸 봐도
    충분하다는 전제로 이미 grouping 단계에서 걸러짐).
    `was_split=True`면 그 쪼갠 항목이라는 표시를 결과에도 그대로 남긴다
    (`BatchItem.was_split` 참고 — 라벨링 우선순위 판단용)."""
    companies = registry.list_companies(db_path=db_path)
    roster_path, roster_sheet = merge_sources[0] if merge_sources else (path, source_sheet)

    filename_guess = guess_company_name_from_filename(path.stem)
    filename_match = next((c for c in companies if c.name == filename_guess), None)
    if filename_match is not None:
        return _run(
            path, filename_match.name, _registered_reference(filename_match, path, source_sheet, merge_sources, db_path), company_id=filename_match.id,
            source_sheet=source_sheet, was_split=was_split, merge_sources=merge_sources,
            learned_mappings=registry.get_column_mappings(filename_match.id, db_path=db_path),
            source_template=_template_for(filename_match.id, path, source_sheet, merge_sources, db_path),
        )

    try:
        ledger_names = list_source_employee_names(roster_path, roster_sheet)
    except Exception as exc:
        return BatchItem(path, "error", error=f"급여대장을 읽지 못했습니다: {exc}")

    match = match_company(ledger_names, _rosters_from_companies(companies))
    if match.status != "matched":
        return BatchItem(path, "needs_reference", match_candidates=match.candidates)

    company = next(c for c in companies if c.id == match.company_id)
    return _run(
        path, company.name, _registered_reference(company, path, source_sheet, merge_sources, db_path), company_id=company.id,
        source_sheet=source_sheet, was_split=was_split, merge_sources=merge_sources,
        learned_mappings=registry.get_column_mappings(company.id, db_path=db_path),
        source_template=_template_for(company.id, path, source_sheet, merge_sources, db_path),
    )


def process_with_reference(
    path: Path,
    client_name: str,
    reference_path: Path,
    source_sheet: str | None = None,
    reference_sheet: str | None = None,
    manual_overrides: dict[int, str | None] | None = None,
    db_path: Path = registry.DEFAULT_DB_PATH,
    reference_dir: Path = registry.DEFAULT_REFERENCE_DIR,
    auto_register: bool = True,
    company_id: int | None = None,
    was_split: bool = False,
    merge_sources: list[tuple[Path, str | None]] | None = None,
) -> BatchItem:
    """자동 인식이 안 됐거나(needs_reference) 사용자가 직접 다른 위하고 파일을
    쓰고 싶을 때, 또는 검토표에서 항목을 수동으로 다시 매칭했을 때(manual_overrides)
    쓴다. `auto_register=True`(기본)면 성공 시 이 사업장명으로 자동 등록/갱신된다 —
    수동 매핑만 다시 적용하는 재실행에서는 매번 새로 복사/등록할 필요가 없으므로
    호출부(ui/main_window.py)가 이미 등록된 항목의 재실행에는 False를 넘긴다.
    `company_id`는 이미 사업장이 정해진 상태에서(시트 변경/수동 매칭 재실행)
    사업장 정체성과 이 사업장에 저장된 `learned_mappings`(core/registry.py의
    `column_mappings`)를 그대로 이어받기 위한 것 — `auto_register=False`라도
    이걸 넘기면 결과 `BatchItem.company_id`가 비지 않는다. `was_split`/
    `merge_sources`도 같은 이유로 재실행 때 이어받는다(BatchItem 참고)."""
    learned_mappings = registry.get_column_mappings(company_id, db_path=db_path) if company_id is not None else None
    item = _run(
        path, client_name, reference_path,
        source_sheet=source_sheet, reference_sheet=reference_sheet,
        manual_overrides=manual_overrides, learned_mappings=learned_mappings, was_split=was_split,
        merge_sources=merge_sources,
        source_template=_template_for(company_id, path, source_sheet, merge_sources, db_path),
    )
    if item.status == "matched":
        if auto_register:
            company = registry.register_company(
                client_name, reference_path, reference_sheet, db_path=db_path, reference_dir=reference_dir
            )
            item.company_id = company.id
        elif company_id is not None:
            item.company_id = company_id
    return item


def process_batch(paths: list[Path], db_path: Path = registry.DEFAULT_DB_PATH) -> list[BatchItem]:
    return [process_one(p, db_path=db_path) for p in paths]
