"""사업장별 수정본 양식 규칙 저장(core/registry.py source_templates)과 큐
분리/적용(core/batch.py) — 2026-10-08."""

from __future__ import annotations

import sys
from pathlib import Path

from core import registry
from core.batch import _template_for, plan_batch_add
from core.profile import parse_template_token, payroll_type_from_sheet_name, template_token
from core.source_template import learn_recipe

sys.path.insert(0, str(Path(__file__).parent))
import test_source_template as T  # noqa: E402


def test_category_reference_follows_payroll_type_and_folder(tmp_path):
    from core.batch import _registered_reference
    import shutil
    db, company, recipe, fixed = _setup(tmp_path)
    original = Path(company.reference_file_path)
    registry.save_payroll_reference(company.id, '급여', original, db_path=db)
    moved = tmp_path / 'moved'
    shutil.copytree(db.parent, moved)
    moved_db = moved / 'registry.db'
    co = registry.get_company_by_name(company.name, db_path=moved_db)
    path = tmp_path / '가나다센터 8월귀속 9월지급 급여대장(급여).xlsx'
    chosen = _registered_reference(co, path, None, None, moved_db)
    assert chosen.name == f'company_{company.id}_pay.xlsx'
    assert chosen.parent == moved / 'reference_files'
    assert _registered_reference(co, path, template_token('추급'), None, moved_db) == Path(co.reference_file_path)
    assert _registered_reference(co, path, None, [(path, '급여1'), (path, '급여2')], moved_db) == chosen

GOLDEN = Path(__file__).parent / "fixtures" / "golden"


def _setup(tmp_path: Path):
    db = tmp_path / "data" / "registry.db"
    company = registry.register_company(
        "가나다센터", GOLDEN / "sample3-wehago-before.xlsx", db_path=db, reference_dir=tmp_path / "data" / "ref"
    )
    orig, fixed = tmp_path / "가나다센터 8월귀속 9월지급 급여대장.xlsx", tmp_path / "fixed.xlsx"
    T._original(orig, 8, T.PEOPLE)
    T._fixed(fixed, T.PEOPLE)
    recipe = learn_recipe(orig, fixed)
    assert recipe.complete
    return db, company, recipe, fixed


def test_template_token_roundtrip() -> None:
    assert parse_template_token(template_token("급+상")) == "급+상"
    assert parse_template_token("급여대장") is None
    assert payroll_type_from_sheet_name(template_token("추급")) == "추급"


def test_save_and_get_template_copies_file(tmp_path: Path) -> None:
    db, company, recipe, fixed = _setup(tmp_path)
    stored = registry.save_source_template(company.id, "급여", recipe.to_json(), fixed, "원본.xlsx", db_path=db)
    assert stored.exists() and stored.parent == db.parent / "source_templates"
    got = registry.get_source_templates(company.id, db_path=db)
    assert set(got) == {"급여"}
    assert got["급여"][1] == stored
    registry.delete_company(company.id, db_path=db)
    assert registry.get_source_templates(company.id, db_path=db) == {}


def test_file_with_two_types_is_split_only_when_both_templates_exist(tmp_path: Path) -> None:
    db, company, recipe, fixed = _setup(tmp_path)
    both = tmp_path / "가나다센터 8월귀속 9월지급 급여대장(급여)(급+상).xlsx"
    T._original(both, 8, T.PEOPLE)

    registry.save_source_template(company.id, "급여", recipe.to_json(), fixed, db_path=db)
    assert plan_batch_add([both], db_path=db) == [[(both, None)]]  # 급+상 규칙이 없으면 나누지 않음

    registry.save_source_template(company.id, "급+상", recipe.to_json(), fixed, db_path=db)
    groups = plan_batch_add([both], db_path=db)
    assert groups == [[(both, template_token("급여"))], [(both, template_token("급+상"))]]

    tpl = _template_for(company.id, both, template_token("급+상"), None, db)
    assert tpl is not None and tpl[1].name.endswith("_bonus.xlsx")
    # 사용자가 시트를 직접 고른 경우, 합치는 항목, 종류가 둘인데 표시가 없는 경우엔 쓰지 않는다.
    assert _template_for(company.id, both, "급여대장", None, db) is None
    assert _template_for(company.id, both, None, None, db) is None
    assert _template_for(company.id, both, None, [(both, None), (both, None)], db) is None


def test_untyped_file_uses_untyped_template(tmp_path: Path) -> None:
    db, company, recipe, fixed = _setup(tmp_path)
    plain = tmp_path / "가나다센터 8월귀속 9월지급 급여대장.xlsx"
    assert _template_for(company.id, plain, None, None, db) is None
    registry.save_source_template(company.id, None, recipe.to_json(), fixed, db_path=db)
    assert _template_for(company.id, plain, None, None, db) is not None

def test_reference_follows_database_when_app_folder_moves(tmp_path: Path) -> None:
    import shutil
    db, company, _, _ = _setup(tmp_path)
    relocated = tmp_path / 'moved' / 'data'
    relocated.mkdir(parents=True)
    shutil.copy2(db, relocated / 'registry.db')
    refs = relocated / 'reference_files'
    refs.mkdir()
    local = refs / f'company_{company.id}.xlsx'
    shutil.copy2(company.reference_file_path, local)
    found = registry.get_company(company.id, db_path=relocated / 'registry.db')
    assert Path(found.reference_file_path) == local
