"""사업장 이름 ↔ 위하고 참고파일 경로를 저장하는 최소 등록 DB.

이 앱은 원래 "아무것도 저장하지 않는다"는 원칙이었다(매번 원본+위하고 파일 둘 다
직접 골라 그 자리에서 처리) — 그런데 매번 위하고 파일을 직접 고르는 게 번거롭다는
사용자 요청으로, **급여대장 파일만 올려도 사업장을 자동 인식해서 미리 등록해둔
위하고 참고파일을 자동으로 쓰는** 기능을 위해 딱 이 용도로만 SQLite 등록 DB를
둔다. 사원/항목/금액은 저장하지 않는다 — canon 기반 매칭 엔진(core/reconcile.py)은
과거 상태가 전혀 필요 없어서, 여기 저장하는 건 "사업장 이름 인식용 사원 이름
지문"과 "위하고 참고파일 경로"뿐이다.

등록은 명시적 버튼이 아니라 **자동**이다(사용자 지시) — 위하고 참고파일을 사람이
직접 골라서 처리에 성공하면 그 사업장명으로 조용히 upsert된다(ui/main_window.py).
"""
from __future__ import annotations

import datetime
import json
import shutil
import sqlite3
from dataclasses import dataclass
from pathlib import Path

from core.canon import norm
from core.engine import list_dest_employee_names
from core.name_match import name_base3

_APP_DATA_DIR = Path(__file__).resolve().parent.parent / "data"
DEFAULT_DB_PATH = _APP_DATA_DIR / "registry.db"
DEFAULT_REFERENCE_DIR = _APP_DATA_DIR / "reference_files"


@dataclass
class Company:
    id: int
    name: str
    reference_file_path: str
    reference_file_name: str
    registered_at: str
    base_names: set[str]

    @property
    def employee_count(self) -> int:
        return len(self.base_names)


def _connect(db_path: Path) -> sqlite3.Connection:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS companies (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            reference_file_path TEXT NOT NULL DEFAULT '',
            reference_file_name TEXT NOT NULL DEFAULT '',
            registered_at TEXT NOT NULL DEFAULT '',
            employee_base_names_json TEXT NOT NULL DEFAULT '[]'
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS column_mappings (
            company_id INTEGER NOT NULL,
            source_header_norm TEXT NOT NULL,
            dest_key TEXT NOT NULL,
            PRIMARY KEY (company_id, source_header_norm)
        )
        """
    )
    # 사업장별 "원본 -> 사람이 손으로 다시 짠 수정본 양식" 변환 규칙
    # (core/source_template.py, 2026-10-08). payroll_type은 급여/급+상/추급
    # 또는 ""(종류 구분 없음). 수정본 파일 사본은 DB 옆 source_templates/에
    # 두고 파일 이름만 저장한다(앱 폴더를 옮겨도 그대로 찾도록).
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS source_templates (
            company_id INTEGER NOT NULL,
            payroll_type TEXT NOT NULL DEFAULT '',
            recipe_json TEXT NOT NULL,
            template_file TEXT NOT NULL,
            source_file_name TEXT NOT NULL DEFAULT '',
            updated_at TEXT NOT NULL DEFAULT '',
            PRIMARY KEY (company_id, payroll_type)
        )
        """
    )
    conn.execute("""CREATE TABLE IF NOT EXISTS payroll_references (
        company_id INTEGER NOT NULL, payroll_type TEXT NOT NULL,
        reference_file TEXT NOT NULL, PRIMARY KEY (company_id, payroll_type))""")
    return conn


def save_payroll_reference(company_id: int, payroll_type: str, reference_path: Path,
                           db_path: Path = DEFAULT_DB_PATH) -> Path:
    """Keep the verified destination schema for each payroll category."""
    tag = {'': 'all', '급여': 'pay', '급+상': 'bonus', '추급': 'extra'}[payroll_type]
    target = db_path.parent / 'reference_files' / f'company_{company_id}_{tag}.xlsx'
    target.parent.mkdir(parents=True, exist_ok=True)
    if reference_path.resolve() != target.resolve():
        shutil.copy2(reference_path, target)
    with _connect(db_path) as conn:
        conn.execute('INSERT INTO payroll_references VALUES (?,?,?) ON CONFLICT(company_id,payroll_type) DO UPDATE SET reference_file=excluded.reference_file',
                     (company_id, payroll_type, target.name))
    return target


def get_payroll_reference(company_id: int, payroll_type: str,
                          db_path: Path = DEFAULT_DB_PATH) -> Path | None:
    with _connect(db_path) as conn:
        row = conn.execute('SELECT reference_file FROM payroll_references WHERE company_id=? AND payroll_type=?',
                           (company_id, payroll_type)).fetchone()
    path = db_path.parent / 'reference_files' / row[0] if row else None
    return path if path is not None and path.is_file() else None


def _template_dir(db_path: Path) -> Path:
    return db_path.parent / "source_templates"


def save_source_template(
    company_id: int,
    payroll_type: str | None,
    recipe_json: dict,
    template_path: Path,
    source_file_name: str = "",
    db_path: Path = DEFAULT_DB_PATH,
) -> Path:
    """수정본 양식 변환 규칙을 저장/갱신한다(같은 사업장+종류면 덮어씀).
    `template_path`(사람이 고친 수정본)는 DB 옆 폴더로 복사해 둔다."""
    ptype = payroll_type or ""
    folder = _template_dir(db_path)
    folder.mkdir(parents=True, exist_ok=True)
    safe_type = {"": "all", "급여": "pay", "급+상": "bonus", "추급": "extra"}.get(ptype, "other")
    stable = folder / f"company_{company_id}_{safe_type}.xlsx"
    if Path(template_path).resolve() != stable.resolve():
        shutil.copyfile(template_path, stable)
    conn = _connect(db_path)
    try:
        conn.execute(
            "INSERT INTO source_templates (company_id, payroll_type, recipe_json, template_file, source_file_name, updated_at) "
            "VALUES (?, ?, ?, ?, ?, ?) ON CONFLICT(company_id, payroll_type) DO UPDATE SET "
            "recipe_json=excluded.recipe_json, template_file=excluded.template_file, "
            "source_file_name=excluded.source_file_name, updated_at=excluded.updated_at",
            (
                company_id, ptype, json.dumps(recipe_json, ensure_ascii=False), stable.name,
                source_file_name, datetime.datetime.now().isoformat(timespec="seconds"),
            ),
        )
        conn.commit()
    finally:
        conn.close()
    return stable


def get_source_templates(company_id: int, db_path: Path = DEFAULT_DB_PATH) -> dict[str, tuple[dict, Path]]:
    """{급여종류("" 포함): (변환 규칙 JSON, 수정본 양식 파일 경로)} — 파일이
    없어진 항목은 뺀다."""
    conn = _connect(db_path)
    try:
        rows = conn.execute(
            "SELECT payroll_type, recipe_json, template_file FROM source_templates WHERE company_id = ?",
            (company_id,),
        ).fetchall()
    finally:
        conn.close()
    out: dict[str, tuple[dict, Path]] = {}
    for ptype, recipe_json, fname in rows:
        path = _template_dir(db_path) / fname
        if path.exists():
            out[ptype] = (json.loads(recipe_json), path)
    return out


def delete_source_template(company_id: int, payroll_type: str | None, db_path: Path = DEFAULT_DB_PATH) -> None:
    conn = _connect(db_path)
    try:
        conn.execute(
            "DELETE FROM source_templates WHERE company_id = ? AND payroll_type = ?", (company_id, payroll_type or "")
        )
        conn.commit()
    finally:
        conn.close()


def _row_to_company(row: tuple, db_path: Path = DEFAULT_DB_PATH) -> Company:
    id_, name, ref_path, ref_name, registered_at, base_names_json = row
    local_ref = db_path.parent / "reference_files" / f"company_{id_}.xlsx"
    if local_ref.is_file():
        ref_path = str(local_ref)
    return Company(id_, name, ref_path, ref_name, registered_at, set(json.loads(base_names_json)))


def register_company(
    name: str,
    reference_path: Path,
    reference_sheet: str | None = None,
    db_path: Path = DEFAULT_DB_PATH,
    reference_dir: Path = DEFAULT_REFERENCE_DIR,
) -> Company:
    """이름 + 위하고 참고파일로 사업장을 등록/갱신한다(upsert — 같은 이름이
    이미 있으면 참고파일과 사원 지문을 새로 갱신). 원본 파일이 나중에 지워지거나
    옮겨져도 안전하게 재처리할 수 있도록 앱 데이터 폴더로 복사해둔다."""
    name = name.strip()
    if not name:
        raise ValueError("사업장 이름이 비어 있습니다.")
    reference_dir.mkdir(parents=True, exist_ok=True)

    conn = _connect(db_path)
    try:
        existing = conn.execute("SELECT id FROM companies WHERE name = ?", (name,)).fetchone()
        if existing:
            company_id = existing[0]
        else:
            cur = conn.execute("INSERT INTO companies (name) VALUES (?)", (name,))
            company_id = cur.lastrowid

        stable_path = reference_dir / f"company_{company_id}.xlsx"
        shutil.copyfile(reference_path, stable_path)

        names = list_dest_employee_names(stable_path, reference_sheet)
        base_names = sorted({b for n in names if (b := name_base3(n))})

        conn.execute(
            "UPDATE companies SET reference_file_path=?, reference_file_name=?, "
            "registered_at=?, employee_base_names_json=? WHERE id=?",
            (
                str(stable_path),
                reference_path.name,
                datetime.datetime.now().isoformat(timespec="seconds"),
                json.dumps(base_names, ensure_ascii=False),
                company_id,
            ),
        )
        conn.commit()
        row = conn.execute(
            "SELECT id, name, reference_file_path, reference_file_name, registered_at, "
            "employee_base_names_json FROM companies WHERE id = ?",
            (company_id,),
        ).fetchone()
        return _row_to_company(row, db_path)
    finally:
        conn.close()


def list_companies(db_path: Path = DEFAULT_DB_PATH) -> list[Company]:
    conn = _connect(db_path)
    try:
        rows = conn.execute(
            "SELECT id, name, reference_file_path, reference_file_name, registered_at, "
            "employee_base_names_json FROM companies ORDER BY name"
        ).fetchall()
        return [_row_to_company(r, db_path) for r in rows]
    finally:
        conn.close()


def get_company(company_id: int, db_path: Path = DEFAULT_DB_PATH) -> Company | None:
    conn = _connect(db_path)
    try:
        row = conn.execute(
            "SELECT id, name, reference_file_path, reference_file_name, registered_at, "
            "employee_base_names_json FROM companies WHERE id = ?",
            (company_id,),
        ).fetchone()
        return _row_to_company(row, db_path) if row else None
    finally:
        conn.close()


def get_company_by_name(name: str, db_path: Path = DEFAULT_DB_PATH) -> Company | None:
    conn = _connect(db_path)
    try:
        row = conn.execute(
            "SELECT id, name, reference_file_path, reference_file_name, registered_at, "
            "employee_base_names_json FROM companies WHERE name = ?",
            (name,),
        ).fetchone()
        return _row_to_company(row, db_path) if row else None
    finally:
        conn.close()


def delete_company(company_id: int, db_path: Path = DEFAULT_DB_PATH) -> None:
    conn = _connect(db_path)
    try:
        conn.execute("DELETE FROM companies WHERE id = ?", (company_id,))
        conn.execute("DELETE FROM column_mappings WHERE company_id = ?", (company_id,))
        conn.execute("DELETE FROM source_templates WHERE company_id = ?", (company_id,))
        conn.execute("DELETE FROM payroll_references WHERE company_id = ?", (company_id,))
        conn.commit()
    finally:
        conn.close()


def save_column_mapping(company_id: int, source_header: str, dest_key: str, db_path: Path = DEFAULT_DB_PATH) -> None:
    """사업장별로 손으로 고른 항목 매칭(항목매핑결과 탭의 "직접 매칭")을
    소스 헤더 텍스트(정규화)와 위하고 목적지 키로 기억해둔다 — 다음 달
    같은 사업장 급여대장에서 canon.py가 자동으로 못 잡는 같은 항목이 다시
    나오면, `core/reconcile.py`의 `learned_mappings`로 자동 적용된다
    (사용자 요청, 2026-09-08). 같은 (사업장, 헤더)에 이미 저장된 게 있으면
    최신 선택으로 덮어쓴다(upsert).

    `dest_key`로 빈 문자열("")을 넘기면 "이 항목은 이 사업장에서 항상
    매칭하지 마라"(콤보박스의 "선택 안 함")는 뜻으로 저장된다 — 컬럼이
    `NOT NULL`이라 파이썬 `None` 대신 빈 문자열을 그 자리 대용으로 쓴다.
    `core/reconcile.py::reconcile()`이 이 값을 보면 `build_map`이 뭘로
    자동매칭했든 강제로 미매칭 처리한다(사용자 요청, 2026-09-14 — 예전엔
    "선택 안 함"은 일회성 판단이라며 저장하지 않았는데, 그러면 다음 실행/
    빌드마다 매번 다시 손으로 꺼줘야 했다)."""
    key = norm(source_header)
    if not key:
        return
    conn = _connect(db_path)
    try:
        conn.execute(
            "INSERT INTO column_mappings (company_id, source_header_norm, dest_key) VALUES (?, ?, ?) "
            "ON CONFLICT(company_id, source_header_norm) DO UPDATE SET dest_key=excluded.dest_key",
            (company_id, key, dest_key),
        )
        conn.commit()
    finally:
        conn.close()


def get_column_mappings(company_id: int, db_path: Path = DEFAULT_DB_PATH) -> dict[str, str]:
    """{정규화된 소스 헤더 텍스트: 위하고 목적지 키} — 이 사업장에 저장된
    전부. 비어있으면(신규 사업장이거나 아직 손으로 고친 적 없으면) 빈 dict."""
    conn = _connect(db_path)
    try:
        rows = conn.execute(
            "SELECT source_header_norm, dest_key FROM column_mappings WHERE company_id = ?", (company_id,)
        ).fetchall()
        return {row[0]: row[1] for row in rows}
    finally:
        conn.close()
