"""위하고 참고파일을 모아두는 폴더를 통째로 스캔해서 등록된 사업장을 한 번에
갱신한다. 지금까지는 "위하고 파일 선택"/"사업장 관리"에서 파일 하나씩만
등록/재등록할 수 있었는데, 사용자가 "`작업파일\\Wehago` 폴더에 최신 양식을
전부 넣어둘 테니 버튼 하나로 한꺼번에 업데이트되게 해달라"고 요청해서 추가함
(2026-09-08). `_wontax_register.py`(원천세 신고리스트 대량 등록 스크립트)에서
이미 검증된 "업체명-YYYYMM.xlsx" 파일명 관례 + 같은 업체 중복 시 최신 수정일
파일 선택 로직을 그대로 재사용한다."""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

from core import registry

DEFAULT_FOLDER = Path(r"C:\Users\<사용자>\Desktop\작업파일\Wehago")

# "업체명-YYYYMM.xlsx" / "업체명 (YYYY)-YYYYMM.xlsx" / 끝에 " (n)" 중복다운로드
# 접미사까지 허용(_wontax_register.py와 동일한 패턴). 날짜 접미사가 아예 없는
# "업체명.xlsx"도 이 폴더에서는 흔할 수 있어 두 번째 패턴으로 폴백한다.
_DATED_RE = re.compile(
    r"^(?P<name>.+?)\s*(?:\((?:19|20)\d{2}\))?-(?:19|20)\d{2}(?:0[1-9]|1[0-2])(?:\s*\(\d+\))?\.(?:xlsx|xlsm)$",
    re.IGNORECASE,
)
_DUP_SUFFIX_RE = re.compile(r"\s*\(\d+\)$")


def _normalize(name: str) -> str:
    # "㈜"(원문자)와 "(주)"(괄호+글자)는 같은 뜻의 다른 표기라 안전하게 같은
    # 업체로 합친다. 그 외 괄호 안 메모("(신)"/"(반기)" 등)는 실제로 다른
    # 지점을 구분하는 데도 쓰이는 걸 확인해서(CLAUDE.md 참고) 벗기지 않는다.
    return re.sub(r"\s+", "", name).replace("㈜", "(주)").strip()


def _guess_company_name(filename: str) -> str:
    match = _DATED_RE.match(filename)
    if match:
        return match.group("name").strip()
    stem = Path(filename).stem
    return _DUP_SUFFIX_RE.sub("", stem).strip()


@dataclass
class SyncResult:
    registered: list[str] = field(default_factory=list)
    failed: list[tuple[str, str]] = field(default_factory=list)  # (사업장명, 오류메시지)
    skipped_files: list[str] = field(default_factory=list)  # 이름을 못 뽑은 파일명
    folder_missing: bool = False


def sync_from_folder(folder: Path = DEFAULT_FOLDER, db_path: Path = registry.DEFAULT_DB_PATH) -> SyncResult:
    """`folder` 안의 위하고 양식 파일을 전부 스캔해서, 파일명에서 뽑은 사업장명으로
    등록/갱신한다. 같은 사업장 이름으로 파일이 여러 개면(예: 이번 달/지난 달 양식이
    같이 남아있는 경우) 파일 자체의 수정일시가 가장 최신인 것만 쓴다. 한 파일이
    실패해도 나머지는 계속 처리한다(배치 처리 원칙과 동일)."""
    result = SyncResult()
    if not folder.is_dir():
        result.folder_missing = True
        return result

    candidates: dict[str, list[Path]] = {}
    for path in list(folder.glob("*.xlsx")) + list(folder.glob("*.xlsm")):
        name = _guess_company_name(path.name)
        if not name:
            result.skipped_files.append(path.name)
            continue
        candidates.setdefault(_normalize(name), []).append(path)

    for paths in candidates.values():
        latest = max(paths, key=lambda p: p.stat().st_mtime)
        name = _guess_company_name(latest.name)
        try:
            registry.register_company(name, latest, db_path=db_path)
            result.registered.append(name)
        except Exception as exc:
            result.failed.append((name, str(exc)))

    result.registered.sort()
    return result
