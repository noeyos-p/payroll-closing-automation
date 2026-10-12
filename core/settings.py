"""앱 전역 설정 — 지금까지 하드코딩돼 있던 경로들(위하고 양식 폴더, 위하고
자동 업로드용 엑셀 저장 폴더)을 환경설정 화면에서 바꿀 수 있게 저장한다.
`data/settings.json`(사업장 등록 DB `registry.db`와 같은 자리)에 저장하는
단순 JSON — 새 의존성 없이 표준 라이브러리만 쓴다."""
from __future__ import annotations

import json
from dataclasses import asdict, dataclass
from pathlib import Path

_SETTINGS_PATH = Path(__file__).resolve().parent.parent / "data" / "settings.json"

# 기존에 core/wehago_sync.py에 하드코딩돼 있던 값 그대로 — 처음 실행 시
# 기존 사용자가 아무것도 안 바꿔도 이전과 똑같이 동작하게 하는 기본값.
DEFAULT_WEHAGO_FORM_FOLDER = r"C:\Users\<사용자>\Desktop\작업파일\Wehago"
DEFAULT_UPLOAD_EXPORT_FOLDER = r"C:\Users\<사용자>\Desktop\작업파일\Exel"


@dataclass
class Settings:
    wehago_form_folder: str = DEFAULT_WEHAGO_FORM_FOLDER
    upload_export_folder: str = DEFAULT_UPLOAD_EXPORT_FOLDER
    # 위하고 자동 로그인용 — 평문 저장(사용자 요청, 2026-09-21). 이 앱은
    # 개인용 로컬 도구라 암호화보다 평문 하드코딩을 선호한다는 기존 방침과
    # 같은 이유(이식성 — PC를 옮기거나 백업/복원할 때 별도 키 관리 없이
    # 그대로 옮겨 쓸 수 있음). 비워두면 기존처럼 매번 수동 로그인으로
    # 폴백한다.
    wehago_id: str = ""
    wehago_password: str = ""


def load_settings() -> Settings:
    """파일이 없거나 읽기/파싱에 실패하면 조용히 기본값으로 폴백한다 —
    설정 파일 하나 깨졌다고 앱 전체가 못 뜨면 안 되므로."""
    if _SETTINGS_PATH.exists():
        try:
            data = json.loads(_SETTINGS_PATH.read_text(encoding="utf-8"))
            defaults = asdict(Settings())
            return Settings(**{**defaults, **{k: v for k, v in data.items() if k in defaults}})
        except Exception:
            pass
    return Settings()


def save_settings(settings: Settings) -> None:
    _SETTINGS_PATH.parent.mkdir(parents=True, exist_ok=True)
    _SETTINGS_PATH.write_text(
        json.dumps(asdict(settings), ensure_ascii=False, indent=2), encoding="utf-8"
    )
