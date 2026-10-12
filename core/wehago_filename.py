"""이 앱이 만드는 "위하고 업로드용" 다운로드 파일명을 다시 파싱한다.

`ui/main_window.py::_download_label`/`_on_download_current`/
`_export_items_to_folder`가 실제로 쓰는 규칙 그대로다:

    f"{date.today():%y%m%d}_{거래처명}{귀속월:02d}({지급구분})?.xlsx"

예: "260921_업체00708(급여).xlsx" — 날짜(6자리) 뒤 "_" 다음은 전부
거래처명이고, 그 거래처명 뒤에 곧바로(공백 없이) 귀속월 2자리 숫자가
붙는다. 지급구분("(급여)"/"(급+상)"/"(추급)")은 그 달에 지급건이 하나뿐이면
아예 안 붙을 수 있다(사용자 확인, 2026-09-21) — 이 경우 `pay_type`은
`None`이고, 위하고 쪽 지급일자 목록에 그 달 지급건이 정말 하나뿐이라고
가정하고 그걸 그대로 연다.

거래처명 자체에 숫자가 섞여 있어도(예: "업체071") 문제
없다 — 정규식이 "귀속월"을 "마지막 2자리"로 고정해서 뽑기 때문에, 거래처명
파싱은 그 앞부분을 전부 그대로 가져가면 된다(사용자 확인 사례,
2026-09-21)."""
from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

_NAME_RE = re.compile(
    r"^(?P<date>\d{6})_(?P<company>.+?)(?P<month>\d{2})(?:\((?P<paytype>[^)]+)\))?\.xlsx$",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class UploadFileInfo:
    path: Path
    company: str
    month: int  # 1-12, 귀속월
    pay_type: str | None  # "급여" / "급+상" / "추급" / None(그 달 지급건이 하나뿐)


def parse_upload_filename(path: Path) -> UploadFileInfo | None:
    """규칙에 안 맞는 파일명이면 조용히 `None`을 돌려준다(추측하지 않음) —
    호출부가 "이 파일은 자동 인식이 안 됩니다"로 표시하고 건너뛰게 한다."""
    match = _NAME_RE.match(path.name)
    if not match:
        return None
    company = match.group("company").strip()
    if not company:
        return None
    month = int(match.group("month"))
    if not 1 <= month <= 12:
        return None
    return UploadFileInfo(path=path, company=company, month=month, pay_type=match.group("paytype"))
