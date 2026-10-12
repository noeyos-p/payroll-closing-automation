"""사용자 요청(2026-09-14): 항목매핑결과 탭에서 자동매칭된 항목을 "선택 안
함"으로 바꿔서 저장하면(`core/registry.py::save_column_mapping`이 빈 문자열
dest_key로 저장), 다음 실행(그리고 이 앱을 새로 빌드한 뒤)에도 그 항목이
build_map의 자동매칭 결과와 무관하게 계속 미매칭으로 남아야 한다 — 예전엔
"선택 안 함"은 일회성 판단이라며 저장 자체를 안 했다.

`core/reconcile.py::reconcile()`이 `learned_mappings`에서 빈 문자열 값을
"이 항목은 항상 매칭하지 마라"로 해석해 강제로 미매칭 처리하는지, 실제 골든
픽스처(sample3)로 검증한다."""

from __future__ import annotations

from pathlib import Path

from core.engine import process
from core.reconcile import learned_mapping_key
from core.canon import norm

FIXTURES = Path(__file__).parent / "fixtures" / "golden"


def _matched_row(mapping_rows, item_label: str):
    for row in mapping_rows:
        if row.kind != "미매칭" and row.item_label == item_label:
            return row
    raise AssertionError(f"expected an auto-matched row for {item_label!r}")


def test_without_learned_mappings_기본급_is_auto_matched() -> None:
    result = process("sample3", FIXTURES / "sample3-source.xlsx", FIXTURES / "sample3-wehago-before.xlsx")
    assert result is not None
    row = _matched_row(result.mapping_rows, "기본급")
    assert row.dest_key == "기본급"


def test_learned_mapping_with_empty_dest_key_forces_unmapped() -> None:
    # 1st pass without learned_mappings just to find the real (group, item_label,
    # source_col) for 기본급 in this fixture — mirrors what
    # ui/main_window.py::_on_apply_mapping_tab does with `item.result.mapping_rows`.
    baseline = process("sample3", FIXTURES / "sample3-source.xlsx", FIXTURES / "sample3-wehago-before.xlsx")
    assert baseline is not None
    base_row = _matched_row(baseline.mapping_rows, "기본급")
    key = norm(learned_mapping_key(base_row.group, base_row.item_label))

    result = process(
        "sample3",
        FIXTURES / "sample3-source.xlsx",
        FIXTURES / "sample3-wehago-before.xlsx",
        learned_mappings={key: ""},
    )
    assert result is not None

    # 더 이상 매칭된 행에 없어야 한다.
    still_matched = [r for r in result.mapping_rows if r.kind != "미매칭" and r.item_label == "기본급"]
    assert still_matched == []

    # 대신 미매칭 목록에, "저장된 매칭 안 함" 사유가 붙은 채로 남아야 한다
    # (원래 사유인 "자동 매칭하지 못했습니다"가 아니라).
    unmapped = [r for r in result.mapping_rows if r.kind == "미매칭" and r.item_label == "기본급"]
    assert len(unmapped) == 1
    assert "매칭 안 함" in unmapped[0].note
    assert unmapped[0].source_col == base_row.source_col

    # 실제 계산에도 반영돼야 한다 — 기본급 열이 통째로 빠졌으니 최소 한 명은
    # 원본 대비 지급액계가 달라져 "차이"로 뜬다(자동매칭됐을 땐 golden
    # 픽스처가 "정상"이었던 것과 대비).
    assert any(row.status in ("차이", "추가됨_차이") for row in result.verification_rows)
