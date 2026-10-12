"""Merge-aware cell text lookup, shared by core/sheet_layout.py and
core/classify.py. Ported from the VBA macro's `H(ws, r, c)` — Excel only
stores a value on the top-left cell of a merged range, so a naive
`.cell(row, col).value` read on any other cell in that range comes back
None even though the header clearly has a label there (real case: the
destination sheet's group headers like "수당"/"공제" span many merged
columns; the source sheet's "임금상세내역" group header does too)."""

from __future__ import annotations

import weakref

# 시트별 (행, 열) -> 병합 범위의 왼쪽 위 셀 좌표. 예전엔 셀 하나 읽을 때마다
# 병합 범위 전체를 처음부터 훑어서, 병합이 많은 원본은 파일 하나 처리에
# 수십 초씩 걸렸다(2026-10-07 전체 급여대장 일괄 검증 중 확인). 결과는 같고
# 속도만 다르다 — 이 앱은 읽은 워크시트를 수정하지 않으므로 캐시가 낡지 않는다.
_MERGE_CACHE: "weakref.WeakKeyDictionary[object, tuple[dict, list]]" = weakref.WeakKeyDictionary()
_BIG_MERGE_AREA = 20000


def _merge_index(ws) -> tuple[dict[tuple[int, int], tuple[int, int]], list]:
    try:
        cached = _MERGE_CACHE.get(ws)
    except TypeError:
        cached = None
    if cached is not None:
        return cached
    anchors: dict[tuple[int, int], tuple[int, int]] = {}
    big: list = []
    for rng in ws.merged_cells.ranges:
        if (rng.max_row - rng.min_row + 1) * (rng.max_col - rng.min_col + 1) > _BIG_MERGE_AREA:
            # 시트 대부분을 덮는 비정상적으로 큰 병합은 펼치지 않고 따로 둔다.
            big.append(rng)
            continue
        anchor = (rng.min_row, rng.min_col)
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                anchors.setdefault((r, c), anchor)
    result = (anchors, big)
    try:
        _MERGE_CACHE[ws] = result
    except TypeError:
        pass
    return result


def cell_text(ws, row: int, col: int) -> str | None:
    if row < 1 or col < 1:
        return None
    value = ws.cell(row=row, column=col).value
    if value is not None:
        return value
    if getattr(ws, "merged_cells", None) is None:
        return None
    anchors, big = _merge_index(ws)
    anchor = anchors.get((row, col))
    if anchor is not None:
        return ws.cell(row=anchor[0], column=anchor[1]).value
    for rng in big:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            return ws.cell(row=rng.min_row, column=rng.min_col).value
    return None
