"""정답 파일(사람이 금액을 맞춰 실제로 위하고에 올린 정리본)에서 사업장별
수당/공제 매칭을 거꾸로 알아낸다 — 사용자 요청(2026-10-08): "작업파일 exel
파일에서 적용시켰던 수당, 공제 매칭을 다시 데이터베이스에 전체 저장".

방법: 정답 파일을 위하고 참고파일 자리에 그대로 넣고 한 번 처리하면
(core/reconcile.py의 `_LEARN_SINK`), 사원별로 어느 원본 행이 짝지어졌는지와
엔진이 자동으로 고른 열 매칭이 남는다. 원본의 각 항목(저장 매칭 키 단위 —
`learned_mapping_key(그룹, 항목명)`)을 사원별 금액 벡터로 보고, 정답의
사원별·열별 금액과 맞는 칸 수가 **엄격히 늘어나는** 경우에만 그 항목을 다른
위하고 열로 옮기거나(또는 매칭 해제), 두 항목의 자리를 맞바꾼다. 맞는 칸이
늘지 않는 변경은 절대 받아들이지 않으므로, 우연히 값이 같은 열로 흘러가는
일은 금액이 정확히 맞아떨어지는 경우로만 제한된다.

결과는 저장 매칭 키 -> 위하고 열 키("" = 매칭 안 함) 사전이며,
`core/registry.py::save_column_mapping`으로 그대로 저장할 수 있다."""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field

import core.reconcile as _rec
from core.canon import canon, norm
from core.classify import group_of_src
from core.reconcile import _as_number, _block_number, learned_mapping_key
from core.sheet_layout import source_item

_TOL = 1.0
_TOTAL_WORDS = ("소계", "합계", "총계", "총액", "총지급", "지급총", "공제총", "실지급", "실수령", "차인")
_TOTAL_CANONS = {"지급액계", "공제액계", "차인지급액"}


def collect_contexts(run):
    """`run()`(엔진 처리 한 번 — 합치기면 시트 여러 개)을 실행하는 동안
    reconcile()이 남긴 학습용 정보를 모아 (run의 반환값, 정보 목록)을 준다."""
    prev = _rec._LEARN_SINK
    sink: list = []
    _rec._LEARN_SINK = sink
    try:
        out = run()
    finally:
        _rec._LEARN_SINK = prev
    return out, sink


def _is_total_like(item: str) -> bool:
    k = norm(item)
    if not k:
        return True
    if canon(item) in _TOTAL_CANONS:
        return True
    if k.endswith("계") and len(k) > 1:
        return True
    return any(w in k for w in _TOTAL_WORDS)


@dataclass
class LearnResult:
    # 저장 매칭 키(norm 적용) -> 학습 후 위하고 열 키("" = 매칭 안 함)
    assignment: dict[str, str] = field(default_factory=dict)
    # 같은 키의 엔진 기본값(학습 전) — 저장할 때 기본값과 같은 건 생략할 수 있게
    initial: dict[str, str] = field(default_factory=dict)
    # 사람이 보기 쉬운 원문 키(그룹|항목명)
    labels: dict[str, str] = field(default_factory=dict)
    mismatches_before: int = 0
    mismatches_after: int = 0
    cells: int = 0
    verified: set[str] = field(default_factory=set)

    @property
    def changed(self) -> dict[str, str]:
        return {k: v for k, v in self.assignment.items() if self.initial.get(k, "") != v}


def learn_from_contexts(contexts: list[dict]) -> LearnResult | None:
    if not contexts:
        return None
    dest_ws = contexts[0]["dest_ws"]
    col_to_key: dict[int, str] = contexts[0]["col_to_key"]
    dkeys = sorted(set(col_to_key.values()))
    emps = sorted({p[0] for ctx in contexts for p in ctx["pairs"]})
    if not emps:
        return None
    emp_idx = {e: i for i, e in enumerate(emps)}
    n = len(emps)

    target: dict[str, list[float]] = {k: [0.0] * n for k in dkeys}
    for col, key in col_to_key.items():
        for e in emps:
            target[key][emp_idx[e]] += _as_number(dest_ws.cell(row=e, column=col).value) or 0.0

    # 고정 기여(학습 대상이 아닌 매칭 열) + 단위(저장 키)별 벡터와 현재 위치
    fixed: dict[str, list[float]] = {k: [0.0] * n for k in dkeys}
    unit_vec: dict[str, list[float]] = {}
    unit_parts: dict[str, list[tuple[str | None, list[float]]]] = defaultdict(list)
    labels: dict[str, str] = {}

    for ctx in contexts:
        ws, layout = ctx["source_ws"], ctx["source_layout"]
        cmap: dict[int, int] = ctx["column_map"]
        ctk: dict[int, str] = ctx["col_to_key"]
        bh = layout.block_height
        cand: dict[int, str] = {}
        for c in range(1, ws.max_column + 1):
            g = group_of_src(ws, layout, c)
            if g not in ("수당", "공제", ""):
                continue
            item = source_item(ws, layout, c)
            if not item or _is_total_like(item):
                continue
            raw = learned_mapping_key(g, item)
            cand[c] = norm(raw)
            labels.setdefault(norm(raw), raw)
        cols = set(cand) | set(cmap)
        for c in cols:
            vec = [0.0] * n
            for dest_row, src_row, cont in ctx["pairs"]:
                vec[emp_idx[dest_row]] += _block_number(ws, src_row, bh, c, cont) or 0.0
            cur = ctk.get(cmap[c]) if c in cmap else None
            if c in cand:
                if not any(abs(v) > _TOL for v in vec) and cur is None:
                    continue
                unit_parts[cand[c]].append((cur, vec))
            elif cur is not None:
                f = fixed[cur]
                for i in range(n):
                    f[i] += vec[i]

    if not unit_parts:
        return None
    for k, parts in unit_parts.items():
        v = [0.0] * n
        for _, pv in parts:
            for i in range(n):
                v[i] += pv[i]
        unit_vec[k] = v

    # 시작 상태: 각 단위의 현재 위치(열마다 다르면 열 단위로 그대로 반영)
    pred = {k: list(fixed[k]) for k in dkeys}
    initial: dict[str, str] = {}
    pos: dict[str, str | None] = {}
    for k, parts in unit_parts.items():
        dests = {d for d, _ in parts}
        for d, pv in parts:
            if d is not None:
                p = pred[d]
                for i in range(n):
                    p[i] += pv[i]
        initial[k] = (next(iter(dests)) or "") if len(dests) == 1 else "\x00mixed"
        pos[k] = next(iter(dests)) if len(dests) == 1 else "\x00mixed"

    def bad(col_vals: list[float], key: str) -> int:
        t = target[key]
        return sum(1 for i in range(n) if abs(col_vals[i] - t[i]) > _TOL)

    col_bad = {k: bad(pred[k], k) for k in dkeys}
    before = sum(col_bad.values())

    def remove_unit(k: str, cur_pred: dict[str, list[float]]) -> None:
        # 단위 k의 현재 기여를 cur_pred에서 뺀다(mixed면 열별 원래 위치에서)
        if pos[k] == "\x00mixed":
            for d, pv in unit_parts[k]:
                if d is not None:
                    p = cur_pred[d]
                    for i in range(n):
                        p[i] -= pv[i]
        elif pos[k] is not None:
            p = cur_pred[pos[k]]
            v = unit_vec[k]
            for i in range(n):
                p[i] -= v[i]

    def touched(k: str) -> set[str]:
        if pos[k] == "\x00mixed":
            return {d for d, _ in unit_parts[k] if d is not None}
        return {pos[k]} if pos[k] is not None else set()

    def evaluate(moves: list[tuple[str, str | None]]) -> int:
        keys_aff: set[str] = set()
        for k, o in moves:
            keys_aff |= touched(k)
            if o is not None:
                keys_aff.add(o)
        tmp = {d: list(pred[d]) for d in keys_aff}
        for k, _ in moves:
            remove_unit(k, tmp)
        for k, o in moves:
            if o is not None:
                p, v = tmp[o], unit_vec[k]
                for i in range(n):
                    p[i] += v[i]
        return sum(bad(tmp[d], d) for d in keys_aff) - sum(col_bad[d] for d in keys_aff)

    def apply(moves: list[tuple[str, str | None]]) -> None:
        keys_aff: set[str] = set()
        for k, o in moves:
            keys_aff |= touched(k)
            if o is not None:
                keys_aff.add(o)
        for k, _ in moves:
            remove_unit(k, pred)
        for k, o in moves:
            if o is not None:
                p, v = pred[o], unit_vec[k]
                for i in range(n):
                    p[i] += v[i]
            pos[k] = o
        for d in keys_aff:
            col_bad[d] = bad(pred[d], d)

    units = sorted(unit_vec)
    options: list[str | None] = [None, *dkeys]
    for _ in range(300):
        if sum(col_bad.values()) == 0:
            break
        best: tuple[int, list] = (0, [])
        for k in units:
            for o in options:
                if o == pos[k]:
                    continue
                d = evaluate([(k, o)])
                if d < best[0]:
                    best = (d, [(k, o)])
        if not best[1]:
            # 하나만 옮겨서는 나아지지 않을 때 — 두 항목 자리 맞바꾸기
            for a in range(len(units)):
                ka = units[a]
                if pos[ka] == "\x00mixed":
                    continue
                for b in range(a + 1, len(units)):
                    kb = units[b]
                    if pos[kb] == "\x00mixed" or pos[ka] == pos[kb]:
                        continue
                    d = evaluate([(ka, pos[kb]), (kb, pos[ka])])
                    if d < best[0]:
                        best = (d, [(ka, pos[kb]), (kb, pos[ka])])
        if not best[1]:
            break
        apply(best[1])

    assignment = {k: (pos[k] or "") for k in units if pos[k] != "\x00mixed"}
    return LearnResult(
        assignment=assignment,
        initial={k: v for k, v in initial.items() if v != "\x00mixed"},
        labels={k: labels.get(k, k) for k in units},
        mismatches_before=before,
        mismatches_after=sum(col_bad.values()),
        cells=n * len(dkeys),
        verified={k for k, dest in assignment.items() if dest and col_bad[dest] == 0
                  and any(abs(v) > _TOL for v in unit_vec[k])},
    )
