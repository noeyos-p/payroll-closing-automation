"""사업장별 "원본 -> 이지케어 표준양식" 변환 규칙(2026-10-08 사용자 요청).

배경: 엔진이 원본 급여대장을 제대로 못 읽는 사업장은 사람이
`★이지케어양식(함수)★ver2.xlsx`(3행 수당/공제 그룹, 4행 항목명, 5행부터
사원, 합계/공제액계/지급급여는 수식) 모양으로 다시 짜서 저장해 두었다
(`작업파일\\급여대장`). 이 모듈은 그 수정본과 같은 달 원본을 비교해서
"수정본의 각 열이 원본의 어느 시트/어느 머리글 열(같은 행, 몇 줄 아래, 또는
같은 사람 여러 행의 합)에서 왔는지"를 **금액이 사원 전원에 대해 정확히 같을
때만** 규칙으로 기록한다(`learn_recipe`). 다음 달 원본이 오면 그 규칙대로
수정본과 같은 양식의 표를 새로 만들어(`apply_recipe`) 그 표를 엔진에
넣는다 — 머리글 텍스트로 열을 다시 찾기 때문에 열이 몇 칸 밀려도 된다.

규칙을 정확히 재현할 수 없는 수정본(사람이 금액 자체를 손으로 고친 경우
등)은 `complete=False`로 남기고 쓰지 않는다 — 금액을 추측하지 않는다."""

from __future__ import annotations

import re
from dataclasses import asdict, dataclass, field
from pathlib import Path

import openpyxl

from core.canon import norm
from core.gridutil import cell_text
from core.sheet_layout import _is_plausible_employee_name, detect_source_layout
from core.xlsx_load import load_workbook, worksheets_only

_TOL = 1.0
_MAX_ROWS = 1500
_MAX_COLS = 300
_NAME_SCAN_COLS = 30
_TOTAL_LABELS = ("계", "합계", "총계", "소계", "총합계")
_REF = re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d+)")


def _nm(v: object) -> str:
    return re.sub(r"[\s·]", "", str(v)) if v is not None else ""


def _num(v: object) -> float:
    if isinstance(v, bool) or v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "")
    try:
        return float(s)
    except ValueError:
        return 0.0


_LATIN_NAME = re.compile(r"[A-Za-z][A-Za-z .,()\-가-힣]{1,40}")
_LATIN_LABELS = {"no", "name", "total", "sum"}


def _is_name(v: object) -> bool:
    """사원 이름 칸 — 한글 이름 + 외국인 영문 이름(업체067 CHISHUNJIN,
    업체068 EMPLOYEE_X2(사원013) 등 실제로 있음)."""
    if _is_plausible_employee_name(v):
        return True
    s = str(v).strip() if isinstance(v, str) else ""
    return bool(s) and bool(_LATIN_NAME.fullmatch(s)) and s.lower() not in _LATIN_LABELS


def _is_total_label(v: object) -> bool:
    k = norm(str(v)) if v is not None else ""
    return bool(k) and (k in _TOTAL_LABELS or k.startswith("합계") or k.startswith("총계"))


def _sheet_pattern(title: str) -> str:
    return re.sub(r"\d+", "#", title or "")


_VOLATILE = re.compile(r"\d{4}|\d+\s*월\s*분|\d+\s*월\s*\d+\s*일|생성일|근무기간|^[\d,.\-~%]+$")
_MAX_MISSING_COLUMNS = 2
_CHECK_WORDS = ("계", "합", "총", "지급액", "실지급", "실수령", "수령", "차인", "차감", "공제액")


def _header_signature(ws, data_row: int, col: int) -> str:
    """원본 열의 머리글(데이터 바로 위 몇 줄) — 달마다 바뀌는 제목/날짜 줄
    ("2026년08월 급여대장", "근무기간 20260801~20260831", "생성일 …", 숫자만
    있는 줄)은 빼고 데이터에 가까운 3단까지만 쓴다. 이걸로 다음 달 원본에서
    같은 항목 열을 다시 찾는다(열이 밀려도 된다)."""
    parts: list[str] = []
    for r in range(max(1, data_row - 6), data_row):
        t = cell_text(ws, r, col)
        if t is None or isinstance(t, (int, float)):
            continue
        raw = str(t)
        k = norm(raw)
        if not k or _VOLATILE.search(raw) or _VOLATILE.search(k):
            continue
        if not parts or parts[-1] != k:
            parts.append(k)
    return ">".join(parts[-3:])


@dataclass
class ColumnRule:
    col: int  # 출력(수정본 양식) 열 번호
    kind: str  # "src" | "formula" | "empty" | "unresolved"
    sheet: str = ""
    src_col: int = 0
    signature: str = ""
    offset: int = 0
    mode: str = "first"  # "first" | "sum"
    formula: str = ""  # "{r}" 자리에 행 번호
    label: str = ""
    # 원본 두 열의 합인 경우(예: 수정본 "기본급" = 원본 "기본급" + "주휴수당")
    src_col2: int = 0
    signature2: str = ""
    group: str = ""  # 수정본에서 이 열의 수당/공제 그룹
    total_kind: str = ""  # 합계 수식 열: "pay"(지급합계) | "ded"(공제액계) | "net"(지급급여)


@dataclass
class SheetRule:
    title: str
    name_col: int
    label: str = ""  # 이름 열의 머리글 글자(예: "성명")
    label_to_data: int = 0  # 머리글 행 -> 첫 사원 행 거리
    first_row: int = 0  # 학습한 달의 첫 사원 행(머리글을 못 찾을 때 대체)
    month_pos: int = -1  # 시트 이름 속 숫자 중 몇 번째가 귀속월인지("8월 (9)" -> 0)
    learned_month: int = 0  # 학습한 원본의 귀속월


@dataclass
class Recipe:
    template_sheet: str
    data_start_row: int
    marker_col: int
    name_col: int
    primary: str
    sheets: dict[str, SheetRule] = field(default_factory=dict)
    columns: list[ColumnRule] = field(default_factory=list)
    gap: int = 1  # 사원 행 사이에 허용하는 빈(이름 없는) 행 수
    exclude_names: list[str] = field(default_factory=list)  # 사람이 수정본에서 뺀 이름
    header_rows: list[int] = field(default_factory=list)  # 수정본 머리글 행
    # 원본에 금액이 있었는데 사람이 수정본에 넣지 않은 열(근로시간, 사업자부담 등)
    # — 다음 달에도 넣지 않는다. 여기 없고 규칙에도 없는 원본 금액 열은 새로
    # 생긴 항목으로 보고 수정본 양식 오른쪽에 덧붙인다(apply_recipe).
    ignored_signatures: list[str] = field(default_factory=list)
    # 검산용 원본 열 [(종류 net/pay/ded, 머리글 서명, 열)] — _learn_groups_and_ignored 참고
    checks: list = field(default_factory=list)
    drop_zero_rows: bool = False
    complete: bool = False
    problems: list[str] = field(default_factory=list)

    def to_json(self) -> dict:
        d = asdict(self)
        return d

    @staticmethod
    def from_json(d: dict) -> "Recipe":
        d = dict(d)
        d["sheets"] = {k: SheetRule(**v) for k, v in d.get("sheets", {}).items()}
        d["columns"] = [ColumnRule(**c) for c in d.get("columns", [])]
        return Recipe(**d)


# --- 수식 계산 -----------------------------------------------------------

def _col_index(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def _eval_formula(pattern: str, row: int, values: dict[int, float]) -> float | None:
    """수정본 데이터 행의 단순 수식(=SUM(C5:S5), =T5-AC5, =C5+D5 …)만 계산한다.
    다른 행을 가리키거나 함수가 섞이면 None."""
    expr = pattern.replace("{r}", str(row)).lstrip("=").upper()

    def sum_range(m: re.Match) -> str:
        a, b = m.group(1), m.group(2)
        ma, mb = _REF.fullmatch(a), _REF.fullmatch(b)
        if not ma or not mb or int(ma.group(4)) != row or int(mb.group(4)) != row:
            raise ValueError
        lo, hi = sorted((_col_index(ma.group(2)), _col_index(mb.group(2))))
        return repr(sum(values.get(c, 0.0) for c in range(lo, hi + 1)))

    try:
        expr = re.sub(r"SUM\(\s*([$A-Z0-9]+)\s*:\s*([$A-Z0-9]+)\s*\)", sum_range, expr)

        def ref(m: re.Match) -> str:
            if int(m.group(4)) != row:
                raise ValueError
            return repr(values.get(_col_index(m.group(2)), 0.0))

        expr = _REF.sub(ref, expr)
    except ValueError:
        return None
    if not re.fullmatch(r"[0-9.eE+\-*/() ]*", expr) or not expr.strip():
        return None
    try:
        return float(eval(expr, {"__builtins__": {}}, {}))  # noqa: S307 — 숫자/사칙연산만 남은 식
    except Exception:
        return None


# --- 원본 읽기 -----------------------------------------------------------

class _Sheet:
    def __init__(self, ws):
        self.ws = ws
        self.title = ws.title
        mr, mc = min(ws.max_row, _MAX_ROWS), min(ws.max_column, _MAX_COLS)
        self.rows = [[c.value for c in row] for row in ws.iter_rows(min_row=1, max_row=mr, max_col=mc)]
        self.max_row, self.max_col = mr, mc

    def v(self, r: int, c: int) -> object:
        if 1 <= r <= len(self.rows) and 1 <= c <= len(self.rows[r - 1]):
            return self.rows[r - 1][c - 1]
        return None

    def num(self, r: int, c: int) -> float:
        return _num(self.v(r, c))

    def name_index(self, col: int) -> dict[str, list[int]]:
        idx: dict[str, list[int]] = {}
        for r in range(1, len(self.rows) + 1):
            k = _nm(self.v(r, col))
            if k:
                idx.setdefault(k, []).append(r)
        return idx


def _load_sheets(path: Path) -> list[_Sheet]:
    wb = load_workbook(path, data_only=True)
    return [_Sheet(ws) for ws in worksheets_only(wb)]


def _find_label(sheet: _Sheet, col: int, first_row: int) -> tuple[str, int]:
    for r in range(first_row - 1, max(0, first_row - 12), -1):
        t = cell_text(sheet.ws, r, col)
        if t is None or isinstance(t, (int, float)):
            continue
        if not _is_plausible_employee_name(t) or norm(str(t)) in ("성명", "이름", "사원명", "종사자명", "직원명"):
            return str(t), first_row - r
    return "", 0


def _extract_rows(sheet: _Sheet, name_col: int, start: int, gap: int) -> list[int]:
    rows: list[int] = []
    blank = 0
    for r in range(start, len(sheet.rows) + 1):
        v = sheet.v(r, name_col)
        if any(_is_total_label(sheet.v(r, c)) for c in range(1, min(4, sheet.max_col) + 1)) or _is_total_label(v):
            break
        if _is_name(v):
            rows.append(r)
            blank = 0
            continue
        blank += 1
        if rows and blank > gap:
            break
    return rows


# --- 학습 ----------------------------------------------------------------

def _fixed_rows(fws, layout) -> list[tuple[int, str]]:
    out: list[tuple[int, str]] = []
    for r in range(layout.data_start_row, fws.max_row + 1):
        v = fws.cell(row=r, column=layout.name_col).value
        if _is_total_label(v) or any(_is_total_label(fws.cell(row=r, column=c).value) for c in range(1, 3)):
            break
        if _is_name(v):
            out.append((r, _nm(v)))
    return out


def learn_recipe(original_path: Path, fixed_path: Path) -> Recipe | None:
    from core.engine import _source_candidates

    cands = _source_candidates(fixed_path)
    if not cands:
        return None
    fws, fl = cands[0]
    if fl.block_height != 1 or getattr(fl, "needs_flatten", False) or not hasattr(fws, "merged_cells"):
        return None
    formulas_ws = openpyxl.load_workbook(fixed_path)[fws.title]
    frows = _fixed_rows(fws, fl)
    if not frows:
        return None
    names = [n for _, n in frows]
    name_set = set(names)

    sheets = _load_sheets(original_path)
    # (시트, 이름열) 후보 — 수정본 사원 60% 이상이 한 열에 나오는 곳
    cand_sheets: list[tuple[_Sheet, int, dict[str, list[int]]]] = []
    for sh in sheets:
        for c in range(1, min(sh.max_col, _NAME_SCAN_COLS) + 1):
            idx = {k: v for k, v in sh.name_index(c).items() if k in name_set}
            if len(idx) >= max(1, 0.6 * len(name_set)):
                cand_sheets.append((sh, c, idx))
    cand_sheets.sort(key=lambda t: -len(t[2]))

    recipe = Recipe(
        template_sheet=fws.title,
        data_start_row=fl.data_start_row,
        marker_col=fl.marker_col,
        name_col=fl.name_col,
        primary="",
    )
    used_sheets: dict[str, int] = {}
    for c in range(1, fws.max_column + 1):
        if c in (fl.name_col, fl.marker_col):
            continue
        label = " > ".join(
            str(t) for r in fl.header_rows if (t := cell_text(fws, r, c)) not in (None, "")
        )
        fvals = [_num(fws.cell(row=r, column=c).value) for r, _ in frows]
        f0 = formulas_ws.cell(row=frows[0][0], column=c).value
        if isinstance(f0, str) and f0.startswith("="):
            pat = _REF.sub(
                lambda m: f"{m.group(1)}{m.group(2)}{m.group(3)}{{r}}" if int(m.group(4)) == frows[0][0] else m.group(0),
                f0,
            )
            ok = all(
                (v := _eval_formula(pat, r, {cc: _num(fws.cell(row=r, column=cc).value) for cc in range(1, fws.max_column + 1)})) is not None
                and abs(v - fv) <= _TOL
                for (r, _), fv in zip(frows, fvals)
            )
            if ok:
                recipe.columns.append(ColumnRule(col=c, kind="formula", formula=pat, label=label))
                continue
        if not any(abs(v) > _TOL for v in fvals):
            recipe.columns.append(ColumnRule(col=c, kind="empty", label=label))
            continue
        found = _find_column(recipe, cand_sheets, frows, fvals, c, label)
        if not found:
            found = _find_pair_sum(cand_sheets[:3], frows, fvals, c, label)
            if found and found.sheet not in recipe.sheets:
                sh, ncol, idx = next(t for t in cand_sheets if f"{t[0].title}|{t[1]}" == found.sheet)
                first_row = min(r for rs in idx.values() for r in rs)
                lab, k = _find_label(sh, ncol, first_row)
                recipe.sheets[found.sheet] = SheetRule(sh.title, ncol, lab, k, first_row)
        if found:
            recipe.columns.append(found)
            used_sheets[found.sheet] = used_sheets.get(found.sheet, 0) + 1
        elif _is_optional_column(fws, fl, c):
            # 근로시간/근로일수, 사업자부담 보험료, 합계류 — 위하고 금액으로 가지
            # 않는 열이라 원본에서 못 찾아도 빈칸으로 둔다.
            recipe.columns.append(ColumnRule(col=c, kind="empty", label=label))
        else:
            recipe.columns.append(ColumnRule(col=c, kind="unresolved", label=label))
            recipe.problems.append(f"'{label}' 열을 원본에서 찾지 못함")

    if not used_sheets:
        recipe.problems.append("원본에서 금액 열을 하나도 찾지 못함")
        return recipe
    month = _month_of(original_path)
    for srule in recipe.sheets.values():
        nums = [int(d) for d in re.findall(r"\d+", srule.title)]
        if month is not None:
            hits = [i for i, d in enumerate(nums) if d == month or (d > 100000 and d % 100 == month)]
            if len(hits) == 1:
                srule.month_pos = hits[0]
            srule.learned_month = month
    recipe.primary = max(used_sheets, key=used_sheets.get)

    # 다음 달에 쓸 사원 행 규칙 — 이번 달 원본에 적용해서 수정본 사원과
    # 정확히 같아지는 가장 작은 빈 행 허용치를 고른다.
    rule = recipe.sheets[recipe.primary]
    prim = next(s for s in sheets if s.title == rule.title)
    want = list(dict.fromkeys(names))
    chosen = None
    for gap in (1, 2, 3, 5, 10, 30):
        for drop in (False, True):
            got = _rows_to_names(prim, rule, recipe, gap, drop, rule.first_row)
            if all(n in got for n in want):
                extra = [n for n in got if n not in want]
                chosen = (gap, drop, extra)
                break
        if chosen:
            break
    if chosen is None:
        recipe.problems.append("원본에서 수정본 사원을 모두 찾지 못함")
    elif len(chosen[2]) > max(len(want), 3):
        recipe.problems.append(f"원본 범위에 수정본에 없는 이름이 너무 많음({len(chosen[2])}개)")
    else:
        # 사람이 뺀 이름(직종 글자, "직접비", 강사 등)은 다음 달에도 뺀다.
        recipe.gap, recipe.drop_zero_rows, recipe.exclude_names = chosen
    _learn_groups_and_ignored(recipe, fws, fl, frows, prim)
    if not recipe.problems:
        # 같은 달 원본에 직접 돌려 보고, 검산 열마다 그달에 이미 안 맞던 사원
        # 수를 허용치로 둔다(원본 자체가 몇 명은 합계와 항목이 안 맞는 경우 —
        # 업체054). 사원의 20%를 넘게 안 맞는 열은 검산에 쓰지 않는다.
        recipe.complete = True
        measured = apply_recipe(recipe, original_path, fixed_path, _measure=True)
        n = len(set(names))
        limit = n // 5 if n >= 5 else 0
        if isinstance(measured, list):
            recipe.checks = [c[:3] + [m] for c, m in zip(recipe.checks, measured) if m <= limit]
        else:
            recipe.checks = []
        if not any(c[0] in ("net", "pay") for c in recipe.checks):
            recipe.problems.append("원본에서 검산할 지급 합계 열(실지급액/지급합계)을 찾지 못함")
    recipe.complete = not recipe.problems
    return recipe


def _total_kind(fws, fl, c: int, group: str) -> str:
    from core.sheet_layout import source_item

    k = norm(source_item(fws, fl, c))
    # 이지케어 표준양식의 "수당" 그룹 끝 합계 열처럼 항목명 칸이 비어 그룹
    # 이름만 있는 열 — 그 그룹의 합계다(수식 대신 값으로 붙여넣은 수정본도 있음).
    bottom = cell_text(fws, fl.header_rows[-1], c) if fl.header_rows else None
    top = cell_text(fws, fl.header_rows[0], c) if fl.header_rows else None
    if group in ("수당", "공제") and len(fl.header_rows) > 1 and (bottom in (None, "") or bottom == top):
        return "pay" if group == "수당" else "ded"
    if "공제" in k and any(w in k for w in ("계", "합", "총")):
        return "ded"
    if any(w in k for w in ("지급급여", "실지급", "실수령", "차인")):
        return "net"
    if group == "수당" and (not k or any(w in k for w in ("계", "합계", "지급액", "총액", "총지급"))):
        return "pay"
    return ""


def _learn_groups_and_ignored(recipe: Recipe, fws, fl, frows, prim: "_Sheet") -> None:
    from core.classify import group_of_src

    recipe.header_rows = list(fl.header_rows)
    for rule in recipe.columns:
        rule.group = group_of_src(fws, fl, rule.col)
        # 합계류 열(소득합계/공제합계/실수령액 …)은 수식이든 원본 복사든 표시만 —
        # 검산·덧붙인 항목 보정에서 수당/공제 항목으로 더하지 않는다.
        rule.total_kind = _total_kind(fws, fl, rule.col, rule.group)
    if not recipe.primary:
        return
    srule = recipe.sheets[recipe.primary]
    used = set()
    item_used = set()  # 합계류가 아닌 항목 열(검산 열 후보에서 뺀다)
    for rule in recipe.columns:
        if rule.kind == "src" and rule.sheet == recipe.primary:
            used.update({rule.src_col, rule.src_col2})
            if not rule.total_kind:
                item_used.update({rule.src_col, rule.src_col2})
    idx = prim.name_index(srule.name_col)
    rows = [idx[n][0] for _, n in frows if n in idx]
    for oc in range(1, prim.max_col + 1):
        if oc in used or oc == srule.name_col:
            continue
        if not any(isinstance(prim.v(r, oc), (int, float)) and abs(prim.num(r, oc)) > _TOL for r in rows):
            continue
        sig = _header_signature(prim.ws, srule.first_row, oc)
        if sig and sig not in recipe.ignored_signatures:
            recipe.ignored_signatures.append(sig)

    # 검산 열: 원본에서 수정본의 "수당 합 - 공제 합"(net) / "수당 합"(pay) /
    # "공제 합"(ded)과 사원별로 같은 열(보통 원본 자체의 실지급액/지급합계/
    # 공제합계). 다음 달에 변환한 표를 이 열들과 맞춰 보고, 하나라도 다르면
    # 열을 잘못 잡은 것이므로 변환을 쓰지 않는다(apply_recipe).
    def fixed_sum(r: int, kind: str) -> float:
        total = 0.0
        for rule in recipe.columns:
            if rule.kind == "src" and rule.group in ("수당", "공제") and not rule.total_kind:
                v = _num(fws.cell(row=r, column=rule.col).value)
                if rule.group == "수당" and kind in ("net", "pay"):
                    total += v
                elif rule.group == "공제" and kind in ("net", "ded"):
                    total += -v if kind == "net" else v
        return total

    recipe.checks = []
    for kind in ("net", "pay", "ded"):
        pairs = [(fixed_sum(r, kind), idx[n][0]) for r, n in frows if n in idx]
        if not pairs or sum(1 for want, _ in pairs if abs(want) > _TOL) < min(2, len(pairs)):
            continue
        allowed = len(pairs) // 5 if len(pairs) >= 5 else 0
        for oc in range(1, prim.max_col + 1):
            # 변환에 쓰는 항목 열은 검산 열이 될 수 없다(학습한 달에 수당이
            # 기본급뿐이면 "기본급" 열이 우연히 지급 합계와 같아진다). 머리글이
            # 합계처럼 보이는 열만 쓴다.
            if oc == srule.name_col or oc in item_used:
                continue
            bad = sum(1 for want, r in pairs if abs(prim.num(r, oc) - want) > _TOL)
            if bad <= allowed:
                sig = _header_signature(prim.ws, srule.first_row, oc)
                if sig and any(w in sig for w in _CHECK_WORDS):
                    recipe.checks.append([kind, sig, oc])
                    break


def _find_column(recipe: Recipe, cand_sheets, frows, fvals, c: int, label: str) -> "ColumnRule | None":
    """수정본 열 하나의 원본 출처. 사원 전원이 정확히 같은 열을 먼저 찾고,
    없으면 사원 80% 이상이 같은 열(나머지는 사람이 그 달 금액을 손으로 고친
    것 — 업체054: 3명 기본급 수정)을 받아들인다. 0이 아닌 값이
    적어도 2명(사원이 1~2명이면 전원) 맞아야 한다."""
    n = len(frows)
    allowed = n // 5 if n >= 5 else 0
    need_nonzero = min(2, sum(1 for v in fvals if abs(v) > _TOL))
    best: tuple[int, tuple] | None = None
    for sh, ncol, idx in cand_sheets:
        for mode in ("first", "sum"):
            for off in (0, 1, 2, -1):
                for oc in range(1, sh.max_col + 1):
                    if oc == ncol:
                        continue
                    bad = 0
                    hit = 0
                    for (_, nm_), fv in zip(frows, fvals):
                        rs = idx.get(nm_)
                        if not rs:
                            ov = 0.0
                        elif mode == "first":
                            ov = sh.num(rs[0] + off, oc)
                        else:
                            ov = sum(sh.num(r0 + off, oc) for r0 in rs)
                        if abs(ov - fv) > _TOL:
                            bad += 1
                            if bad > allowed:
                                break
                        elif abs(fv) > _TOL:
                            hit += 1
                    if bad > allowed or hit < need_nonzero:
                        continue
                    if best is None or bad < best[0]:
                        best = (bad, (sh, ncol, idx, oc, off, mode))
                    if bad == 0:
                        break
                if best and best[0] == 0:
                    break
            if best and best[0] == 0:
                break
        if best and best[0] == 0:
            break
    if best is None:
        return None
    sh, ncol, idx, oc, off, mode = best[1]
    first_row = min(r for rs in idx.values() for r in rs)
    tid = f"{sh.title}|{ncol}"
    if tid not in recipe.sheets:
        lab, k = _find_label(sh, ncol, first_row)
        recipe.sheets[tid] = SheetRule(sh.title, ncol, lab, k, first_row)
    return ColumnRule(
        col=c, kind="src", sheet=tid, src_col=oc, signature=_header_signature(sh.ws, first_row, oc),
        offset=off, mode=mode, label=label,
    )


def _is_optional_column(fws, fl, c: int) -> bool:
    from core.canon import canon
    from core.classify import group_of_src
    from core.sheet_layout import source_item

    if group_of_src(fws, fl, c) not in ("수당", "공제"):
        return True
    item = source_item(fws, fl, c)
    k = norm(item)
    return canon(item) in ("지급액계", "공제액계", "차인지급액") or any(
        w in k for w in ("합계", "소계", "총계", "공제계", "지급액", "실지급", "실수령")
    ) or k.endswith("계")


def _find_pair_sum(cand_sheets, frows, fvals, c: int, label: str) -> "ColumnRule | None":
    """수정본 열 = 원본 같은 행 두 열의 합인지(첫 번째 행만, 같은 줄)."""
    for sh, ncol, idx in cand_sheets:
        vecs: dict[int, tuple[float, ...]] = {}
        for oc in range(1, sh.max_col + 1):
            if oc == ncol:
                continue
            v = tuple(round(sh.num(idx[n][0], oc), 0) if n in idx else 0.0 for _, n in frows)
            if any(abs(x) > _TOL for x in v):
                vecs[oc] = v
        target = tuple(round(x, 0) for x in fvals)
        by_vec: dict[tuple[float, ...], list[int]] = {}
        for oc, v in vecs.items():
            by_vec.setdefault(v, []).append(oc)
        for oc, v in vecs.items():
            need = tuple(t - x for t, x in zip(target, v))
            for oc2 in by_vec.get(need, []):
                if oc2 > oc:
                    first_row = min(r for rs in idx.values() for r in rs)
                    return ColumnRule(
                        col=c, kind="src", sheet=f"{sh.title}|{ncol}", src_col=oc,
                        signature=_header_signature(sh.ws, first_row, oc), label=label,
                        src_col2=oc2, signature2=_header_signature(sh.ws, first_row, oc2),
                    )
    return None


def _rows_to_names(sheet: _Sheet, rule: SheetRule, recipe: Recipe, gap: int, drop: bool, start: int) -> list[str]:
    rows = _extract_rows(sheet, rule.name_col, start, gap)
    if drop:
        cols = [(c.src_col, c.offset) for c in recipe.columns if c.kind == "src" and recipe.sheets.get(c.sheet) is rule]
        rows = [r for r in rows if any(abs(sheet.num(r + off, oc)) > _TOL for oc, off in cols)]
    return list(dict.fromkeys(_nm(sheet.v(r, rule.name_col)) for r in rows))


# --- 적용 ----------------------------------------------------------------

def _month_of(path: Path) -> int | None:
    m = re.search(r"(\d{1,2})\s*월\s*귀속", path.stem)
    return int(m.group(1)) if m else None


def _month_in_title(title: str, pos: int) -> int | None:
    nums = [int(d) for d in re.findall(r"\d+", title)]
    if not (0 <= pos < len(nums)):
        return None
    d = nums[pos]
    return d % 100 if d > 100000 else d


def _resolve_sheet(sheets: list[_Sheet], srule: "SheetRule", month: int | None, sigs: list[str]) -> _Sheet | None:
    """학습한 시트를 이번 원본에서 찾는다: 같은 이름 -> 이름 모양이 같고
    (숫자만 다름) 귀속월 자리 숫자가 이번 달인 시트 -> 머리글이 가장 많이
    맞는 시트(학습한 머리글의 80% 이상)."""
    for s in sheets:
        if s.title == srule.title:
            return s
    pat = _sheet_pattern(srule.title)
    same = [s for s in sheets if _sheet_pattern(s.title) == pat]
    if len(same) == 1:
        return same[0]
    if same and month is not None and srule.month_pos >= 0:
        hit = [s for s in same if _month_in_title(s.title, srule.month_pos) == month]
        if len(hit) == 1:
            return hit[0]
        if len(hit) > 1 and srule.learned_month:
            # 같은 달 시트가 여러 개("7월 (8)"/"7월 (3)" — 연도별로 쌓인 통합문서)면
            # 학습한 시트 이름의 숫자들을 이번 달만큼 옮긴 값에 가장 가까운 것.
            delta = month - srule.learned_month
            base = [int(d) for d in re.findall(r"\d+", srule.title)]

            def dist(sh: _Sheet) -> int:
                nums = [int(d) for d in re.findall(r"\d+", sh.title)]
                return sum(abs(a - (b + delta)) for a, b in zip(nums, base) if b < 100)

            hit.sort(key=dist)
            if dist(hit[0]) < dist(hit[1]):
                return hit[0]
    want = [g for g in sigs if g]
    if not want:
        return None
    scored = []
    for s in sheets:
        have: set[str] = set()
        start = srule.first_row
        for r in range(max(2, start - 3), min(len(s.rows), start + 40) + 1):
            for c in range(1, s.max_col + 1):
                have.add(_header_signature(s.ws, r, c))
        scored.append((sum(1 for g in want if g in have), s))
    scored.sort(key=lambda t: -t[0])
    if scored and scored[0][0] >= 0.8 * len(want) and (len(scored) == 1 or scored[1][0] < scored[0][0]):
        return scored[0][1]
    return None


def _find_label_row(sheet: _Sheet, rule: SheetRule) -> tuple[int, int] | None:
    """(이름 열, 첫 사원 행) — 머리글 글자로 찾고, 못 찾으면 학습한 위치."""
    if rule.label:
        key = norm(rule.label)
        learned_label_row = rule.first_row - rule.label_to_data
        # 이름 열에서 먼저 찾고, 없으면(열이 밀린 경우) 다른 열에서 찾는다.
        # 한 시트에 표가 여러 개면(업체004: 위 지급표/아래 공제표) 같은
        # 머리글이 여러 번 나오므로 학습한 위치에서 가장 가까운 것을 쓴다.
        for cols in ([rule.name_col], range(1, min(sheet.max_col, _NAME_SCAN_COLS) + 1)):
            hits: list[tuple[int, int]] = []
            for c in cols:
                r = 1
                while r <= min(len(sheet.rows), 300):
                    t = cell_text(sheet.ws, r, c)
                    if t is not None and norm(str(t)) == key:
                        # 세로 병합 머리글("성 명"이 2~3행)이면 병합의 마지막 행 기준 —
                        # 학습 때(_find_label)는 아래에서 위로 찾아 마지막 행을 잡는다.
                        while (n := cell_text(sheet.ws, r + 1, c)) is not None and norm(str(n)) == key:
                            r += 1
                        hits.append((c, r))
                    r += 1
            if hits:
                c, r = min(hits, key=lambda h: (abs(h[1] - learned_label_row), h[0] != rule.name_col))
                return c, r + rule.label_to_data
    return rule.name_col, rule.first_row


LAST_APPLY_FAILURE = ""


def _fail(why: str):
    global LAST_APPLY_FAILURE
    LAST_APPLY_FAILURE = why
    return None


def apply_recipe(recipe: Recipe, original_path: Path, template_path: Path, _measure: bool = False):
    """원본을 수정본과 같은 양식의 워크시트로 바꾼다. 규칙대로 재현할 수
    없으면(시트/머리글 열을 못 찾음) None — 호출하는 쪽은 원래 방식으로
    처리한다. 반환: (워크시트, SourceLayout, 메모 목록)."""
    global LAST_APPLY_FAILURE
    LAST_APPLY_FAILURE = ""
    if not recipe.complete:
        return _fail("규칙 불완전")
    sheets = _load_sheets(original_path)
    month = _month_of(original_path)
    resolved: dict[str, _Sheet] = {}
    for tid, srule in recipe.sheets.items():
        sigs = [c.signature for c in recipe.columns if c.sheet == tid]
        sh = _resolve_sheet(sheets, srule, month, sigs)
        if sh is None:
            return _fail(f"시트 '{srule.title}'를 찾지 못함: {[x.title for x in sheets]}")
        resolved[tid] = sh

    prim_rule = recipe.sheets[recipe.primary]
    prim = resolved[recipe.primary]
    name_col, start = _find_label_row(prim, prim_rule)
    rows = _extract_rows(prim, name_col, start, recipe.gap)
    if not rows:
        return _fail(f"사원 행 없음(이름열 {name_col}, 시작 {start})")

    # 열 다시 찾기(머리글 서명) — 학습한 자리의 머리글이 같으면 그대로,
    # 다르면 같은 머리글을 가진 가장 가까운 열. 못 찾으면 변환하지 않는다.
    def relocate(sh: _Sheet, ref_row: int, col: int, sig: str) -> int | None:
        if _header_signature(sh.ws, ref_row, col) == sig:
            return col
        if not sig:
            return None
        hits = [c for c in range(1, sh.max_col + 1) if _header_signature(sh.ws, ref_row, c) == sig]
        return min(hits, key=lambda c: abs(c - col)) if hits else None

    src_cols: dict[int, tuple[int, int]] = {}
    missing: list[str] = []
    for rule in recipe.columns:
        if rule.kind != "src":
            continue
        sh = resolved[rule.sheet]
        srule = recipe.sheets[rule.sheet]
        if rule.sheet == recipe.primary:
            ref_row = start
        else:
            _, ref_row = _find_label_row(sh, srule)
        c1 = relocate(sh, ref_row, rule.src_col, rule.signature)
        c2 = relocate(sh, ref_row, rule.src_col2, rule.signature2) if rule.src_col2 else 0
        if c1 is None or c2 is None:
            # 그 달에만 있던 항목 열(업체037 "오지급공제", 업체069 "프로그램가산")은
            # 다른 달 원본에 아예 없다 — 몇 개까지는 빈칸으로 두고 메모를 남긴다.
            missing.append(rule.label or rule.signature)
            if len(missing) <= _MAX_MISSING_COLUMNS:
                continue
            return _fail(f"머리글 '{rule.signature}' 열을 찾지 못함(기준행 {ref_row}, 지금 {_header_signature(sh.ws, ref_row, rule.src_col)!r})")
        src_cols[rule.col] = (c1, c2)

    other_idx: dict[str, tuple[int, dict[str, list[int]]]] = {}
    for title, sh in resolved.items():
        if title == recipe.primary:
            continue
        nc, _ = _find_label_row(sh, recipe.sheets[title])
        other_idx[title] = (nc, sh.name_index(nc))

    by_name: dict[str, list[int]] = {}
    skip = set(recipe.exclude_names)
    for r in rows:
        k = _nm(prim.v(r, name_col))
        if k not in skip:
            by_name.setdefault(k, []).append(r)

    out_rows: list[tuple[str, dict[int, float]]] = []
    for name, rs in by_name.items():
        vals: dict[int, float] = {}
        for rule in recipe.columns:
            if rule.kind != "src" or rule.col not in src_cols:
                continue
            oc, oc2 = src_cols[rule.col]
            if rule.sheet == recipe.primary:
                use = rs
                sh = prim
            else:
                sh = resolved[rule.sheet]
                use = other_idx[rule.sheet][1].get(name, [])
            if not use:
                continue
            take = use[:1] if rule.mode == "first" else use
            vals[rule.col] = sum(
                sh.num(r + rule.offset, oc) + (sh.num(r + rule.offset, oc2) if oc2 else 0.0) for r in take
            )
        if recipe.drop_zero_rows and not any(abs(v) > _TOL for v in vals.values()):
            continue
        out_rows.append((name, vals))

    extras, unknown = _extra_columns(recipe, prim, name_col, start, by_name, src_cols)

    # 검산 — 변환한 표의 수당 합/공제 합/(수당-공제)가 원본의 해당 합계 열과
    # 사원 전원 같아야 쓴다. 학습한 달과 시트 구성이 달라 열을 잘못 잡았으면
    # 여기서 걸린다(업체070 8월: 열이 밀려 금액이 엉뚱하게 들어감).
    if not _measure and not any(k in ("net", "pay") for k, *_ in recipe.checks):
        return _fail("검산할 원본 지급 합계 열(실지급액/지급합계)이 없어 변환하지 않음")
    groups = {rule.col: rule.group for rule in recipe.columns if not rule.total_kind}

    def conv(name: str, vals: dict[int, float], kind: str) -> float:
        pay = sum(v for col, v in vals.items() if groups.get(col) == "수당")
        pay += sum(v.get(name, 0.0) for g, _, v in extras if g == "수당")
        ded = sum(v for col, v in vals.items() if groups.get(col) == "공제")
        ded += sum(v.get(name, 0.0) for g, _, v in extras if g == "공제")
        return {"net": pay - ded, "pay": pay, "ded": ded}[kind]

    measured: list[int] = []
    for check in recipe.checks:
        kind, sig, col = check[:3]
        allowed = check[3] if len(check) > 3 else 0
        cc = relocate(prim, start, col, sig)
        if cc is None:
            if _measure:
                measured.append(10**9)
                continue
            return _fail(f"검산 열 '{sig}'을 찾지 못함")
        bad = [n for n, vals in out_rows if abs(conv(n, vals, kind) - prim.num(by_name[n][0], cc)) > _TOL]
        if _measure:
            measured.append(len(bad))
        elif len(bad) > allowed:
            return _fail(f"검산 불일치({kind}, {len(bad)}명: {', '.join(bad[:5])}) — 원본 '{sig}'과 다름")
    if _measure:
        return measured

    wb = openpyxl.load_workbook(template_path)
    ws = wb[recipe.template_sheet]
    for rng in list(ws.merged_cells.ranges):
        if rng.max_row >= recipe.data_start_row:
            ws.unmerge_cells(str(rng))
    if ws.max_row >= recipe.data_start_row:
        ws.delete_rows(recipe.data_start_row, ws.max_row - recipe.data_start_row + 1)
    for i, (name, vals) in enumerate(out_rows):
        r = recipe.data_start_row + i
        if recipe.marker_col:
            ws.cell(row=r, column=recipe.marker_col, value=i + 1)
        ws.cell(row=r, column=recipe.name_col, value=name)
        for col, v in vals.items():
            ws.cell(row=r, column=col, value=v if abs(v) > 1e-9 else None)
        for rule in recipe.columns:
            if rule.kind == "formula":
                v = _eval_formula(rule.formula, r, {c: _num(ws.cell(row=r, column=c).value) for c in range(1, ws.max_column + 1)})
                ws.cell(row=r, column=rule.col, value=v)
    # 공식 열끼리 서로 참조할 수 있으니 한 번 더 계산
    for i in range(len(out_rows)):
        r = recipe.data_start_row + i
        for rule in recipe.columns:
            if rule.kind == "formula":
                v = _eval_formula(rule.formula, r, {c: _num(ws.cell(row=r, column=c).value) for c in range(1, ws.max_column + 1)})
                ws.cell(row=r, column=rule.col, value=v)
    # 새로 생긴 항목 열을 오른쪽에 붙이고 합계 열을 그만큼 고친다.
    if extras:
        top = recipe.header_rows[0] if recipe.header_rows else recipe.data_start_row - 2
        bottom = recipe.header_rows[-1] if recipe.header_rows else recipe.data_start_row - 1
        base_col = ws.max_column
        for j, (group, item, values) in enumerate(extras):
            col = base_col + 1 + j
            ws.cell(row=top, column=col, value=group)
            if bottom != top:
                ws.cell(row=bottom, column=col, value=item)
            for i, (name, _) in enumerate(out_rows):
                v = values.get(name, 0.0)
                ws.cell(row=recipe.data_start_row + i, column=col, value=v if abs(v) > 1e-9 else None)
        for i, (name, _) in enumerate(out_rows):
            r = recipe.data_start_row + i
            add_pay = sum(v.get(name, 0.0) for g, _, v in extras if g == "수당")
            add_ded = sum(v.get(name, 0.0) for g, _, v in extras if g == "공제")
            for rule in recipe.columns:
                delta = {"pay": add_pay, "ded": add_ded, "net": add_pay - add_ded}.get(rule.total_kind)
                if rule.kind in ("formula", "src") and delta:
                    cur = ws.cell(row=r, column=rule.col).value
                    if cur is not None:
                        ws.cell(row=r, column=rule.col, value=_num(cur) + delta)
    layout = detect_source_layout(ws)
    if layout is None:
        return _fail("변환한 표의 구조를 인식하지 못함")
    notes = [f"원본을 저장된 수정본 양식('{template_path.name}')으로 변환해서 처리했습니다(사원 {len(out_rows)}명)."]
    if missing:
        notes.append("원본에 다음 항목 열이 없어 비워두었습니다: " + ", ".join(m.replace(chr(10), " ") for m in missing))
    if extras:
        notes.append("수정본 양식에 없던 원본 항목을 덧붙였습니다: " + ", ".join(f"[{g}] {i}" for g, i, _ in extras))
    if unknown:
        notes.append("수당/공제를 판단하지 못해 넣지 않은 원본 금액 열: " + ", ".join(unknown) + " — 확인 필요")
    return ws, layout, notes


_SKIP_WORDS = ("시간", "일수", "시급", "단가", "횟수", "주민", "번호", "근무", "요율", "%", "율", "나이", "인원")


def _extra_columns(recipe: Recipe, prim: _Sheet, name_col: int, start: int, by_name, src_cols):
    """규칙에도, 학습 때 사람이 뺀 목록에도 없는데 이번 원본에 금액이 있는
    열 — 학습한 달엔 0원이라 수정본에 없던 항목(연차수당, 휴일근로수당,
    조정수당 …)이다. 버리면 돈이 빠지므로 수당/공제를 정해서 덧붙인다."""
    from core.canon import canon
    from core.classify import is_allowance_canon, is_deduction_canon

    used: dict[int, str] = {}
    for rule in recipe.columns:
        if rule.kind == "src" and rule.sheet == recipe.primary and rule.col in src_cols:
            for c in src_cols[rule.col]:
                if c:
                    used[c] = rule.group
    ignored = set(recipe.ignored_signatures)
    extras: list[tuple[str, str, dict[str, float]]] = []
    unknown: list[str] = []
    for oc in range(1, prim.max_col + 1):
        if oc in used or oc == name_col:
            continue
        values: dict[str, float] = {}
        for name, rs in by_name.items():
            v = prim.v(rs[0], oc)
            if isinstance(v, (int, float)) and not isinstance(v, bool) and abs(v) > _TOL:
                values[name] = float(v)
        if not values:
            continue
        sig = _header_signature(prim.ws, start, oc)
        if not sig or sig in ignored:
            continue
        item = sig.split(">")[-1]
        if any(w in sig for w in _SKIP_WORDS) or item.endswith("계") or any(
            w in item for w in ("합계", "총액", "지급액", "실지급", "실수령", "차인", "소계")
        ) or canon(item) in ("지급액계", "공제액계", "차인지급액"):
            continue
        ca = canon(item)
        if is_deduction_canon(ca) or "공제" in sig:
            group = "공제"
        elif is_allowance_canon(ca):
            group = "수당"
        else:
            left = [used[c] for c in sorted(used) if c < oc and used[c] in ("수당", "공제")]
            right = [used[c] for c in sorted(used) if c > oc and used[c] in ("수당", "공제")]
            group = left[-1] if left and right and left[-1] == right[0] else ""
        if not group:
            unknown.append(item)
            continue
        extras.append((group, item, values))
    return extras, unknown
