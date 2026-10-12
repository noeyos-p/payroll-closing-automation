"""Name / birth-digit employee matching, ported from
파일/이지케어_Sheet1_현재시트_수정본_v3.bas's `NameKey`/`FullNameKey`/
`NameBase3`/`DigitsOnly`/`BirthDigitsFromName`/`PersonBirthKey`/`BirthYY`/
`BirthDigitsMatch`/`ResolveSrcRow`/`ShouldAddSourceRow`.

Two additions layered on top of the ported VBA algorithm (kept because
they're real, previously-confirmed fixes from this app's sister project,
wehago-payroll-tool-py, and are NOT fully subsumed by VBA's own
exact-name -> base-3-char -> birth-digit cascade — see each function's
docstring for exactly which VBA case does/doesn't already cover it):

  - A source row named "이름1/이름2" (two people sharing one row the same
    pay period) is indexed under BOTH halves, not just read as one literal
    string — VBA's own NameBase3 already resolves ONE common case here (the
    registered name happens to be the first 2-3 characters of the compound
    string) purely by accident of Korean given names being short, but not
    the reversed-order case.
  - A destination name's trailing parenthetical content is also tried as an
    alternate lookup key — VBA's FullNameKey strips the parentheses but
    keeps the content glued onto the base text (so "사원007" still
    resolves via NameBase3 accidentally), but when the parenthetical
    content is itself a real alternate name for that person (not a title —
    confirmed real case: 업체047's roster lists
    "사원008" but the source ledger only ever calls her "사원X1"),
    the base-3 fallback can't bridge two completely different names."""

from __future__ import annotations

import re

_TAG_SUBSTRINGS = ("(재)", "（재）", "[재]", "【재】", "(임시)", "（임시）", "[임시]", "【임시】")
_PAREN_CHARS = "()（）[]【】"


def name_key(value: object) -> str:
    s = str(value if value is not None else "")
    for ch in ("\r", "\n", "\t", " ", "　", ","):
        s = s.replace(ch, "")
    return s.strip()


def full_name_key(value: object) -> str:
    s = name_key(value)
    for tag in _TAG_SUBSTRINGS:
        s = s.replace(tag, "")
    for ch in _PAREN_CHARS:
        s = s.replace(ch, "")
    return s.strip()


def name_base3(value: object) -> str:
    s = full_name_key(value)
    result = "".join(ch for ch in s if not ("0" <= ch <= "9") and ch not in ("년", "생"))
    result = result.strip()
    return result[:3] if len(result) >= 3 else result


def digits_only(value: object) -> str:
    return "".join(ch for ch in str(value if value is not None else "") if ch.isdigit())


def birth_digits_from_name(value: object) -> str:
    return digits_only(full_name_key(value))


def person_birth_key(name: object, birth_cell_value: object = None) -> str:
    digits = birth_digits_from_name(name)
    if not digits and birth_cell_value is not None:
        digits = digits_only(birth_cell_value)
    return digits


def _birth_yy(digits: str) -> str:
    if len(digits) >= 4 and digits[:2] in ("19", "20"):
        return digits[2:4]
    return digits[:2] if len(digits) >= 2 else ""


def birth_digits_match(a: str, b: str) -> bool:
    if not a or not b:
        return False
    if a == b:
        return True
    if len(a) >= len(b) and a.startswith(b):
        return True
    if len(b) >= len(a) and b.startswith(a):
        return True
    ay, by = _birth_yy(a), _birth_yy(b)
    return bool(ay) and bool(by) and ay == by


def _parenthetical_content(value: object) -> str | None:
    match = re.search(r"\(([^)]*)\)\s*$", str(value or ""))
    if match and match.group(1).strip():
        return match.group(1).strip()
    return None


def split_compound_name(raw_name: object) -> list[str]:
    """"이름1/이름2" -> ["이름1", "이름2"]; a plain name -> [name]."""
    return [p.strip() for p in str(raw_name or "").split("/") if p.strip()]


class SourceNameIndex:
    """Indexes source rows by exact full-name-key and by base-3-char key —
    port of VBA's `BuildSourceIndexes` (`srcExact`/`srcBase` Dictionaries),
    extended to also index each half of a "이름1/이름2" compound name."""

    def __init__(self) -> None:
        self.exact: dict[str, list[int]] = {}
        self.base: dict[str, list[int]] = {}
        self._row_names: dict[int, str] = {}

    def add(self, row: int, raw_name: object) -> None:
        names = split_compound_name(raw_name) or [str(raw_name or "")]
        self._row_names[row] = names[0]
        for name in names:
            full_key = full_name_key(name)
            base_key = name_base3(name)
            if not base_key:
                continue
            self.exact.setdefault(full_key, []).append(row)
            self.base.setdefault(base_key, []).append(row)

    def source_name(self, row: int) -> str:
        return self._row_names.get(row, "")


def resolve_source_row(
    dest_name: object,
    dest_birth: str,
    index: SourceNameIndex,
    used_rows: set[int],
    birth_lookup,
) -> tuple[int | None, str]:
    """Port of VBA `ResolveSrcRow`. `birth_lookup(row)` returns that source
    row's own birth-digit key (core/reconcile.py supplies this from the
    source layout's birth column, or the name-embedded digits)."""
    for candidate_name in [dest_name] + ([_parenthetical_content(dest_name)] if _parenthetical_content(dest_name) else []):
        full_key = full_name_key(candidate_name)
        base_key = name_base3(candidate_name)

        if full_key and full_key in index.exact:
            candidates = [r for r in index.exact[full_key] if r not in used_rows]
            if len(candidates) == 1:
                return candidates[0], ""
            if len(candidates) > 1:
                row = _match_by_birth(candidates, dest_birth, birth_lookup)
                if row is not None:
                    return row, "동명이인 중 생년월일 숫자로 매칭했습니다."
                return None, "같은 이름이 여러 명입니다. 생년월일 숫자로도 1명을 특정하지 못했습니다."

        if base_key and base_key in index.base:
            candidates = [r for r in index.base[base_key] if r not in used_rows]
            if len(candidates) == 1:
                row = candidates[0]
                memo = (
                    ""
                    if full_name_key(index.source_name(row)) == full_key
                    else "정확한 이름은 다르지만 앞 3글자가 같은 1명으로 자동 매칭했습니다."
                )
                return row, memo
            if len(candidates) > 1:
                row = _match_by_birth(candidates, dest_birth, birth_lookup)
                if row is not None:
                    return row, "앞 3글자가 같은 동명이인 중 생년월일 숫자로 매칭했습니다."
                return None, "앞 3글자가 같은 사람이 여러 명입니다. 생년월일 숫자로도 1명을 특정하지 못했습니다."

    return None, ""


def _match_by_birth(candidates: list[int], dest_birth: str, birth_lookup) -> int | None:
    if not dest_birth:
        return None
    hits = [row for row in candidates if birth_digits_match(dest_birth, birth_lookup(row))]
    return hits[0] if len(hits) == 1 else None


def should_add_source_row(
    src_name: object,
    src_birth: str,
    dest_base_index: dict[str, list[int]],
    dest_birth_lookup,
) -> bool:
    """Port of VBA `ShouldAddSourceRow` — should a source-only employee
    become a brand-new destination row, or be left as "확인필요" because a
    same-base3 destination row already plausibly represents them?"""
    base_key = name_base3(src_name)
    if not base_key:
        return False
    if base_key not in dest_base_index:
        return True
    if not src_birth:
        return False
    for row in dest_base_index[base_key]:
        if birth_digits_match(src_birth, dest_birth_lookup(row)):
            return False
    return True
