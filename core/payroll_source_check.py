"""Read the ledger's stated net amounts independently of learned conversion rules."""
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from pathlib import Path
import re

from core.engine import _source_candidates, _pick_candidate
from core.profile import parse_template_token, payroll_type_from_sheet_name
from core.reconcile import _find_src_canon_col, _continuation_rows
from core.sheet_layout import find_total_row, last_nonblank_row


@dataclass(frozen=True)
class SourceNet:
    name: str
    net: Decimal
    source: str
    row: int
    # Set when the ledger value had sub-won digits hidden by an integer display
    # format; net is then the displayed (Excel-rounded) value. User rule 2026-10-08.
    raw: Decimal | None = None


def strict_won(value) -> Decimal:
    if value is None or isinstance(value, bool):
        raise ValueError('금액이 비어 있거나 숫자가 아닙니다.')
    text = str(value).strip().replace(',', '').replace('원', '')
    if text.startswith('(') and text.endswith(')'):
        text = '-' + text[1:-1]
    try:
        amount = Decimal(text)
    except InvalidOperation as exc:
        raise ValueError('금액을 읽지 못했습니다.') from exc
    if not amount.is_finite() or amount != amount.to_integral_value():
        raise ValueError('원 단위 정수 금액이 아닙니다.')
    return amount


# Strip colour/condition brackets, quoted text, escaped chars (\x), padding (_x),
# fill (*x), sign/paren literals and spaces; what remains must be digits only.
_FORMAT_NOISE = re.compile(r'\[[^\]]*\]|"[^"]*"|\\.|_.|\*.|[()+\- ]')


def _integer_only_format(number_format: str | None) -> bool:
    sections = (number_format or 'General').split(';')[:2]
    return all(re.fullmatch(r'[#,0]+', _FORMAT_NOISE.sub('', s)) for s in sections)


def displayed_won(value, number_format: str | None) -> tuple[Decimal, Decimal | None]:
    """Whole won as Excel shows it (user rule 2026-10-08). Returns (won, raw);
    raw is set only when sub-won digits were hidden by an integer-only format
    and rounded half up like Excel's display. A format that shows decimals
    (or General) is refused: the visible value itself is not whole won."""
    if isinstance(value, float) and not value.is_integer():
        if not _integer_only_format(number_format):
            raise ValueError('원 미만 금액이 소수로 표시되는 셀입니다.')
        raw = Decimal(repr(value))
        return raw.quantize(Decimal(1), rounding=ROUND_HALF_UP), raw
    return strict_won(value), None


def read_source_net(path: Path, sheet: str | None = None) -> tuple[list[SourceNet], Decimal]:
    candidates = _source_candidates(path)
    if not candidates:
        raise ValueError('원본 급여대장 구조를 읽지 못했습니다.')
    category = parse_template_token(sheet)
    if category is not None:
        candidates = [(ws, layout) for ws, layout in candidates
                      if payroll_type_from_sheet_name(ws.title) == category]
        if len(candidates) != 1:
            raise ValueError('수정 양식과 대응하는 원본 지급구분 시트를 확정하지 못했습니다.')
        ws, layout = candidates[0]
    else:
        ws, layout = _pick_candidate(candidates, sheet)
    column = _find_src_canon_col(ws, layout, '차인지급액')
    if column is None:
        stacked = _stacked_net_table(path, ws, layout)
        if stacked is not None:
            return stacked
        raise ValueError('원본 실지급액(차인지급액) 열이 없습니다.')
    last = last_nonblank_row(ws)
    total_row = find_total_row(ws, layout.data_start_row, last)
    end = total_row or last + 1
    continuation = _continuation_rows(ws, layout, end)
    continuation_rows = {r for rows in continuation.values() for r in rows}
    people = []
    for row in range(layout.data_start_row, end, layout.block_height):
        if row in continuation_rows:
            continue
        name = ws.cell(row, layout.name_col).value
        if not isinstance(name, str) or not name.strip():
            continue
        if name.startswith('#'):
            raise ValueError('원본 사원명에 수식 오류가 있습니다.')
        physical = list(range(row, row + layout.block_height)) + list(continuation.get(row, ()))
        cells = [ws.cell(r, column) for r in physical if ws.cell(r, column).value is not None]
        if not cells:
            raise ValueError(f'{name}: 원본 실지급액이 비어 있습니다.')
        parsed = [displayed_won(c.value, getattr(c, 'number_format', None)) for c in cells]  # flattened Grid cells carry no format
        amount = sum((won for won, _ in parsed), Decimal(0))
        rounded = any(raw is not None for _, raw in parsed)
        raw_total = sum(((raw if raw is not None else won) for won, raw in parsed), Decimal(0))
        people.append(SourceNet(name.strip(), amount, str(path), row, raw_total if rounded else None))
    if not people:
        raise ValueError('원본 사원별 실지급액이 없습니다.')
    total = sum((p.net for p in people), Decimal(0))
    if total_row is not None:
        printed = ws.cell(total_row, column).value
        if printed is not None:
            raw_sum = sum(((p.raw if p.raw is not None else p.net) for p in people), Decimal(0))
            if isinstance(printed, float) and not printed.is_integer():
                # Compare unrounded values: rounding each person can drift from
                # rounding the printed grand total by a won or two.
                if abs(Decimal(repr(printed)) - raw_sum) >= Decimal('0.01'):
                    raise ValueError('원본 사원별 실지급액 합계와 원본 합계행이 다릅니다.')
            elif strict_won(printed) != raw_sum:
                raise ValueError('원본 사원별 실지급액 합계와 원본 합계행이 다릅니다.')
    return people, total


_TOTAL_NAMES = ('계', '합계', '총계', '총합계')


def _stacked_net_table(path: Path, ws, layout) -> tuple[list[SourceNet], Decimal] | None:
    """A second table below the 수당 table with its own 성명 … 실지급액 header
    (업체004, 2026-10-12: '4대 보험공제 내역' table). Used only when its
    names are exactly the main table's names, so an unrelated table is refused."""
    from core.canon import canon, norm
    from core.sheet_layout import _SRC_NAME_LABELS
    last = last_nonblank_row(ws)
    for header in range(layout.data_start_row, last + 1):
        name_col = net_col = None
        for c in range(1, ws.max_column + 1):
            value = ws.cell(header, c).value
            if not isinstance(value, str):
                continue
            if name_col is None and norm(value) in _SRC_NAME_LABELS:
                name_col = c
            elif canon(value) == '차인지급액':
                net_col = c
        if name_col is not None and net_col is not None:
            break
    else:
        return None
    main_names = sorted(
        v.strip() for r in range(layout.data_start_row, header)
        if isinstance(v := ws.cell(r, layout.name_col).value, str) and v.strip()
        and norm(v) not in _TOTAL_NAMES)
    people, total_row = [], None
    for row in range(header + 1, last + 1):
        name = ws.cell(row, name_col).value
        if isinstance(name, str) and (norm(name) in _TOTAL_NAMES or norm(name).startswith('합계')):
            total_row = row
            break
        if not isinstance(name, str) or not name.strip():
            continue
        cell = ws.cell(row, net_col)
        if cell.value is None:
            raise ValueError(f'{name}: 원본 실지급액이 비어 있습니다.')
        won, raw = displayed_won(cell.value, cell.number_format)
        people.append(SourceNet(name.strip(), won, str(path), row, raw))
    if not people or sorted(p.name for p in people) != main_names:
        raise ValueError('원본 아래쪽 실지급액 표의 사원이 수당 표와 다릅니다.')
    total = sum((p.net for p in people), Decimal(0))
    if total_row is not None and ws.cell(total_row, net_col).value is not None:
        if strict_won(ws.cell(total_row, net_col).value) != sum(((p.raw or p.net) for p in people), Decimal(0)):
            raise ValueError('원본 사원별 실지급액 합계와 원본 합계행이 다릅니다.')
    return people, total


def payment_period(path: Path, year: int, payment_month: int) -> tuple[int, int]:
    accrual = re.search(r'(?<!\d)(\d{1,2})\s*월\s*귀속', path.stem)
    payment = re.search(r'(?<!\d)(\d{1,2})\s*월\s*지급', path.stem)
    if not accrual or not payment:
        raise ValueError('파일명에서 귀속월/지급월을 확정하지 못했습니다.')
    am, pm = int(accrual[1]), int(payment[1])
    if not 1 <= am <= 12 or pm != payment_month:
        raise ValueError('선택한 지급월과 원본 파일명이 다릅니다.')
    explicit = re.search(r'(?<!\d)(20\d{2})\s*년', path.stem)
    if explicit and int(explicit[1]) != year:
        raise ValueError('선택한 연도와 원본 연도가 다릅니다.')
    return (year - 1 if am > pm else year), am
