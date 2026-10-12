"""한 업체 급여입력·마감 (2026-10-08 업체002/업체003에서 확인한 순서).

실행(작업 폴더 = 바탕화면\\　):
  "All2 - 복사본/.venv/Scripts/python.exe" "All2 - 급여입력 마감 자동화/run_one_company.py" "<원본 파일명 일부>" [--apply] [--close]

  인자 없음   : 읽기만(상태·대조)
  --apply     : 전월 복사·임시등록·업로드까지(대조 일치 시 마감 전 멈춤)
  --close     : 대조가 사원별·합계 모두 일치하면 '완료'(마감)까지
  --net-override 이름=금액: 원본 실지급액이 틀렸다고 사용자가 정한 사원만 그 금액으로 대조
                (예: 업체004 사원001=2306050, 2026-10-12). 이름은 원본에 정확히 한 번 있어야 한다
  --list-name 이름: 업체리스트 B열 이름이 파일명의 업체명과 다를 때(행 선택은 여전히 사업자번호로 확인)
  --verify-closed: 이미 마감된 지급건을 Ctrl+G로 재대조해 일치하면 업체리스트 복사본에 기록
  --reuse-copy: 대상 지급건이 이미 있어도, 전월 지급건과 인원·차인지급액이 같은(손대지 않은
                복사본) 미마감 건이면 이번 실행이 만든 것처럼 업로드한다(사용자 승인 2026-10-12)

규칙(CLAUDE.md): 이미 마감 → 건너뜀. 마감 전 기존 지급건 → 비교만(다르면 비고, 덮어쓰지 않음).
지급건 없음 → 전월 같은 종류 지급건이 하나면 복사 → Ctrl+G 기준 변환 → '사원코드 없음'은
임시등록(동명 사원이 명부에 있으면 '이름(재)(임시)') → 다시 Ctrl+G → 업로드 → 사원별 대조 → 마감.
업체리스트는 data/live_isak/업체리스트(10월지급)_복사본.xlsx 에만 기록한다: 마감 O, 급여자료입력
사원 목록에 '(중도퇴사)' 사원이 있으면 '중도'. 업로드 파일은 작업파일/Exel 에 저장한다.
"""
from __future__ import annotations

import argparse
import asyncio
import json
from dataclasses import replace
from datetime import datetime
from decimal import Decimal
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

FOLDER = Path(r"\\SERVER\share"
              r"\2026-10월 급여대장(지급기준)\사원002(10월지급 급여입력)")
LIST = Path(r"C:\Users\<사용자>\Desktop\업체리스트(10월지급).xlsx")
LIST_COPY = ROOT / "data" / "live_isak" / "업체리스트(10월지급)_복사본.xlsx"
# Per-kind results, so a company with 급여 and 급+상 files is O only when every
# kind in the folder is closed (업체005, 2026-10-12); partial is written "O,X".
CLOSURES = ROOT / "data" / "live_isak" / "closures.json"
# State of each payment right after this tool copied it: WEHAGO may recalculate
# while copying (업체006 2026-10-12: 61명 141,349,820 → 138,947,130),
# so "same as last month" alone cannot recognise an untouched copy.
COPIES = ROOT / "data" / "live_isak" / "copies.json"
UPLOAD_DIR = Path(r"C:\Users\<사용자>\Desktop\작업파일\Exel")  # user request 2026-10-12
CODE_KIND = {"1": "급여", "2": "급+상", "5": "추급"}
KIND_CODE = {v: k for k, v in CODE_KIND.items()}
KIND_SCREEN = {"1": "1.급여", "2": "2.급여+상여", "5": "5."}


class Stop(Exception):
    pass


def upload_path_for(name: str, accrual_month: int, kind: str) -> Path:
    """Same naming as the app's download (`260921_업체00708(급여).xlsx`);
    never overwrites a file already there."""
    stem = f"{datetime.now():%y%m%d}_{name}{accrual_month:02d}({kind})"
    target, n = UPLOAD_DIR / f"{stem}.xlsx", 2
    while target.exists():
        target, n = UPLOAD_DIR / f"{stem}_{n}.xlsx", n + 1
    return target


# 급여자료입력 left grid (focused via #Left_grid_line). Screen shows a leaver as
# "사원003(중도퇴사)" in blue (user screenshot 2026-10-12); rows carry da_retire.
PAYMENT_EMPLOYEES_JS = """() => { const g = window.Grids && window.Grids.getActiveGrid && window.Grids.getActiveGrid();
  if (!g) return null;
  return g.getDataSource().getJsonRows().map(r => ({cd_emp: r.cd_emp, nm_krname: r.nm_krname, da_retire: r.da_retire})); }"""


def record_closure(number: str, kind: str, kinds: list[str], closed: bool, reason: str = "",
                   midway: tuple[str, ...] = ()) -> list:
    """Remember this kind's outcome and rewrite the list copy for every kind of the
    company in the folder: all closed → O(중도), some → "O,X", none → blank;
    failure reasons go to 비고 (user rule 2026-10-12: 오류는 비고에 적는다)."""
    from core.payroll_company_list import ClosureResult, read_company_list, update_company_list
    store = json.loads(CLOSURES.read_text(encoding="utf-8")) if CLOSURES.exists() else {}
    store.setdefault(number, {})[kind] = {"closed": closed, "reason": reason, "midway": list(midway)}
    CLOSURES.write_text(json.dumps(store, ensure_ascii=False, indent=1), encoding="utf-8")
    done = store[number]
    results = []
    for k in kinds:
        rec = done.get(k)
        ok = bool(rec and rec["closed"])
        why = "" if ok else (rec.get("reason") if rec else "") or "아직 마감 안 함"
        results.append(ClosureResult(k, ok, why, tuple(rec.get("midway", ()) if rec else ())))
    update_company_list(read_company_list(LIST_COPY), {number: results})
    return results


def short_issues(issues, limit: int = 6) -> str:
    names = [i.note.split(":")[0].strip() for i in issues]
    more = f" 외 {len(names) - limit}명" if len(names) > limit else ""
    return f"{len(names)}명: " + ", ".join(names[:limit]) + more


def overridden_net(read, overrides: dict[str, Decimal]):
    """Wrap read_source_net: replace the stated net of user-corrected people."""
    def read_net(path):
        people, _ = read(path)
        for name, net in overrides.items():
            hits = [i for i, p in enumerate(people) if p.name == name]
            if len(hits) != 1:
                raise Stop(f"--net-override {name}: 원본에서 이 이름을 하나로 찾지 못함({len(hits)}건)")
            people[hits[0]] = replace(people[hits[0]], net=net, raw=None)
        return people, sum((p.net for p in people), Decimal(0))
    return read_net


def midway_leavers(rows: list[dict]) -> tuple[str, ...]:
    """Payment employees marked 중도퇴사 or carrying a 퇴사일 → "이름(YYYYMMDD)"."""
    out = []
    for r in rows:
        name, retire = str(r.get("nm_krname") or "").strip(), str(r.get("da_retire") or "").strip()
        if "중도퇴사" in name or retire:
            base = name.replace("(중도퇴사)", "").strip()
            out.append(f"{base}(퇴사 {retire})" if retire else f"{base}(중도퇴사)")
    return tuple(out)


async def main(args, report: dict) -> dict:
    from core import wehago_upload as wu
    from core.payroll_browser import (PayrollBrowserError, WehagoPayrollAdapter, employees_from_rows, is_closed,
                                      matching_payment_rows, payment_with_date, _read_roster_rows)
    from core.payroll_closing_run import discover_jobs, _key
    from core.payroll_company_list import ClosureResult, read_company_list, update_company_list
    from core.payroll_live import (LiveStepError, close_payment, copy_payment, export_ledger,
                                   register_employee, wait_gone)
    from core.payroll_prepare import prepare
    from core.payroll_workflow import PaymentIdentity, PayrollSnapshot, normalize_name, reconcile
    from core.registry import get_company_by_name
    from core.settings import load_settings
    from core.wehago_export import read_wehago_export

    out = ROOT / "data" / "live_runs" / datetime.now().strftime("%Y%m%d_%H%M%S")
    out.mkdir(parents=True, exist_ok=True)
    report["out"] = str(out)

    def log(*parts):
        line = " ".join(map(str, parts))
        report["steps"].append(line)
        print(line, flush=True)

    all_jobs, _ = discover_jobs(FOLDER, 2026, 10)
    all_jobs = [j for j in all_jobs if not j.path.name.startswith("~$")]
    jobs = [j for j in all_jobs if args.file in j.path.name]
    if len(jobs) != 1:
        raise Stop(f"원본 파일을 하나로 찾지 못함({len(jobs)}건)")
    job = jobs[0]
    from core.month_block import extract_month, load_rules
    block_rule = load_rules(job.company_name)  # annual cumulative ledgers (업체008, 2026-10-12)
    list_name = args.list_name or (block_rule or {}).get("list_name", job.company_name)
    rows = [r for r in read_company_list(LIST).rows if _key(r.company_name) == _key(list_name)]
    if len(rows) != 1 or not rows[0].business_number:
        raise Stop("업체리스트에서 업체를 하나로 찾지 못함")
    row = rows[0]
    number, name = row.business_number, row.company_name
    report.update(company=name, number=number, file=job.path.name)
    if row.completed:
        raise Stop("업체리스트 이미 O: 건너뜀")
    overrides = {}
    for item in args.net_override:
        who, _, amount = item.partition("=")
        if not who.strip() or not amount.strip().isdigit():
            raise Stop(f"--net-override 형식 오류: {item}")
        overrides[who.strip()] = Decimal(amount.strip())
    if overrides:
        report["net_override"] = {k: int(v) for k, v in overrides.items()}
    registered = (get_company_by_name(name, db_path=ROOT / "data" / "registry.db")
                  or get_company_by_name(job.company_name, db_path=ROOT / "data" / "registry.db"))
    company_id = registered.id if registered else None  # its saved column mappings apply

    source_path = job.path
    if block_rule:
        source_path, block_notes = extract_month(job.path, job.accrual_month, block_rule,
                                                 out / f"{job.company_name}_{job.accrual_month}월구간.xlsx")
        log("연간 시트에서 월 구간 추출", source_path.name, block_notes)
    settings = load_settings()
    main_page = await wu._get_main_page(300_000, settings.wehago_id, settings.wehago_password)
    adapter = WehagoPayrollAdapter(main_page, {number: name}, out / "dl", allow_mutation=args.apply)
    keys = ("fg_pay", "dt_pay", "num", "yn_close", "am_orpay")
    show = lambda rs: json.dumps([{k: r.get(k) for k in keys} for r in rs], ensure_ascii=False)
    opening = PaymentIdentity(number, job.accrual_year, job.accrual_month, job.payment_month, None, "급여")
    try:
        page = await adapter._company_page(opening)
    except PayrollBrowserError as exc:
        # The list may spell the name without the file's spaces (업체009, 2026-10-12);
        # the search term only finds candidates — the row is still chosen by 사업자번호.
        if "검색 결과가 없습니다" not in str(exc) or job.company_name == name:
            raise
        log("업체 검색어 재시도:", job.company_name)
        adapter.companies[number] = job.company_name
        page = await adapter._company_page(opening)

    prev_month = job.accrual_month - 1
    if prev_month < 1:
        raise Stop("1월 귀속은 전월(전년도) 복사를 아직 지원하지 않음")
    prev = await adapter._payment_rows(
        page, PaymentIdentity(number, job.accrual_year, prev_month, job.accrual_month, None, "급여"))
    log("전월 지급건", show(prev))
    await page.keyboard.press("Escape")
    await page.wait_for_timeout(800)
    if job.kind_assumed:
        if len(prev) != 1 or str(prev[0].get("fg_pay")) not in CODE_KIND:
            raise Stop("파일명에 지급구분이 없고 전월 지급건이 하나가 아님: 보류")
        kind = CODE_KIND[str(prev[0]["fg_pay"])]
    else:
        kind = job.kind
    code = KIND_CODE[kind]
    report["kind"] = kind
    siblings = [j for j in all_jobs if _key(j.company_name) == _key(job.company_name)]
    kinds = sorted({kind} | {j.kind for j in siblings if not j.kind_assumed},
                   key=lambda k: ("급여", "급+상", "추급").index(k))
    report["_list"] = (number, kind, kinds)
    payment = PaymentIdentity(number, job.accrual_year, job.accrual_month, job.payment_month, None, kind)

    async def target_rows(attempts: int = 1):
        for attempt in range(attempts):
            found = await adapter._payment_rows(page, payment)
            hits_ = matching_payment_rows(found, payment)
            if hits_ or attempt == attempts - 1:
                return found, hits_
            log("지급건 목록 재조회", show(found))
            await page.keyboard.press("Escape")
            await page.wait_for_timeout(2000)

    async def fresh_prepare(tag, dated):
        ref, file_name = await export_ledger(page, out, tag)
        log("Ctrl+G", file_name)
        read_net = None
        if overrides:
            from core.payroll_source_check import read_source_net
            read_net = overridden_net(read_source_net, overrides)
            log("원본 실지급액 정정(사용자 결정)", {k: int(v) for k, v in overrides.items()})
        prepared_ = prepare(source_path, dated, db_path=ROOT / "data" / "registry.db", read_net=read_net,
                            reference_path=ref, client_name=name, company_id=company_id)
        expect = (block_rule or {}).get("expect_codes", {})
        if expect:
            # Matched people carry WEHAGO's names (both 사원004), so check by the
            # extracted row's own net: that amount must sit on the expected code.
            from core.payroll_source_check import read_source_net
            stated = {p.name: p.net for p in read_source_net(source_path)[0]}
            for who, code in expect.items():
                got = [p.amounts.net for p in prepared_.original.people if p.employee_code == code]
                if who not in stated or got != [stated[who]]:
                    raise Stop(f"{who}: 규칙상 사원코드 {code}에 {stated.get(who)}원이어야 하는데 {got} — 업로드하지 않음")
            log("규칙 사원코드 확인", expect)
        return ref, prepared_

    async def finish(dated, prepared, result_name):
        """Closed payment: re-export and reconcile, check 중도퇴사자, record the list copy."""
        final, _ = await export_ledger(page, out, "final")
        final_export = read_wehago_export(final)
        final_result = reconcile(prepared.original,
                                 PayrollSnapshot(dated, final_export.people, final_export.totals, True, True))
        log("마감 후 재대조", "일치" if final_result.exact else "불일치", [i.note for i in final_result.issues][:8])
        if not final_result.exact:
            report["result"] = "close_unconfirmed"
            return
        await page.bring_to_front()
        await page.locator("#Left_grid_line").focus()
        rows = await page.evaluate(PAYMENT_EMPLOYEES_JS)
        if not isinstance(rows, list) or not rows:
            report["result"] = "close_unconfirmed"
            raise Stop("급여자료입력 사원 목록을 읽지 못해 중도퇴사 여부 미확인: 업체리스트 미기록")
        midway = midway_leavers(rows)
        log("중도퇴사자", list(midway) or "없음", f"(사원 {len(rows)}명 확인)")
        report["result"], report["midway"] = result_name, list(midway)
        results = record_closure(number, kind, kinds, True, "", midway)
        report["listed"] = True
        log("업체리스트 복사본 기록:", [(r.kind, "O" if r.closed else "X") for r in results],
            "| 중도" if midway else "")

    current, hits = await target_rows()
    log("대상 지급건", show(current), "일치", len(hits))
    if len(hits) > 1:
        raise Stop("같은 지급월·구분 지급건이 여러 개: 보류")
    existing = bool(hits)
    if existing and is_closed(current[hits[0]]):
        if not args.verify_closed:
            await page.keyboard.press("Escape")
            report["result"] = "skipped_closed"
            raise Stop("위하고에 이미 마감된 지급건: 건너뜀")
        dated = payment_with_date(payment, current[hits[0]])
        await adapter._open_row(page, current, hits[0])
        ref, prepared = await fresh_prepare("verify", dated)
        await finish(dated, prepared, "closed_verified")
        return report
    reused = False
    prev_same = [p for p in prev if str(p.get("fg_pay")) == code]
    copy_key = f"{number}|{job.accrual_year}-{job.accrual_month:02d}|{kind}"
    copies = json.loads(COPIES.read_text(encoding="utf-8")) if COPIES.exists() else {}
    if existing and args.reuse_copy and len(prev_same) == 1:
        target = current[hits[0]]
        as_prev = all(str(target.get(k)) == str(prev_same[0].get(k)) for k in ("fg_pay", "num", "am_orpay"))
        recorded = copies.get(copy_key)
        as_copied = bool(recorded) and all(str(target.get(k)) == str(recorded.get(k)) for k in ("dt_pay", "num", "am_orpay"))
        same = as_prev or as_copied
        log("복사본 재사용 확인", "전월과 같음" if as_prev else "복사 직후 기록과 같음" if as_copied else "전월·복사기록과 다름(손댄 지급건)")
        if same:
            existing, reused = False, True
    if not existing and not reused:
        await page.keyboard.press("Escape")
        await page.wait_for_timeout(800)
        if not args.apply:
            raise Stop("지급건 없음: --apply 시 전월 복사부터 진행")
        if len([p for p in prev if str(p.get("fg_pay")) == code]) != 1:
            raise Stop("복사할 전월 같은 종류 지급건이 하나가 아님: 보류")
        dt = await copy_payment(page, job.accrual_year, prev_month, job.accrual_month, code)
        log("전월 복사 완료, 위하고 제안 지급일", dt)
        current, hits = await target_rows()
        if len(hits) != 1 or str(current[hits[0]].get("dt_pay")) != dt or is_closed(current[hits[0]]):
            raise Stop("복사로 만든 지급건 확인 실패")
        copies[copy_key] = {k: current[hits[0]].get(k) for k in ("dt_pay", "num", "am_orpay")}
        COPIES.write_text(json.dumps(copies, ensure_ascii=False, indent=1), encoding="utf-8")
        log("복사 직후 기록", copies[copy_key])
    dated = payment_with_date(payment, current[hits[0]])
    await adapter._open_row(page, current, hits[0])

    ref, prepared = await fresh_prepare("ref1", dated)
    log("원본", len(prepared.original.people), "명", prepared.original.totals.net, "| 메모", prepared.notes)

    if existing:
        before = read_wehago_export(ref)
        result = reconcile(prepared.original, PayrollSnapshot(dated, before.people, before.totals, True))
        log("기존 지급건 대조", "일치" if result.exact else "불일치", [i.note for i in result.issues][:8])
        if not result.exact:
            report["result"] = "blocked_existing_mismatch"
            report["note"] = "이미 있던 지급건이 원본과 다름(덮어쓰지 않음) " + short_issues(result.issues)
            return report
    else:
        missing = [p for p in prepared.original.people if p.employee_code is None]
        if missing:
            roster = await adapter._roster_page(page)
            employees = employees_from_rows(await _read_roster_rows(roster), payment.join_date)
            names = {normalize_name(e.name) for e in employees}
            next_code = max(int(e.code) for e in employees if e.code.isdigit()) + 1
            for person in missing:
                base = normalize_name(person.name)
                same = [e for e in employees if normalize_name(e.name).split("(")[0] == base]
                new_name = f"{person.name}(재)(임시)" if same else f"{person.name}(임시)"
                if normalize_name(new_name) in names:
                    raise Stop(f"{new_name}: 같은 이름이 이미 명부에 있음(중복 등록 방지)")
                log("임시등록", new_name, "코드", next_code, "| 동명", [(e.code, e.excluded_reason) for e in same])
                try:
                    await register_employee(roster, str(next_code), new_name, payment.join_date)
                except LiveStepError:
                    await roster.screenshot(path=str(out / f"register_fail_{next_code}.png"))
                    raise
                next_code += 1
            employees = await adapter.list_employees(dated)
            log("등록 확인", [(e.code, e.name) for e in employees if "(임시)" in e.name])
            await page.bring_to_front()
            await page.reload()
            await page.wait_for_load_state()
            await page.wait_for_timeout(3000)
            await wu._dismiss_stray_settings_screen(page)
            await wu._dismiss_overlays(page, rounds=6, settle_ms=500, min_empty_rounds=2)
            await wu._dismiss_known_popups(page)
            current, hits = await target_rows()
            await adapter._open_row(page, current, hits[0])
            ref, prepared = await fresh_prepare("ref2", dated)
            if any(p.employee_code is None for p in prepared.original.people):
                raise Stop("등록 후에도 사원코드 없음이 남음")
        upload_path = prepared.plan.write(upload_path_for(name, job.accrual_month, kind), {})
        log("업로드 파일", upload_path, len(prepared.plan.records), "행")
        await wu._upload_excel(page, upload_path, adapter.ui_lock)
        await wu._confirm_and_convert(page)
        try:
            await wait_gone(page, "재계산 반영중", 120)
        except LiveStepError:
            pass
        current, hits = await target_rows()
        log("저장 후", show(current))
        await adapter._open_row(page, current, hits[0])

    after, _ = await export_ledger(page, out, "after")
    export = read_wehago_export(after)
    result = reconcile(prepared.original, PayrollSnapshot(dated, export.people, export.totals, True))
    log("최종 대조", "일치" if result.exact else "불일치", "| 위하고", len(export.people), "명", export.totals.net,
        [i.note for i in result.issues][:8])
    if not result.exact:
        report["result"] = "blocked_mismatch"
        report["note"] = "업로드 후 위하고 금액이 원본과 다름 " + short_issues(result.issues)
        return report
    if not args.close:
        report["result"] = "ready_to_close"
        return report
    await close_payment(page, adapter.companies[number], job.accrual_year, job.accrual_month, dated.payment_date, KIND_SCREEN[code])
    current, hits = await target_rows(attempts=3)
    if len(hits) != 1:
        report["result"] = "close_unconfirmed"
        raise Stop("마감 후 지급건 목록에서 대상을 다시 찾지 못함: --verify-closed로 재확인 필요")
    log("마감 후", show([current[hits[0]]]))
    if not is_closed(current[hits[0]]):
        report["result"] = "close_unconfirmed"
        return report
    await adapter._open_row(page, current, hits[0])
    await finish(dated, prepared, "closed")
    return report


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")  # Korean log lines without a PYTHONIOENCODING prefix
    sys.stderr.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser()
    parser.add_argument("file")
    parser.add_argument("--apply", action="store_true")
    parser.add_argument("--close", action="store_true")
    parser.add_argument("--reuse-copy", action="store_true")
    parser.add_argument("--verify-closed", action="store_true")
    parser.add_argument("--net-override", action="append", default=[])
    parser.add_argument("--list-name")  # 업체리스트 이름이 파일명과 다를 때(업체010 → …기관)
    args = parser.parse_args()
    report: dict = {"steps": []}
    try:
        asyncio.run(main(args, report))
    except Stop as exc:
        report.setdefault("result", "stopped")
        report["note"] = str(exc)
    except Exception as exc:  # LiveStepError and anything unexpected: stop, nothing guessed
        report["result"] = "error"
        report["note"] = f"{type(exc).__name__}: {exc}"
    # Failed real attempts are written to 비고 (user rule 2026-10-12); read-only
    # runs and companies already closed in WEHAGO are left alone.
    if args.apply and "_list" in report and not report.get("listed") and report.get("result") not in (
            "closed", "closed_verified", "skipped_closed", "ready_to_close"):
        number, kind, kinds = report["_list"]
        reason = (str(report.get("note", "")).splitlines() or [""])[0][:150]  # no Playwright call logs in 비고
        try:
            record_closure(number, kind, kinds, False, reason)
            report["list_note"] = reason
        except Exception as exc:  # the run's own result is still printed
            report["list_note_error"] = f"{type(exc).__name__}: {exc}"
    report.pop("_list", None)
    print("RESULT", json.dumps({k: v for k, v in report.items() if k != "steps"}, ensure_ascii=False))
