"""폴더 일괄 급여입력·마감 실행기.

기본은 읽기 전용이며 업체리스트는 결과 폴더의 복사본에만 기록한다.
  python run_payroll_closing.py --offline            위하고 없이 원본·변환 대조만
  python run_payroll_closing.py --only 업체011  위하고 읽기 전용(지급건/내보내기 확인)
  --apply        위하고 변경 허용(복사·업로드·등록·마감). 미관측 단계는 멈춘다.
  --write-list   바탕화면 원본 업체리스트에 직접 기록(백업 자동 생성)
"""
from __future__ import annotations

import argparse
import asyncio
from datetime import datetime
from pathlib import Path
import shutil
import sys

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

DEFAULT_FOLDER = Path(r"\\SERVER\share"
                      r"\2026-10월 급여대장(지급기준)\사원002(10월지급 급여입력)")
DEFAULT_LIST = Path(r"C:\Users\<사용자>\Desktop\업체리스트(10월지급).xlsx")


def _save_results(rows: list[dict], path: Path) -> None:
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "결과"
    headers = ["파일", "업체명", "사업자번호", "지급구분", "결과", "사유"]
    ws.append(headers)
    for row in rows:
        ws.append([row[h] for h in headers])
    for col, width in zip("ABCDEF", (48, 24, 14, 8, 10, 100)):
        ws.column_dimensions[col].width = width
    wb.save(path)


class _OfflineAdapter:
    """No WEHAGO: reports every payment as not read, so nothing proceeds."""
    def register_upload_plan(self, payment, plan):
        pass

    async def list_closed(self, payment):
        raise RuntimeError("오프라인 확인: 위하고 조회 안 함(원본·변환 대조까지 통과)")


async def main(args) -> int:
    from core.payroll_closing_run import outcome_rows, run_folder
    from core.payroll_company_list import read_company_list

    out = ROOT / "data" / "closing_runs" / datetime.now().strftime("%Y%m%d_%H%M%S")
    out.mkdir(parents=True, exist_ok=True)
    folder = Path(args.folder)
    if args.only:
        staged = out / "only"
        staged.mkdir()
        for path in folder.iterdir():
            if path.is_file() and args.only in path.name:
                shutil.copy2(path, staged / path.name)
        folder = staged
    list_path = Path(args.list)
    if not args.write_list:
        copy = out / list_path.name
        shutil.copy2(list_path, copy)
        list_path = copy

    if args.offline:
        adapter = _OfflineAdapter()
    else:
        from core import wehago_upload as wu
        from core.payroll_browser import WehagoPayrollAdapter
        from core.settings import load_settings
        settings = load_settings()
        main_page = await wu._get_main_page(300_000, settings.wehago_id, settings.wehago_password)
        companies = {r.business_number: r.company_name for r in read_company_list(list_path).rows
                     if r.business_number}
        adapter = WehagoPayrollAdapter(main_page, companies, out / "downloads", allow_mutation=args.apply,
                                       export_closed=args.verify_export)

    outcomes, update = await run_folder(folder, list_path, args.year, args.month, adapter,
                                        write_list=True, db_path=ROOT / "data" / "registry.db")
    if not args.offline:
        import json
        (out / "observations.json").write_text(
            json.dumps(adapter.observations, ensure_ascii=False, indent=1, default=str), encoding="utf-8")
        for number, page in adapter._pages.items():
            try:
                await page.screenshot(path=str(out / f"last_{number}.png"))
            except Exception:
                pass
        try:
            await adapter.main_page.screenshot(path=str(out / "last_main.png"))
        except Exception:
            pass
    rows = outcome_rows(outcomes)
    _save_results(rows, out / "결과표.xlsx")
    for row in rows:
        print(f"[{row['결과']}] {row['업체명'] or row['파일']} {row['지급구분']} — {row['사유']}")
    print(f"결과 폴더: {out}")
    print(f"업체리스트 기록: {list_path}" + (f" (백업 {update.backup_path})" if update else " (변경 없음)"))
    return 0


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--folder", default=str(DEFAULT_FOLDER))
    parser.add_argument("--list", default=str(DEFAULT_LIST))
    parser.add_argument("--year", type=int, default=2026, help="지급연도")
    parser.add_argument("--month", type=int, default=10, help="지급월")
    parser.add_argument("--only", help="파일명에 이 글자가 들어간 원본만 처리")
    parser.add_argument("--offline", action="store_true")
    parser.add_argument("--apply", action="store_true")
    parser.add_argument("--write-list", action="store_true")
    parser.add_argument("--verify-export", action="store_true", help="마감된 지급건도 내보내 금액 확인(읽기 전용)")
    sys.exit(asyncio.run(main(parser.parse_args())))
