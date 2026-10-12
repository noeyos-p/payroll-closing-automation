"""WEHAGO print → PDF via the native "Duzon - PrintDialog" (WehagoPrint2.exe).

Ported 2026-10-12 from `　\\신고서 저장 자동화 - 복사본\\core\\wehago_print.py`
(observed 2026-09-23). The print window is a separate Windows program, so it is
driven with pywinauto, not Playwright:
  1) the page's print button (#print) is clicked by the caller (Playwright);
     the context needs Chrome's local-network-access permission, otherwise the
     page falls back to the download center (granted per context, 2026-10-12);
  2) wait for "Duzon - PrintDialog";
  3) click its btnSavePDF (auto_id, no coordinates);
  4) in the "다른 이름으로 저장" dialog (#32770) put the full path into the Edit
     and press Enter — clicking "저장(&S)" did nothing when observed;
  5) the file must start with %PDF- and end with %%EOF.
Functions are synchronous; call them from async code with asyncio.to_thread.
"""
from __future__ import annotations

import time
from pathlib import Path

PRINT_DIALOG_TITLE = "Duzon - PrintDialog"
SAVE_PDF_AUTO_ID = "btnSavePDF"


class PrintPdfError(RuntimeError):
    pass


def window_handles() -> set[int]:
    from pywinauto import findwindows
    return set(findwindows.find_windows())


def find_window(title_contains: str, class_name: str | None = None, timeout: float = 20,
                exclude_handles=(), process_id=None):
    from pywinauto import Desktop, findwindows
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        for handle in findwindows.find_windows():
            if handle in exclude_handles:
                continue
            try:
                w = Desktop(backend="win32").window(handle=handle).wrapper_object()
                if process_id is not None and w.process_id() != process_id:
                    continue
                if title_contains in w.window_text() and (class_name is None or w.element_info.class_name == class_name):
                    return w
            except Exception:
                continue
        time.sleep(0.5)
    return None


def print_dialog_contents(print_win) -> str:
    """Selected 인쇄형태 (cbContents) of an open print dialog."""
    from pywinauto import Application
    win = Application(backend="win32").connect(handle=print_win.handle).window(handle=print_win.handle)
    combos = [c for c in win.descendants() if c.friendly_class_name() == "ComboBox"
              and "연말정산근로소득원천징수영수증_전체" in c.texts()]
    return combos[0].selected_text() if len(combos) == 1 else ""


def save_open_dialog_as_pdf(print_win, dest: Path, close_after: bool = True) -> Path:
    """Save the already-open print dialog as `dest` (never overwrites). The print
    window is closed only after the saved file is complete (user rule 2026-10-12);
    on any failure it is left open so nothing is cut off."""
    from pywinauto import Application
    dest = Path(dest)
    if dest.exists():
        raise PrintPdfError(f"같은 이름의 파일이 이미 있습니다: {dest}")
    dest.parent.mkdir(parents=True, exist_ok=True)
    try:
        print_win.set_focus()  # without this the PDF click was ignored once (2026-10-12)
    except Exception:
        pass  # SetForegroundWindow can refuse; the click below still works when visible
    win = Application(backend="uia").connect(handle=print_win.handle).window(handle=print_win.handle)
    win.child_window(auto_id=SAVE_PDF_AUTO_ID).click_input()
    save_dlg = find_window("다른 이름으로 저장", class_name="#32770", timeout=15, process_id=print_win.process_id())
    if save_dlg is None:
        raise PrintPdfError("'다른 이름으로 저장' 창이 15초 안에 뜨지 않았습니다.")
    dlg = Application(backend="win32").connect(handle=save_dlg.handle).window(handle=save_dlg.handle)
    edit = dlg.child_window(class_name="Edit")
    edit.set_edit_text(str(dest))
    edit.click_input()
    edit.type_keys("{ENTER}")
    # WehagoPrint2 rewrites the file page by page: it briefly disappears and a
    # 1-page PDF already has %%EOF (업체008, 2026-10-12, cut at 1 of 9 pages when
    # the dialog was closed early). Wait until a valid file stops changing.
    deadline = time.monotonic() + 120
    last, stable = None, 0
    while time.monotonic() < deadline:
        time.sleep(1)
        try:
            size = dest.stat().st_size
            with dest.open("rb") as stream:
                head = stream.read(5)
                stream.seek(max(0, size - 2048))
                tail = stream.read()
        except OSError:
            last, stable = None, 0
            continue
        valid = size > 0 and head == b"%PDF-" and b"%%EOF" in tail
        stable = stable + 1 if valid and size == last else 0
        last = size
        if stable >= 4:
            if close_after:
                try:
                    print_win.close()
                except Exception:
                    pass
            return dest
    raise PrintPdfError(f"저장 후에도 완성된 PDF가 생기지 않았습니다(인쇄창은 열어 둠): {dest}")
