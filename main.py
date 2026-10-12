import sys
from pathlib import Path

from PySide6.QtGui import QFontDatabase, QIcon
from PySide6.QtWidgets import QApplication

from ui.main_window import MainWindow
from ui.style import STYLESHEET


def _app_base_dir() -> Path:
    """Directory that holds bundled `datas` (assets/, RELEASE_README.txt, ...).

    Frozen (PyInstaller --onedir) builds put `datas` inside `_internal/`,
    a sibling of the .exe, not next to the .exe itself — confirmed by
    inspecting an actual build's dist/ folder (PyInstaller 6.22). `sys.executable`
    is the .exe's own path, so `Path(sys.executable).parent` looks one level
    too shallow and silently finds nothing; `sys._MEIPASS` is the path
    PyInstaller itself sets to the real bundled-data directory (`_internal/`
    for onedir, or the temp extraction dir for onefile) and is the correct
    one to use here."""
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
    return Path(__file__).parent


def _load_bundled_fonts() -> None:
    """Registers NanumGothic from assets/fonts/ with Qt's font database so
    the UI renders it correctly even on another person's machine that
    doesn't have it installed as a system font — this app is handed off to
    other users (PyInstaller --onedir build, see main.spec), not just run
    on this dev machine. Resolves via `_app_base_dir()` when frozen (see
    that function's docstring — bundled `datas` end up in `_internal/`,
    not next to the .exe).

    Was Pretendard (4 static OTF weights) until 2026-08-07 — on this
    machine it rendered visibly broken (glyph corruption) once bundled and
    loaded via `addApplicationFont` instead of used as a plain installed
    system font. Most likely cause: those 4 files register under
    inconsistent family names (`Pretendard`, `Pretendard Medium`,
    `Pretendard SemiBold` — confirmed via `QFontDatabase.applicationFontFamilies`
    per weight), so Qt's style-matching for the stylesheet's `font-weight:
    600` can land on a different registered face than intended, and for
    the Medium/SemiBold weight it's landing on a face Qt then has to
    synthesize a style for — that synthesis is what was corrupting Hangul
    glyphs specifically. Switched to NanumGothic instead: 2 TTF weights
    (Regular/Bold) that both register cleanly under the single family name
    "NanumGothic", so there's no ambiguous weight to mismatch. `.ttf` here
    is deliberate too, not just an artifact of which file this font ships
    as — TrueType glyf outlines are the safer, more universally-supported
    choice for a bundled application font than OTF/CFF."""
    fonts_dir = _app_base_dir() / "assets" / "fonts"
    for font_file in list(fonts_dir.glob("*.otf")) + list(fonts_dir.glob("*.ttf")):
        QFontDatabase.addApplicationFont(str(font_file))


def main() -> None:
    app = QApplication(sys.argv)
    _load_bundled_fonts()
    icon_path = _app_base_dir() / "assets" / "app_icon.ico"
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))
    app.setStyleSheet(STYLESHEET)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
