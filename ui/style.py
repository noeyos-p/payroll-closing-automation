"""App-wide look: quiet, monochrome, generous whitespace — no color beyond
grayscale (the review grid's functional red mismatch highlight is the one
deliberate exception, kept as-is since it's a warning signal, not decor).
Applied once via QApplication.setStyleSheet() in main.py.
"""

# Grayscale palette only. Named by role, not by hue, so nothing here ever
# tempts a future edit to sneak in a color.
BG = "#fafafa"  # window background
SURFACE = "#ffffff"  # card/panel background
BORDER = "#e2e2e2"  # hairline dividers
BORDER_STRONG = "#d0d0d0"  # input/button borders
TEXT = "#1c1c1c"  # primary text
TEXT_MUTED = "#787878"  # secondary/help text
INK = "#1c1c1c"  # primary button fill
INK_HOVER = "#333333"
INK_PRESSED = "#000000"
GHOST_HOVER = "#f0f0f0"
GHOST_PRESSED = "#e6e6e6"

STYLESHEET = f"""
* {{
    font-family: "NanumGothic", "Segoe UI", "Malgun Gothic", "맑은 고딕", "Apple SD Gothic Neo", sans-serif;
}}

QMainWindow, QWidget#central {{
    background: {BG};
}}

QLabel {{
    color: {TEXT};
}}

QLabel[role="title"] {{
    font-size: 15px;
    font-weight: 600;
    color: {TEXT};
}}

QLabel[role="subtitle"] {{
    font-size: 12px;
    color: {TEXT_MUTED};
}}

QLabel[role="field"] {{
    font-size: 12px;
    color: {TEXT_MUTED};
}}

QFrame#card {{
    background: {SURFACE};
    border: 1px solid {BORDER};
    border-radius: 10px;
}}

QFrame#dropzone {{
    background: {SURFACE};
    border: 1px dashed {BORDER_STRONG};
    border-radius: 8px;
}}

QLineEdit {{
    background: {SURFACE};
    border: 1px solid {BORDER_STRONG};
    border-radius: 6px;
    padding: 6px 10px;
    color: {TEXT};
    selection-background-color: {BORDER_STRONG};
}}

QLineEdit:focus {{
    border: 1px solid {TEXT_MUTED};
}}

QPushButton {{
    background: {SURFACE};
    border: 1px solid {BORDER_STRONG};
    border-radius: 6px;
    padding: 7px 16px;
    color: {TEXT};
}}

QPushButton:hover {{
    background: {GHOST_HOVER};
}}

QPushButton:pressed {{
    background: {GHOST_PRESSED};
}}

QPushButton:disabled {{
    color: {TEXT_MUTED};
    border-color: {BORDER};
}}

QPushButton[role="primary"] {{
    background: {INK};
    border: 1px solid {INK};
    color: {SURFACE};
    font-weight: 600;
    padding: 9px 18px;
}}

QPushButton[role="primary"]:hover {{
    background: {INK_HOVER};
    border-color: {INK_HOVER};
}}

QPushButton[role="primary"]:pressed {{
    background: {INK_PRESSED};
    border-color: {INK_PRESSED};
}}

QPushButton[role="ghost"] {{
    background: transparent;
    border: 1px solid transparent;
    color: {TEXT_MUTED};
}}

QPushButton[role="ghost"]:hover {{
    background: {GHOST_HOVER};
    color: {TEXT};
}}

QTableWidget {{
    background: {SURFACE};
    border: 1px solid {BORDER};
    border-radius: 8px;
    gridline-color: {BORDER};
}}

QStatusLabel, QLabel#status {{
    color: {TEXT_MUTED};
    font-size: 12px;
}}

QScrollBar:vertical, QScrollBar:horizontal {{
    background: transparent;
    border: none;
    margin: 0;
}}

QScrollBar::handle {{
    background: {BORDER_STRONG};
    border-radius: 4px;
}}

QScrollBar::add-line, QScrollBar::sub-line {{
    height: 0;
    width: 0;
}}
"""
