"""A QFrame that accepts Excel files dragged onto it, forwarding the local
paths to whoever connects to `files_dropped` — used by ui/main_window.py's
batch file queue so a user can drag one or several .xlsx files straight from
Finder instead of always going through the file-open dialog."""

from __future__ import annotations

from PySide6.QtCore import Signal
from PySide6.QtWidgets import QFrame

_ACCEPTED_SUFFIXES = (".xlsx", ".xlsm")


def _local_excel_paths(mime_data) -> list[str]:
    if not mime_data.hasUrls():
        return []
    paths = []
    for url in mime_data.urls():
        if url.isLocalFile():
            path = url.toLocalFile()
            if path.lower().endswith(_ACCEPTED_SUFFIXES):
                paths.append(path)
    return paths


class DropZone(QFrame):
    files_dropped = Signal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("dropzone")
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event) -> None:
        if _local_excel_paths(event.mimeData()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event) -> None:
        paths = _local_excel_paths(event.mimeData())
        if paths:
            self.files_dropped.emit(paths)
            event.acceptProposedAction()
        else:
            event.ignore()
