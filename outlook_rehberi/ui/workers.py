from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QObject, Signal, Slot

from ..exporter import build_workbook, save_workbook_atomic
from ..models import CancellationToken, ScanResult
from ..scanner import ScanCallbacks


class ScanWorker(QObject):
    progress = Signal(object)
    warning = Signal(object)
    completed = Signal(object)
    failed = Signal(str)
    cancelled = Signal(object)
    finished = Signal()

    def __init__(
        self,
        scanner,
        selections,
        full_rebuild: bool,
        cancellation: CancellationToken,
    ) -> None:
        super().__init__()
        self._scanner = scanner
        self._selections = selections
        self._full_rebuild = full_rebuild
        self._cancellation = cancellation

    @Slot()
    def run(self) -> None:
        try:
            result = self._scanner.run(
                self._selections,
                self._full_rebuild,
                self._cancellation,
                ScanCallbacks(self.progress.emit, self.warning.emit),
            )
            if self._cancellation.is_cancelled:
                self.cancelled.emit(result)
            else:
                self.completed.emit(result)
        except Exception:
            self.failed.emit("Tarama başlatılamadı. Bağlantınızı ve hesap iznini kontrol edin.")
        finally:
            self.finished.emit()


class FolderLoadWorker(QObject):
    completed = Signal(object)
    failed = Signal(str)
    finished = Signal()

    def __init__(self, scanner, accounts, cancellation: CancellationToken) -> None:
        super().__init__()
        self._scanner = scanner
        self._accounts = accounts
        self._cancellation = cancellation

    @Slot()
    def run(self) -> None:
        try:
            self.completed.emit(
                self._scanner.load_folder_trees(self._accounts, self._cancellation)
            )
        except Exception:
            self.failed.emit("Posta klasörleri alınamadı. Hesap bağlantısını yenileyin.")
        finally:
            self.finished.emit()


class ExportWorker(QObject):
    completed = Signal(str)
    failed = Signal(str)
    finished = Signal()

    def __init__(self, result: ScanResult, destination: Path) -> None:
        super().__init__()
        self._result = result
        self._destination = destination

    @Slot()
    def run(self) -> None:
        try:
            workbook = build_workbook(self._result.contacts, self._result)
            path = save_workbook_atomic(workbook, self._destination)
            self.completed.emit(str(path))
        except Exception:
            self.failed.emit("Excel dosyası kaydedilemedi. Konumu ve dosya adını kontrol edin.")
        finally:
            self.finished.emit()
