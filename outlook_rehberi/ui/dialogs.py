from __future__ import annotations

from datetime import datetime
from pathlib import Path

from PySide6.QtWidgets import QFileDialog, QMessageBox, QWidget

from ..exporter import default_export_name


def show_scan_complete(parent: QWidget) -> None:
    QMessageBox.information(parent, "Tarama", "Tarama tamamlandı")


def choose_export_path(parent: QWidget) -> Path | None:
    suggested = Path.home() / default_export_name(datetime.now())
    filename, _ = QFileDialog.getSaveFileName(
        parent,
        "Excel Dosyasını Farklı Kaydet",
        str(suggested),
        "Excel Çalışma Kitabı (*.xlsx)",
    )
    if not filename:
        return None
    destination = Path(filename)
    if destination.suffix.casefold() != ".xlsx":
        destination = destination.with_suffix(".xlsx")
    return destination
