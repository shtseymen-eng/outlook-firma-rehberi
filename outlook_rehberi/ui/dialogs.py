from __future__ import annotations

from datetime import datetime
from pathlib import Path

from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QLineEdit,
    QMessageBox,
    QWidget,
)

from ..exporter import default_export_name
from ..models import Contact


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


def edit_contact(parent: QWidget, contact: Contact) -> tuple[str, str, str] | None:
    dialog = QDialog(parent)
    dialog.setWindowTitle(f"Kişiyi Düzenle — {contact.email}")
    layout = QFormLayout(dialog)
    name = QLineEdit(contact.name)
    title = QLineEdit(contact.title)
    company = QLineEdit(contact.company)
    layout.addRow("Ad Soyad", name)
    layout.addRow("Görev / Unvan", title)
    layout.addRow("Firma", company)
    buttons = QDialogButtonBox(
        QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel
    )
    buttons.accepted.connect(dialog.accept)
    buttons.rejected.connect(dialog.reject)
    layout.addRow(buttons)
    if dialog.exec() != QDialog.DialogCode.Accepted:
        return None
    return name.text().strip(), title.text().strip(), company.text().strip()
