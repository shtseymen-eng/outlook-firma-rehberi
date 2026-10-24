from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
from threading import Event

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QFileDialog, QMessageBox

from outlook_rehberi.models import (
    AccountRef,
    CancellationToken,
    Contact,
    FolderRef,
    ScanProgress,
    ScanResult,
    ScanStats,
)
from outlook_rehberi.ui.main_window import MainWindow


NOW = datetime(2026, 10, 24, 10, 0, tzinfo=timezone.utc)


def complete_result(*, complete: bool = True) -> ScanResult:
    return ScanResult(
        contacts=(Contact("contact@firma.com", "firma.com", company="Firma"),),
        stats=ScanStats(NOW, NOW, accounts=1, folders=2, messages=15, unique_contacts=1),
        complete=complete,
    )


class FakeAuthenticator:
    def __init__(self) -> None:
        self.values = (AccountRef("a1", "owner@firma.com", "Owner", "tenant"),)
        self.disconnected: list[str] = []

    def accounts(self):
        return self.values

    def add_account(self):
        return self.values[0]

    def disconnect(self, account_id: str) -> None:
        self.disconnected.append(account_id)


class FakeStore:
    def clear_directory(self) -> None:
        pass

    def import_legacy_json(self, path: Path):
        return object()


class FakeScanner:
    def __init__(self) -> None:
        self.release = Event()
        self.started = Event()
        owner = AccountRef("a1", "owner@firma.com", "Owner", "tenant")
        self.owner = owner
        self.folder = FolderRef("a1", "inbox", "Gelen", None, ("Gelen",), 0, False)

    def load_folder_trees(self, accounts, cancellation: CancellationToken):
        return {"a1": (self.folder,)}

    def run(self, selections, full_rebuild, cancellation, callbacks):
        self.started.set()
        callbacks.on_progress(ScanProgress("owner@firma.com", ("Gelen",), 7, 3))
        while not self.release.wait(0.005):
            if cancellation.is_cancelled:
                return complete_result(complete=False)
        return complete_result()


def make_window(qtbot) -> tuple[MainWindow, FakeScanner]:
    scanner = FakeScanner()
    window = MainWindow(FakeAuthenticator(), scanner, FakeStore(), auto_refresh=False)
    qtbot.addWidget(window)
    window.set_accounts((scanner.owner,), {"a1": (scanner.folder,)})
    window.folder_tree.topLevelItem(0).setCheckState(0, Qt.CheckState.Checked)
    return window, scanner


def test_accounts_and_read_only_permission_copy_are_visible(qtbot) -> None:
    window, scanner = make_window(qtbot)
    del scanner

    assert window.folder_tree.topLevelItem(0).text(0) == "Owner — owner@firma.com"
    copy = window.permission_label.text().casefold()
    assert "yalnızca okur" in copy
    assert "göndermez" in copy
    assert "silmez" in copy


def test_scan_runs_in_background_updates_progress_and_cancel_stops_it(qtbot) -> None:
    window, scanner = make_window(qtbot)

    window.start_scan()
    qtbot.waitUntil(scanner.started.is_set, timeout=1000)
    qtbot.waitUntil(lambda: not window.start_button.isEnabled(), timeout=1000)
    qtbot.waitUntil(lambda: window.progress_label.text().startswith("7 mesaj"), timeout=1000)
    assert window.cancel_button.isEnabled()

    window.cancel_scan()
    qtbot.waitUntil(lambda: window.start_button.isEnabled(), timeout=2000)
    assert window.save_button.isEnabled() is False
    assert "iptal" in window.status_label.text().casefold()


def test_completed_scan_shows_completion_before_native_save_as(monkeypatch, qtbot, tmp_path) -> None:
    window, scanner = make_window(qtbot)
    del scanner
    calls: list[tuple[str, str]] = []
    destination = tmp_path / "Benim Rehberim.xlsx"

    def information(parent, title, text):
        calls.append(("information", text))
        return QMessageBox.StandardButton.Ok

    def save_name(parent, title, suggested, file_filter):
        calls.append(("save", suggested))
        assert suggested.endswith(".xlsx")
        assert "Excel" in file_filter
        return str(destination), file_filter

    monkeypatch.setattr(QMessageBox, "information", information)
    monkeypatch.setattr(QFileDialog, "getSaveFileName", save_name)
    saved: list[Path] = []
    monkeypatch.setattr(window, "start_export", lambda path: saved.append(path))

    window.handle_scan_completed(complete_result())

    assert calls[0] == ("information", "Tarama tamamlandı")
    assert calls[1][0] == "save"
    assert saved == [destination]
    assert window.save_button.isEnabled()


def test_cancelled_save_as_can_be_reopened_with_explicit_button(monkeypatch, qtbot) -> None:
    window, scanner = make_window(qtbot)
    del scanner
    answers = iter([("", ""), ("/tmp/yeniden.xlsx", "Excel (*.xlsx)")])
    monkeypatch.setattr(QMessageBox, "information", lambda *args: QMessageBox.StandardButton.Ok)
    monkeypatch.setattr(QFileDialog, "getSaveFileName", lambda *args: next(answers))
    saved: list[Path] = []
    monkeypatch.setattr(window, "start_export", lambda path: saved.append(path))

    window.handle_scan_completed(complete_result())
    assert saved == []
    window.save_button.click()

    assert saved == [Path("/tmp/yeniden.xlsx")]


def test_incomplete_scan_does_not_open_save_as_and_shows_retry(monkeypatch, qtbot) -> None:
    window, scanner = make_window(qtbot)
    del scanner
    opened = []
    monkeypatch.setattr(QFileDialog, "getSaveFileName", lambda *args: opened.append(True))
    monkeypatch.setattr(QMessageBox, "warning", lambda *args: QMessageBox.StandardButton.Ok)

    window.handle_scan_completed(complete_result(complete=False))

    assert opened == []
    assert window.save_button.isEnabled() is False
    assert "yeniden" in window.status_label.text().casefold()


def test_successful_export_is_atomic_and_errors_are_displayed(monkeypatch, qtbot, tmp_path) -> None:
    window, scanner = make_window(qtbot)
    del scanner
    window.last_result = complete_result()
    destination = tmp_path / "rehber.xlsx"
    messages: list[str] = []
    monkeypatch.setattr(
        QMessageBox,
        "information",
        lambda parent, title, text: messages.append(text) or QMessageBox.StandardButton.Ok,
    )
    monkeypatch.setattr(
        QMessageBox,
        "critical",
        lambda parent, title, text: messages.append(text) or QMessageBox.StandardButton.Ok,
    )

    window.start_export(destination)
    qtbot.waitUntil(destination.exists, timeout=3000)
    qtbot.waitUntil(lambda: any("kaydedildi" in item.casefold() for item in messages), timeout=3000)
    assert not list(tmp_path.glob("*.tmp.xlsx"))

    bad_destination = tmp_path / "klasor"
    bad_destination.mkdir()
    window.start_export(bad_destination)
    qtbot.waitUntil(lambda: any("kaydedilemedi" in item.casefold() for item in messages), timeout=3000)
