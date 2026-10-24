from __future__ import annotations

import os
from pathlib import Path
import sys

from PySide6.QtWidgets import QApplication, QMessageBox

from .auth import MicrosoftAuthenticator
from .config import AppConfig, ConfigurationError
from .graph import GraphClient
from .scanner import ScanService
from .storage import LocalStore
from .ui.main_window import MainWindow


def create_application(config: AppConfig) -> QApplication:
    application = QApplication.instance() or QApplication(sys.argv)
    application.setApplicationName("Outlook Firma Rehberi")
    application.setApplicationDisplayName("Outlook Firma Rehberi")
    application.setOrganizationName("S.SEYMEN")
    application.setProperty("outlook_rehberi_config", config)
    return application


def run_application() -> int:
    try:
        config = AppConfig.load(os.environ)
    except ConfigurationError as exc:
        application = QApplication.instance() or QApplication(sys.argv)
        QMessageBox.critical(None, "Yapılandırma Eksik", str(exc))
        return 2

    application = create_application(config)
    authenticator = MicrosoftAuthenticator(config)
    store = LocalStore()
    graph = GraphClient(authenticator, base_url=config.graph_base_url)
    scanner = ScanService(graph, store)
    _import_legacy_if_present(store)
    window = MainWindow(authenticator, scanner, store)
    window.show()
    return application.exec()


def _import_legacy_if_present(store: LocalStore) -> None:
    candidates = (
        Path.cwd() / "rehber_data.json",
        Path(sys.executable).resolve().parent / "rehber_data.json",
    )
    for candidate in candidates:
        if candidate.is_file():
            try:
                store.import_legacy_json(candidate)
            except (OSError, ValueError):
                pass
            return
