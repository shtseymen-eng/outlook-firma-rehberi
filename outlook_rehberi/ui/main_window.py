from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QThread, QTimer, Qt
from PySide6.QtGui import QCloseEvent
from PySide6.QtWidgets import (
    QCheckBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from ..models import AccountRef, CancellationToken, ScanProgress, ScanResult, ScanWarning
from .dialogs import choose_export_path, show_scan_complete
from .folder_tree import FolderTreeWidget, SelectionRequiredError
from .workers import ExportWorker, FolderLoadWorker, ScanWorker


class MainWindow(QMainWindow):
    def __init__(
        self,
        authenticator,
        scanner,
        store,
        *,
        auto_refresh: bool = True,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.authenticator = authenticator
        self.scanner = scanner
        self.store = store
        self.last_result: ScanResult | None = None
        self._accounts: tuple[AccountRef, ...] = ()
        self._scan_cancellation: CancellationToken | None = None
        self._folder_cancellation: CancellationToken | None = None
        self._threads: set[QThread] = set()
        self._workers: set[object] = set()

        self.setWindowTitle("Outlook Firma Rehberi")
        self.resize(860, 680)
        self._build_ui()
        if auto_refresh:
            self.refresh_accounts()

    def _build_ui(self) -> None:
        central = QWidget(self)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(12)

        title = QLabel("Outlook Firma Rehberi")
        title_font = title.font()
        title_font.setPointSize(title_font.pointSize() + 7)
        title_font.setBold(True)
        title.setFont(title_font)
        layout.addWidget(title)

        self.permission_label = QLabel(
            "Microsoft hesabınızdaki postaları yalnızca okur. "
            "E-posta göndermez, silmez, taşımaz veya değiştirmez."
        )
        self.permission_label.setWordWrap(True)
        layout.addWidget(self.permission_label)

        account_actions = QHBoxLayout()
        self.add_account_button = QPushButton("Hesap Ekle")
        self.disconnect_button = QPushButton("Seçili Hesabı Ayır")
        self.refresh_button = QPushButton("Klasörleri Yenile")
        account_actions.addWidget(self.add_account_button)
        account_actions.addWidget(self.disconnect_button)
        account_actions.addWidget(self.refresh_button)
        account_actions.addStretch(1)
        layout.addLayout(account_actions)

        self.folder_tree = FolderTreeWidget()
        layout.addWidget(self.folder_tree, 1)

        options = QHBoxLayout()
        self.full_rebuild_box = QCheckBox("Tüm kayıtları baştan oluştur")
        self.clear_button = QPushButton("Yerel Rehberi Temizle")
        self.import_button = QPushButton("Eski v4 Verisini İçe Aktar")
        options.addWidget(self.full_rebuild_box)
        options.addStretch(1)
        options.addWidget(self.import_button)
        options.addWidget(self.clear_button)
        layout.addLayout(options)

        self.progress_label = QLabel("Henüz tarama yapılmadı")
        self.status_label = QLabel("Hazır")
        self.status_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        layout.addWidget(self.progress_label)
        layout.addWidget(self.status_label)

        scan_actions = QHBoxLayout()
        self.start_button = QPushButton("Taramayı Başlat")
        self.cancel_button = QPushButton("İptal Et")
        self.save_button = QPushButton("Excel'i Farklı Kaydet")
        self.cancel_button.setEnabled(False)
        self.save_button.setEnabled(False)
        scan_actions.addWidget(self.start_button)
        scan_actions.addWidget(self.cancel_button)
        scan_actions.addStretch(1)
        scan_actions.addWidget(self.save_button)
        layout.addLayout(scan_actions)

        self.setCentralWidget(central)
        self.add_account_button.clicked.connect(self.add_account)
        self.disconnect_button.clicked.connect(self.disconnect_selected_account)
        self.refresh_button.clicked.connect(self.refresh_accounts)
        self.start_button.clicked.connect(self.start_scan)
        self.cancel_button.clicked.connect(self.cancel_scan)
        self.save_button.clicked.connect(self.prompt_save_as)
        self.clear_button.clicked.connect(self.clear_local_directory)
        self.import_button.clicked.connect(self.import_legacy_data)

    def set_accounts(self, accounts, trees) -> None:
        self._accounts = tuple(accounts)
        aliases = {
            account.account_id: self.store.aliases(account.account_id)
            for account in self._accounts
            if hasattr(self.store, "aliases")
        }
        self.folder_tree.set_own_addresses(aliases)
        self.folder_tree.set_data(self._accounts, trees)
        self.status_label.setText(
            f"{len(self._accounts)} Microsoft hesabı hazır. Klasörleri seçin."
        )

    def refresh_accounts(self) -> None:
        try:
            accounts = tuple(self.authenticator.accounts())
        except Exception:
            self.folder_tree.set_error("Bağlı hesaplar okunamadı")
            return
        self._accounts = accounts
        if not accounts:
            self.folder_tree.set_error("Hesap bulunamadı. “Hesap Ekle” ile başlayın.")
            self.status_label.setText("Microsoft hesabı ekleyin.")
            return
        self.folder_tree.set_loading()
        self._folder_cancellation = CancellationToken()
        worker = FolderLoadWorker(self.scanner, accounts, self._folder_cancellation)
        self._start_worker(worker)
        worker.completed.connect(lambda trees: self.set_accounts(accounts, trees))
        worker.failed.connect(self._folder_load_failed)

    def _folder_load_failed(self, message: str) -> None:
        self.folder_tree.set_error(message)
        self.status_label.setText(message)

    def add_account(self) -> None:
        try:
            self.authenticator.add_account()
        except Exception:
            QMessageBox.critical(
                self, "Hesap Eklenemedi", "Microsoft hesabı eklenemedi. Yeniden deneyin."
            )
            return
        self.refresh_accounts()

    def disconnect_selected_account(self) -> None:
        account = self.folder_tree.selected_account()
        if account is None:
            QMessageBox.warning(self, "Hesap Seçin", "Ayırmak için bir hesap seçin.")
            return
        try:
            self.authenticator.disconnect(account.account_id)
        except Exception:
            QMessageBox.critical(self, "Hata", "Microsoft hesabı ayrılamadı.")
            return
        self.refresh_accounts()

    def start_scan(self) -> None:
        try:
            selections = self.folder_tree.selections()
        except SelectionRequiredError as exc:
            QMessageBox.warning(self, "Klasör Seçin", str(exc))
            return
        self.last_result = None
        self.save_button.setEnabled(False)
        self._scan_cancellation = CancellationToken()
        worker = ScanWorker(
            self.scanner,
            selections,
            self.full_rebuild_box.isChecked(),
            self._scan_cancellation,
        )
        self._start_worker(worker)
        worker.progress.connect(self._scan_progress)
        worker.warning.connect(self._scan_warning)
        worker.completed.connect(self.handle_scan_completed)
        worker.cancelled.connect(self._scan_cancelled)
        worker.failed.connect(self._scan_failed)
        self._set_busy(True)
        self.status_label.setText("Tarama sürüyor…")

    def cancel_scan(self) -> None:
        if self._scan_cancellation is not None:
            self._scan_cancellation.cancel()
            self.cancel_button.setEnabled(False)
            self.status_label.setText("Tarama iptal ediliyor…")

    def _scan_progress(self, progress: ScanProgress) -> None:
        folder = " › ".join(progress.folder_path)
        self.progress_label.setText(
            f"{progress.processed_messages} mesaj • "
            f"{progress.unique_contacts} tekil e-posta • "
            f"{progress.account_email} • {folder}"
        )

    def _scan_warning(self, warning: ScanWarning) -> None:
        self.status_label.setText(warning.message)

    def handle_scan_completed(self, result: ScanResult) -> None:
        self._set_busy(False)
        if not result.complete:
            self.last_result = None
            self.save_button.setEnabled(False)
            self.status_label.setText("Tarama tamamlanamadı. Sorunlu klasörleri yeniden deneyin.")
            QMessageBox.warning(
                self,
                "Tarama Tamamlanamadı",
                "Tarama eksik kaldı. Excel kaydı açılmadı; yeniden deneyin.",
            )
            return
        self.last_result = result
        self.save_button.setEnabled(True)
        self.status_label.setText("Tarama tamamlandı")
        show_scan_complete(self)
        self.prompt_save_as()

    def _scan_cancelled(self, result: ScanResult) -> None:
        del result
        self._set_busy(False)
        self.last_result = None
        self.save_button.setEnabled(False)
        self.status_label.setText("Tarama iptal edildi. Excel dosyası oluşturulmadı.")

    def _scan_failed(self, message: str) -> None:
        self._set_busy(False)
        self.last_result = None
        self.save_button.setEnabled(False)
        self.status_label.setText(message)
        QMessageBox.critical(self, "Tarama Hatası", message)

    def prompt_save_as(self) -> None:
        if self.last_result is None or not self.last_result.complete:
            return
        destination = choose_export_path(self)
        if destination is not None:
            self.start_export(destination)

    def start_export(self, destination: Path) -> None:
        if self.last_result is None or not self.last_result.complete:
            return
        self.save_button.setEnabled(False)
        self.status_label.setText("Excel dosyası hazırlanıyor…")
        worker = ExportWorker(self.last_result, destination)
        self._start_worker(worker)
        worker.completed.connect(self._export_completed)
        worker.failed.connect(self._export_failed)

    def _export_completed(self, destination: str) -> None:
        self.save_button.setEnabled(True)
        self.status_label.setText(f"Excel kaydedildi: {destination}")
        QMessageBox.information(self, "Excel Hazır", f"Excel dosyası kaydedildi:\n{destination}")

    def _export_failed(self, message: str) -> None:
        self.save_button.setEnabled(True)
        self.status_label.setText(message)
        QMessageBox.critical(self, "Excel Kaydedilemedi", message)

    def clear_local_directory(self) -> None:
        answer = QMessageBox.question(
            self,
            "Yerel Rehberi Temizle",
            "Yerel rehber ve tarama ilerlemesi temizlensin mi? Bağlı hesaplar korunur.",
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        self.store.clear_directory()
        self.last_result = None
        self.save_button.setEnabled(False)
        self.status_label.setText("Yerel rehber temizlendi.")

    def import_legacy_data(self) -> None:
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "v4 rehber_data.json Dosyasını Seç",
            str(Path.home()),
            "JSON Dosyası (*.json)",
        )
        if not filename:
            return
        try:
            result = self.store.import_legacy_json(Path(filename))
        except Exception:
            QMessageBox.critical(self, "İçe Aktarma Hatası", "Eski rehber verisi okunamadı.")
            return
        imported = getattr(result, "imported", 0)
        self.status_label.setText(f"Eski rehberden {imported} kayıt içe aktarıldı.")

    def _set_busy(self, busy: bool) -> None:
        self.start_button.setEnabled(not busy)
        self.cancel_button.setEnabled(busy)
        self.add_account_button.setEnabled(not busy)
        self.disconnect_button.setEnabled(not busy)
        self.refresh_button.setEnabled(not busy)
        self.folder_tree.setEnabled(not busy)

    def _start_worker(self, worker) -> QThread:
        thread = QThread(self)
        self._threads.add(thread)
        self._workers.add(worker)
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.finished.connect(thread.quit)
        worker.finished.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)
        thread.finished.connect(
            lambda active_thread=thread, active_worker=worker: self._release_worker(
                active_thread, active_worker
            )
        )
        QTimer.singleShot(0, thread.start)
        return thread

    def _release_worker(self, thread: QThread, worker: object) -> None:
        self._threads.discard(thread)
        self._workers.discard(worker)

    def closeEvent(self, event: QCloseEvent) -> None:
        if self._scan_cancellation is not None:
            self._scan_cancellation.cancel()
        if self._folder_cancellation is not None:
            self._folder_cancellation.cancel()
        for thread in tuple(self._threads):
            thread.quit()
            thread.wait(1500)
        super().closeEvent(event)
