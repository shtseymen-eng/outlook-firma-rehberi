from __future__ import annotations

from collections.abc import Mapping, Sequence

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QAbstractItemView, QTreeWidget, QTreeWidgetItem

from ..models import AccountRef, FolderRef, ScanSelection


ACCOUNT_ROLE = int(Qt.ItemDataRole.UserRole)
FOLDER_ROLE = ACCOUNT_ROLE + 1


class SelectionRequiredError(ValueError):
    pass


class FolderTreeWidget(QTreeWidget):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setHeaderLabels(["Hesaplar ve posta klasörleri"])
        self.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.setAlternatingRowColors(True)
        self._updating_checks = False
        self._own_addresses: dict[str, frozenset[str]] = {}
        self.itemChanged.connect(self._propagate_check_state)

    def set_loading(self, message: str = "Klasörler yükleniyor…") -> None:
        self._set_status(message)

    def set_error(self, message: str) -> None:
        self._set_status(message)

    def _set_status(self, message: str) -> None:
        self.clear()
        item = QTreeWidgetItem([message])
        item.setFlags(Qt.ItemFlag.ItemIsEnabled)
        self.addTopLevelItem(item)

    def set_own_addresses(
        self, own_addresses: Mapping[str, frozenset[str] | set[str]]
    ) -> None:
        self._own_addresses = {
            account_id: frozenset(values)
            for account_id, values in own_addresses.items()
        }

    def set_data(
        self,
        accounts: Sequence[AccountRef],
        trees: Mapping[str, tuple[FolderRef, ...]],
    ) -> None:
        self._updating_checks = True
        try:
            self.clear()
            for account in accounts:
                label = (
                    f"{account.display_name} — {account.email}"
                    if account.display_name and account.display_name != account.email
                    else account.email
                )
                account_item = QTreeWidgetItem([label])
                account_item.setData(0, ACCOUNT_ROLE, account)
                account_item.setFlags(
                    account_item.flags()
                    | Qt.ItemFlag.ItemIsUserCheckable
                    | Qt.ItemFlag.ItemIsAutoTristate
                )
                account_item.setCheckState(0, Qt.CheckState.Unchecked)
                self.addTopLevelItem(account_item)

                by_id: dict[str, QTreeWidgetItem] = {}
                seen: set[str] = set()
                folders = sorted(
                    trees.get(account.account_id, ()),
                    key=lambda value: (len(value.path), value.path, value.folder_id),
                )
                for folder in folders:
                    if folder.folder_id in seen:
                        continue
                    seen.add(folder.folder_id)
                    item = QTreeWidgetItem([folder.name])
                    item.setData(0, ACCOUNT_ROLE, account)
                    item.setData(0, FOLDER_ROLE, folder)
                    item.setToolTip(0, " › ".join(folder.path))
                    item.setFlags(
                        item.flags()
                        | Qt.ItemFlag.ItemIsUserCheckable
                        | Qt.ItemFlag.ItemIsAutoTristate
                    )
                    item.setCheckState(0, Qt.CheckState.Unchecked)
                    parent = by_id.get(folder.parent_id or "")
                    (parent or account_item).addChild(item)
                    by_id[folder.folder_id] = item
            self.expandToDepth(1)
        finally:
            self._updating_checks = False

    def selections(self) -> tuple[ScanSelection, ...]:
        selections: list[ScanSelection] = []
        selected_count = 0
        all_own_addresses = {
            value
            for values in self._own_addresses.values()
            for value in values
        }
        for index in range(self.topLevelItemCount()):
            connected = self.topLevelItem(index).data(0, ACCOUNT_ROLE)
            if isinstance(connected, AccountRef):
                all_own_addresses.add(connected.email)
                all_own_addresses.update(connected.own_addresses)
        for index in range(self.topLevelItemCount()):
            account_item = self.topLevelItem(index)
            account = account_item.data(0, ACCOUNT_ROLE)
            if not isinstance(account, AccountRef):
                continue
            folders: list[FolderRef] = []
            seen: set[str] = set()
            for item in _descendants(account_item):
                selected = item.data(0, FOLDER_ROLE)
                if not isinstance(selected, FolderRef):
                    continue
                if item.checkState(0) != Qt.CheckState.Checked:
                    continue
                if selected.folder_id in seen:
                    continue
                seen.add(selected.folder_id)
                folders.append(selected)
            if folders:
                selected_count += len(folders)
                selections.append(
                    ScanSelection(account, tuple(folders), frozenset(all_own_addresses))
                )
        if not selected_count:
            raise SelectionRequiredError("Tarama için en az bir posta klasörü seçin.")
        return tuple(selections)

    def selected_account(self) -> AccountRef | None:
        item = self.currentItem()
        while item is not None:
            account = item.data(0, ACCOUNT_ROLE)
            if isinstance(account, AccountRef):
                return account
            item = item.parent()
        return None

    def _propagate_check_state(self, item: QTreeWidgetItem, column: int) -> None:
        if self._updating_checks or column != 0:
            return
        self._updating_checks = True
        try:
            state = item.checkState(0)
            if state in (Qt.CheckState.Checked, Qt.CheckState.Unchecked):
                _set_descendants(item, state)
            _update_ancestors(item.parent())
        finally:
            self._updating_checks = False


def _descendants(parent: QTreeWidgetItem):
    for index in range(parent.childCount()):
        child = parent.child(index)
        yield child
        yield from _descendants(child)


def _set_descendants(parent: QTreeWidgetItem, state: Qt.CheckState) -> None:
    for index in range(parent.childCount()):
        child = parent.child(index)
        child.setCheckState(0, state)
        _set_descendants(child, state)


def _update_ancestors(item: QTreeWidgetItem | None) -> None:
    while item is not None:
        states = {item.child(index).checkState(0) for index in range(item.childCount())}
        if states == {Qt.CheckState.Checked}:
            item.setCheckState(0, Qt.CheckState.Checked)
        elif states == {Qt.CheckState.Unchecked}:
            item.setCheckState(0, Qt.CheckState.Unchecked)
        else:
            item.setCheckState(0, Qt.CheckState.PartiallyChecked)
        item = item.parent()
