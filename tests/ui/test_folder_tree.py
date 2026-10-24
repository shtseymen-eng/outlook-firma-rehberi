from __future__ import annotations

import pytest
from PySide6.QtCore import Qt

from outlook_rehberi.models import AccountRef, FolderRef
from outlook_rehberi.ui.folder_tree import FolderTreeWidget, SelectionRequiredError


def account(account_id: str, email: str) -> AccountRef:
    return AccountRef(account_id, email, email, "tenant")


def folder(
    owner: AccountRef,
    folder_id: str,
    name: str,
    parent_id: str | None,
    path: tuple[str, ...],
) -> FolderRef:
    return FolderRef(owner.account_id, folder_id, name, parent_id, path, 0, False)


def test_builds_account_and_unlimited_folder_hierarchy(qtbot) -> None:
    owner = account("a1", "owner@firma.com")
    inbox = folder(owner, "inbox", "Gelen", None, ("Gelen",))
    year = folder(owner, "year", "2026", "inbox", ("Gelen", "2026"))
    customer = folder(
        owner, "customer", "Müşteri", "year", ("Gelen", "2026", "Müşteri")
    )
    tree = FolderTreeWidget()
    qtbot.addWidget(tree)

    tree.set_data((owner,), {"a1": (customer, inbox, year)})

    account_item = tree.topLevelItem(0)
    assert account_item.text(0) == "owner@firma.com"
    assert account_item.child(0).text(0) == "Gelen"
    assert account_item.child(0).child(0).text(0) == "2026"
    assert account_item.child(0).child(0).child(0).text(0) == "Müşteri"


def test_tristate_propagates_and_selections_are_unique_across_accounts(qtbot) -> None:
    first = account("a1", "first@firma.com")
    second = account("a2", "second@firma.com")
    inbox = folder(first, "inbox", "Gelen", None, ("Gelen",))
    nested = folder(first, "nested", "Alt", "inbox", ("Gelen", "Alt"))
    archive = folder(first, "archive", "Arşiv", None, ("Arşiv",))
    sent = folder(second, "sent", "Gönderilen", None, ("Gönderilen",))
    tree = FolderTreeWidget()
    qtbot.addWidget(tree)
    tree.set_data(
        (first, second),
        {"a1": (inbox, nested, nested, archive), "a2": (sent,)},
    )

    first_item = tree.topLevelItem(0)
    first_item.setCheckState(0, Qt.CheckState.Checked)
    tree.topLevelItem(1).child(0).setCheckState(0, Qt.CheckState.Checked)
    inbox_item = next(
        first_item.child(index)
        for index in range(first_item.childCount())
        if first_item.child(index).text(0) == "Gelen"
    )

    assert inbox_item.checkState(0) == Qt.CheckState.Checked
    assert inbox_item.child(0).checkState(0) == Qt.CheckState.Checked
    selections = tree.selections()
    assert [selection.account.account_id for selection in selections] == ["a1", "a2"]
    assert [item.folder_id for item in selections[0].folders] == [
        "archive",
        "inbox",
        "nested",
    ]
    assert [item.folder_id for item in selections[1].folders] == ["sent"]

    inbox_item.child(0).setCheckState(0, Qt.CheckState.Unchecked)
    assert first_item.checkState(0) == Qt.CheckState.PartiallyChecked


def test_zero_selection_and_refresh_error_states_are_explicit(qtbot) -> None:
    owner = account("a1", "owner@firma.com")
    tree = FolderTreeWidget()
    qtbot.addWidget(tree)

    tree.set_loading("Klasörler yükleniyor…")
    assert tree.topLevelItem(0).text(0) == "Klasörler yükleniyor…"
    tree.set_error("Klasörler alınamadı")
    assert tree.topLevelItem(0).text(0) == "Klasörler alınamadı"
    tree.set_data((owner,), {"a1": ()})

    with pytest.raises(SelectionRequiredError, match="en az bir"):
        tree.selections()
