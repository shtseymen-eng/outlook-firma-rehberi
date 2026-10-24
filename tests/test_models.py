from dataclasses import FrozenInstanceError
from datetime import datetime, timezone

import pytest

from outlook_rehberi.models import (
    AccountRef,
    CancelledError,
    CancellationToken,
    Category,
    FolderRef,
    Participant,
    ScanSelection,
    ScanStats,
)


def test_categories_have_stable_persisted_values():
    assert {category.value for category in Category} == {
        "polisan_poliport",
        "customs_ygm",
        "corporate",
        "personal",
    }


def test_account_and_folder_identity_models_are_immutable():
    account = AccountRef(
        account_id="home-id",
        email="owner@example.com",
        display_name="Owner",
        tenant_id="tenant-id",
    )
    folder = FolderRef(
        account_id=account.account_id,
        folder_id="inbox-id",
        name="Gelen Kutusu",
        parent_id=None,
        path=("Gelen Kutusu",),
        child_count=2,
        is_hidden=False,
    )

    with pytest.raises(FrozenInstanceError):
        account.email = "changed@example.com"
    with pytest.raises(FrozenInstanceError):
        folder.path = ("Başka",)


def test_scan_selection_carries_account_folder_and_aliases():
    account = AccountRef("a1", "owner@example.com", "Owner", "tenant")
    folder = FolderRef("a1", "f1", "Gelen", None, ("Gelen",), 0, False)

    selection = ScanSelection(
        account=account,
        folders=(folder,),
        own_addresses=frozenset({"owner@example.com", "alias@example.com"}),
    )

    assert selection.account.account_id == "a1"
    assert selection.folders == (folder,)
    assert selection.own_addresses == frozenset({"owner@example.com", "alias@example.com"})


def test_cancellation_token_raises_only_after_cancel():
    token = CancellationToken()
    token.raise_if_cancelled()

    token.cancel()

    with pytest.raises(CancelledError):
        token.raise_if_cancelled()


def test_scan_stats_defaults_are_zero_and_timestamp_is_explicit():
    started = datetime(2026, 10, 24, tzinfo=timezone.utc)

    stats = ScanStats(started_at=started)

    assert stats.started_at is started
    assert stats.accounts == 0
    assert stats.folders == 0
    assert stats.messages == 0
    assert stats.unique_contacts == 0

