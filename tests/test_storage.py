from datetime import datetime, timezone
import json
from pathlib import Path
import sqlite3

import pytest

from outlook_rehberi.models import ContactEvidence, RankedValue, SourceRef
from outlook_rehberi.storage import LocalStore


NOW = datetime(2026, 10, 24, 12, 0, tzinfo=timezone.utc)
SOURCE = SourceRef("account", "owner@example.com", "inbox", ("Gelen",))


def evidence(email: str, name: str = "Person") -> ContactEvidence:
    return ContactEvidence(
        email=email,
        source=SOURCE,
        observed_at=NOW,
        name=RankedValue(name, 30, NOW),
    )


def test_folder_commit_is_append_only_and_updates_checkpoint_atomically(tmp_path):
    store = LocalStore(tmp_path / "data.sqlite3")

    store.commit_folder("account", "inbox", "delta-1", [evidence("one@example.com")])
    store.commit_folder("account", "inbox", "delta-2", [evidence("two@example.com")])

    assert [item.email for item in store.load_directory().contacts()] == [
        "one@example.com",
        "two@example.com",
    ]
    assert store.checkpoint("account", "inbox") == "delta-2"


def test_failed_folder_commit_rolls_back_evidence_and_checkpoint(tmp_path):
    store = LocalStore(tmp_path / "data.sqlite3")
    store.commit_folder("account", "inbox", "good", [evidence("one@example.com")])

    with pytest.raises(TypeError):
        store.commit_folder(
            "account",
            "inbox",
            "must-not-commit",
            [evidence("two@example.com"), object()],
        )

    assert [item.email for item in store.load_directory().contacts()] == ["one@example.com"]
    assert store.checkpoint("account", "inbox") == "good"


def test_clear_directory_keeps_aliases_but_removes_contacts_and_checkpoints(tmp_path):
    store = LocalStore(tmp_path / "data.sqlite3")
    store.set_aliases("account", {" Alias@Example.COM ", "bad"})
    store.commit_folder("account", "inbox", "delta", [evidence("one@example.com")])

    store.clear_directory()

    assert store.load_directory().contacts() == ()
    assert store.checkpoint("account", "inbox") is None
    assert store.aliases("account") == frozenset({"alias@example.com"})


def test_corrupt_evidence_row_is_skipped_without_losing_valid_rows(tmp_path):
    path = tmp_path / "data.sqlite3"
    store = LocalStore(path)
    store.commit_folder("account", "inbox", "delta", [evidence("valid@example.com")])
    connection = sqlite3.connect(path)
    connection.execute(
        "INSERT INTO evidence(account_id, folder_id, email, payload) VALUES (?, ?, ?, ?)",
        ("account", "inbox", "broken@example.com", "{not-json"),
    )
    connection.commit()
    connection.close()

    assert [item.email for item in store.load_directory().contacts()] == [
        "valid@example.com"
    ]


def test_v4_json_import_is_one_time_deduplicated_and_non_destructive(tmp_path):
    path = tmp_path / "rehber_data.json"
    original = {
        "last_scan": "2026-10-20T10:00:00+00:00",
        "contacts": {
            "USER@EXAMPLE.COM": {
                "email": "USER@EXAMPLE.COM",
                "name": "Legacy User",
                "title": "Manager",
                "firma_display": "Example Ltd.",
                "gsm": ["05320000000"],
                "sabit": ["02120000000"],
            },
            "invalid": {"email": "invalid"},
        },
    }
    path.write_text(json.dumps(original), encoding="utf-8")
    before = path.read_bytes()
    store = LocalStore(tmp_path / "data.sqlite3")

    first = store.import_legacy_json(path)
    second = store.import_legacy_json(path)

    assert first.imported == 1
    assert first.skipped == 1
    assert first.already_migrated is False
    assert second.already_migrated is True
    assert [item.email for item in store.load_directory().contacts()] == [
        "user@example.com"
    ]
    assert path.read_bytes() == before


def test_sqlite_never_contains_authentication_cache_text(tmp_path):
    path = tmp_path / "data.sqlite3"
    store = LocalStore(path)
    store.commit_folder("account", "inbox", "delta", [evidence("one@example.com")])

    assert b"ACCESS-TOKEN" not in path.read_bytes()


def test_manual_contact_override_persists_with_highest_precedence(tmp_path):
    store = LocalStore(tmp_path / "data.sqlite3")
    store.commit_folder("account", "inbox", "delta", [evidence("one@example.com")])

    store.save_contact_override(
        "one@example.com",
        name="Düzeltilmiş Kişi",
        title="Satın Alma Müdürü",
        company="ÖRNEK A.Ş.",
    )

    contact = store.load_directory().contacts()[0]
    assert contact.name == "Düzeltilmiş Kişi"
    assert contact.title == "Satın Alma Müdürü"
    assert contact.company == "ÖRNEK A.Ş."

