from __future__ import annotations

from collections.abc import Iterator
from datetime import datetime, timezone

from outlook_rehberi.classifier import ContactDirectory
from outlook_rehberi.models import (
    AccountRef,
    CancellationToken,
    ContactEvidence,
    FolderRef,
    GraphMessage,
    GraphMessagePage,
    Participant,
    SourceRef,
    ScanSelection,
)
from outlook_rehberi.scanner import ScanCallbacks, ScanService


NOW = datetime(2026, 10, 24, 9, 30, tzinfo=timezone.utc)


def account(account_id: str, email: str) -> AccountRef:
    return AccountRef(account_id, email, email.split("@", 1)[0], "tenant")


def folder(owner: AccountRef, folder_id: str, *path: str) -> FolderRef:
    return FolderRef(owner.account_id, folder_id, path[-1], None, path, 0, False)


def message(
    message_id: str,
    sender: str,
    *,
    to: tuple[str, ...] = (),
    body: str = "",
) -> GraphMessage:
    return GraphMessage(
        message_id=message_id,
        sender=Participant(sender, sender.split("@", 1)[0]),
        to=tuple(Participant(value, value.split("@", 1)[0]) for value in to),
        body_text=body,
        received_at=NOW,
    )


class FakeGraph:
    def __init__(self) -> None:
        self.pages: dict[tuple[str, str], list[GraphMessagePage | Exception]] = {}
        self.folder_trees: dict[str, tuple[FolderRef, ...]] = {}
        self.message_calls: list[tuple[str, str, str | None]] = []
        self.folder_calls: list[str] = []

    def iter_messages(
        self,
        owner: AccountRef,
        selected: FolderRef,
        cancellation: CancellationToken,
        delta_link: str | None = None,
    ) -> Iterator[GraphMessagePage]:
        self.message_calls.append((owner.account_id, selected.folder_id, delta_link))
        for item in self.pages[(owner.account_id, selected.folder_id)]:
            cancellation.raise_if_cancelled()
            if isinstance(item, Exception):
                raise item
            yield item

    def list_folder_tree(
        self,
        owner: AccountRef,
        cancellation: CancellationToken,
    ) -> tuple[FolderRef, ...]:
        cancellation.raise_if_cancelled()
        self.folder_calls.append(owner.account_id)
        return self.folder_trees[owner.account_id]


class FakeStore:
    def __init__(self, initial: tuple[ContactEvidence, ...] = ()) -> None:
        self.directory = ContactDirectory(initial)
        self.alias_values: dict[str, frozenset[str]] = {}
        self.checkpoints: dict[tuple[str, str], str | None] = {}
        self.commits: list[tuple[str, str, str | None, tuple[ContactEvidence, ...]]] = []
        self.cleared = 0

    def load_directory(self) -> ContactDirectory:
        return ContactDirectory.from_records(self.directory.to_records())

    def clear_directory(self) -> None:
        self.cleared += 1
        self.directory = ContactDirectory()
        self.checkpoints.clear()

    def aliases(self, account_id: str) -> frozenset[str]:
        return self.alias_values.get(account_id, frozenset())

    def checkpoint(self, account_id: str, folder_id: str) -> str | None:
        return self.checkpoints.get((account_id, folder_id))

    def commit_folder(
        self,
        account_id: str,
        folder_id: str,
        delta_link: str | None,
        evidence: tuple[ContactEvidence, ...] | list[ContactEvidence],
    ) -> None:
        batch = tuple(evidence)
        self.commits.append((account_id, folder_id, delta_link, batch))
        for item in batch:
            self.directory.upsert(item)
        self.checkpoints[(account_id, folder_id)] = delta_link


def legacy_evidence(email: str) -> ContactEvidence:
    return ContactEvidence(
        email=email,
        source=SourceRef("legacy", "legacy", "legacy", ("Eski",)),
        observed_at=NOW,
    )


def test_successful_scan_reads_all_pages_once_clears_and_reports_monotonic_progress() -> None:
    owner = account("a1", "owner@firma.com")
    inbox = folder(owner, "inbox", "Gelen Kutusu")
    nested = folder(owner, "nested", "Gelen Kutusu", "Tedarikçiler")
    graph = FakeGraph()
    graph.pages[(owner.account_id, inbox.folder_id)] = [
        GraphMessagePage(
            (
                message("1", "bir@firma.com", to=(owner.email,)),
                message("2", owner.email, to=("iki@firma.com",)),
            ),
            next_link="page-2",
        ),
        GraphMessagePage((message("3", "uc@firma.com"),), delta_link="delta-inbox"),
    ]
    graph.pages[(owner.account_id, nested.folder_id)] = [
        GraphMessagePage((message("4", "dort@firma.com"),), delta_link="delta-nested")
    ]
    store = FakeStore((legacy_evidence("eski@firma.com"),))
    progress = []

    result = ScanService(graph, store).run(
        (ScanSelection(owner, (inbox, nested, nested), frozenset()),),
        full_rebuild=True,
        cancellation=CancellationToken(),
        callbacks=ScanCallbacks(progress.append, lambda warning: None),
    )

    assert result.complete is True
    assert store.cleared == 1
    assert [(call[0], call[1]) for call in graph.message_calls] == [
        ("a1", "inbox"),
        ("a1", "nested"),
    ]
    assert [item.email for item in result.contacts] == [
        "bir@firma.com",
        "dort@firma.com",
        "iki@firma.com",
        "uc@firma.com",
    ]
    assert [commit[2] for commit in store.commits] == ["delta-inbox", "delta-nested"]
    assert [item.processed_messages for item in progress] == sorted(
        item.processed_messages for item in progress
    )
    assert [item.unique_contacts for item in progress] == sorted(
        item.unique_contacts for item in progress
    )
    assert result.stats.messages == 4
    assert result.stats.folders == 2


def test_multi_account_evidence_is_global_and_own_alias_is_excluded() -> None:
    first = account("a1", "birinci@kurum.com")
    second = account("a2", "ikinci@kurum.com")
    first_folder = folder(first, "f1", "Gelen")
    second_folder = folder(second, "f2", "Gönderilen")
    graph = FakeGraph()
    graph.pages[("a1", "f1")] = [
        GraphMessagePage(
            (
                message(
                    "1",
                    "ortak@tedarikci.com",
                    to=(second.email, "takma@kurum.com"),
                ),
            ),
            delta_link="d1",
        )
    ]
    graph.pages[("a2", "f2")] = [
        GraphMessagePage((message("2", "ortak@tedarikci.com"),), delta_link="d2")
    ]
    store = FakeStore()
    store.alias_values["a1"] = frozenset({"takma@kurum.com"})

    result = ScanService(graph, store).run(
        (
            ScanSelection(first, (first_folder,), frozenset()),
            ScanSelection(second, (second_folder,), frozenset()),
        ),
        full_rebuild=False,
        cancellation=CancellationToken(),
        callbacks=ScanCallbacks(lambda value: None, lambda value: None),
    )

    assert result.complete is True
    assert [contact.email for contact in result.contacts] == ["ortak@tedarikci.com"]
    assert {source.account_id for source in result.contacts[0].sources} == {"a1", "a2"}


def test_incomplete_folder_does_not_commit_checkpoint_and_completed_folder_survives() -> None:
    owner = account("a1", "owner@firma.com")
    complete_folder = folder(owner, "good", "Gelen")
    broken_folder = folder(owner, "bad", "Silinen")
    graph = FakeGraph()
    graph.pages[("a1", "good")] = [
        GraphMessagePage((message("1", "tamam@firma.com"),), delta_link="good-delta")
    ]
    graph.pages[("a1", "bad")] = [
        GraphMessagePage((message("2", "yarim@firma.com"),), next_link="bad-page-2"),
        RuntimeError("secret remote failure"),
    ]
    store = FakeStore()
    warnings = []

    result = ScanService(graph, store).run(
        (ScanSelection(owner, (complete_folder, broken_folder), frozenset()),),
        full_rebuild=False,
        cancellation=CancellationToken(),
        callbacks=ScanCallbacks(lambda value: None, warnings.append),
    )

    assert result.complete is False
    assert [contact.email for contact in result.contacts] == ["tamam@firma.com"]
    assert store.checkpoint("a1", "good") == "good-delta"
    assert store.checkpoint("a1", "bad") is None
    assert [commit[1] for commit in store.commits] == ["good"]
    assert len(warnings) == 1
    assert "secret" not in warnings[0].message


def test_cancellation_stops_between_messages_and_does_not_commit_partial_folder() -> None:
    owner = account("a1", "owner@firma.com")
    inbox = folder(owner, "inbox", "Gelen")
    graph = FakeGraph()
    graph.pages[("a1", "inbox")] = [
        GraphMessagePage(
            (
                message("1", "bir@firma.com"),
                message("2", "iki@firma.com"),
            ),
            delta_link="delta",
        )
    ]
    store = FakeStore()
    cancellation = CancellationToken()

    def cancel_after_first(value: object) -> None:
        if getattr(value, "processed_messages", 0) == 1:
            cancellation.cancel()

    result = ScanService(graph, store).run(
        (ScanSelection(owner, (inbox,), frozenset()),),
        full_rebuild=False,
        cancellation=cancellation,
        callbacks=ScanCallbacks(cancel_after_first, lambda value: None),
    )

    assert result.complete is False
    assert result.contacts == ()
    assert result.stats.messages == 1
    assert store.commits == []


def test_load_folder_trees_keeps_account_boundaries() -> None:
    first = account("a1", "first@firma.com")
    second = account("a2", "second@firma.com")
    first_folder = folder(first, "f1", "Gelen")
    second_folder = folder(second, "f2", "Silinen", "2026")
    graph = FakeGraph()
    graph.folder_trees = {"a1": (first_folder,), "a2": (second_folder,)}

    trees = ScanService(graph, FakeStore()).load_folder_trees(
        (first, second), CancellationToken()
    )

    assert trees == {"a1": (first_folder,), "a2": (second_folder,)}
    assert graph.folder_calls == ["a1", "a2"]
