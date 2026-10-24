from __future__ import annotations

from collections.abc import Callable, Iterator, Mapping, Sequence
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Protocol

from .classifier import ContactDirectory, normalize_email
from .models import (
    AccountRef,
    CancellationToken,
    CancelledError,
    ContactEvidence,
    FolderRef,
    GraphMessagePage,
    ScanProgress,
    ScanResult,
    ScanSelection,
    ScanStats,
    ScanWarning,
    SourceRef,
)
from .parser import message_evidence


class GraphReader(Protocol):
    def list_folder_tree(
        self, account: AccountRef, cancellation: CancellationToken
    ) -> tuple[FolderRef, ...]: ...

    def iter_messages(
        self,
        account: AccountRef,
        folder: FolderRef,
        cancellation: CancellationToken,
        delta_link: str | None = None,
    ) -> Iterator[GraphMessagePage]: ...


class ScanStore(Protocol):
    def load_directory(self) -> ContactDirectory: ...

    def clear_directory(self) -> None: ...

    def aliases(self, account_id: str) -> frozenset[str]: ...

    def checkpoint(self, account_id: str, folder_id: str) -> str | None: ...

    def commit_folder(
        self,
        account_id: str,
        folder_id: str,
        delta_link: str | None,
        evidence: Sequence[ContactEvidence],
    ) -> None: ...


def _ignore_progress(value: ScanProgress) -> None:
    del value


def _ignore_warning(value: ScanWarning) -> None:
    del value


@dataclass(frozen=True, slots=True)
class ScanCallbacks:
    on_progress: Callable[[ScanProgress], None] = _ignore_progress
    on_warning: Callable[[ScanWarning], None] = _ignore_warning


class ScanService:
    def __init__(self, graph: GraphReader, store: ScanStore) -> None:
        self._graph = graph
        self._store = store

    def load_folder_trees(
        self,
        accounts: Sequence[AccountRef],
        cancellation: CancellationToken,
    ) -> Mapping[str, tuple[FolderRef, ...]]:
        trees: dict[str, tuple[FolderRef, ...]] = {}
        for account in accounts:
            cancellation.raise_if_cancelled()
            trees[account.account_id] = self._graph.list_folder_tree(account, cancellation)
        return trees

    def run(
        self,
        selections: Sequence[ScanSelection],
        full_rebuild: bool,
        cancellation: CancellationToken,
        callbacks: ScanCallbacks,
    ) -> ScanResult:
        started_at = datetime.now(timezone.utc)
        if full_rebuild:
            self._store.clear_directory()
        directory = self._store.load_directory()
        unique_contacts = len(directory.contacts())
        work = _deduplicated_work(selections)
        own_addresses = self._own_addresses(selections)
        warnings: list[ScanWarning] = []
        processed_messages = 0
        completed_folders = 0
        stopped = False

        for selection, folder in work:
            account = selection.account
            staged: list[ContactEvidence] = []
            checkpoint = None if full_rebuild else self._store.checkpoint(
                account.account_id, folder.folder_id
            )
            final_delta = checkpoint
            source = SourceRef(
                account_id=account.account_id,
                account_email=account.email,
                folder_id=folder.folder_id,
                folder_path=folder.path,
            )
            try:
                cancellation.raise_if_cancelled()
                pages = self._graph.iter_messages(
                    account,
                    folder,
                    cancellation,
                    delta_link=checkpoint,
                )
                for page in pages:
                    cancellation.raise_if_cancelled()
                    if page.delta_link:
                        final_delta = page.delta_link
                    for graph_message in page.messages:
                        cancellation.raise_if_cancelled()
                        staged.extend(
                            message_evidence(graph_message, source, own_addresses)
                        )
                        processed_messages += 1
                        callbacks.on_progress(
                            ScanProgress(
                                account_email=account.email,
                                folder_path=folder.path,
                                processed_messages=processed_messages,
                                unique_contacts=unique_contacts,
                            )
                        )
                cancellation.raise_if_cancelled()
                self._store.commit_folder(
                    account.account_id,
                    folder.folder_id,
                    final_delta,
                    staged,
                )
                for item in staged:
                    directory.upsert(item)
                unique_contacts = len(directory.contacts())
                completed_folders += 1
                callbacks.on_progress(
                    ScanProgress(
                        account_email=account.email,
                        folder_path=folder.path,
                        processed_messages=processed_messages,
                        unique_contacts=unique_contacts,
                    )
                )
            except CancelledError:
                warning = ScanWarning(
                    account_email=account.email,
                    folder_path=folder.path,
                    message="Tarama iptal edildi. Bu klasörün ilerleme kaydı değiştirilmedi.",
                )
                warnings.append(warning)
                callbacks.on_warning(warning)
                stopped = True
                break
            except Exception:
                warning = ScanWarning(
                    account_email=account.email,
                    folder_path=folder.path,
                    message=(
                        "Klasör tamamlanamadı. Bu klasörün ilerleme kaydı "
                        "değiştirilmedi."
                    ),
                )
                warnings.append(warning)
                callbacks.on_warning(warning)

        completed_at = datetime.now(timezone.utc)
        contacts = directory.contacts()
        complete = not stopped and not warnings
        stats = ScanStats(
            started_at=started_at,
            completed_at=completed_at,
            accounts=len({selection.account.account_id for selection in selections}),
            folders=completed_folders,
            messages=processed_messages,
            unique_contacts=len(contacts),
        )
        return ScanResult(
            contacts=contacts,
            stats=stats,
            complete=complete,
            warnings=tuple(warnings),
        )

    def _own_addresses(self, selections: Sequence[ScanSelection]) -> frozenset[str]:
        own: set[str] = set()
        account_ids: set[str] = set()
        for selection in selections:
            account_ids.add(selection.account.account_id)
            values = (selection.account.email, *selection.own_addresses)
            own.update(
                email
                for value in values
                if (email := normalize_email(value)) is not None
            )
        for account_id in account_ids:
            own.update(self._store.aliases(account_id))
        return frozenset(own)


def _deduplicated_work(
    selections: Sequence[ScanSelection],
) -> tuple[tuple[ScanSelection, FolderRef], ...]:
    work: list[tuple[ScanSelection, FolderRef]] = []
    seen: set[tuple[str, str]] = set()
    for selection in selections:
        for folder in selection.folders:
            key = selection.account.account_id, folder.folder_id
            if key in seen:
                continue
            seen.add(key)
            work.append((selection, folder))
    return tuple(work)
