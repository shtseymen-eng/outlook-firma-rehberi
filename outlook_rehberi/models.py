from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from enum import Enum
from threading import Event
from typing import Protocol


class Category(str, Enum):
    POLISAN_POLIPORT = "polisan_poliport"
    CUSTOMS_YGM = "customs_ygm"
    CORPORATE = "corporate"
    PERSONAL = "personal"


class CancelledError(RuntimeError):
    """Raised when a user-requested cancellation reaches a safe boundary."""


class CancellationToken:
    def __init__(self) -> None:
        self._event = Event()

    @property
    def is_cancelled(self) -> bool:
        return self._event.is_set()

    def cancel(self) -> None:
        self._event.set()

    def raise_if_cancelled(self) -> None:
        if self.is_cancelled:
            raise CancelledError("Tarama kullanıcı tarafından iptal edildi.")


class AccessTokenProvider(Protocol):
    def get_access_token(self, account_id: str) -> str: ...


@dataclass(frozen=True, slots=True)
class AccountRef:
    account_id: str
    email: str
    display_name: str
    tenant_id: str


@dataclass(frozen=True, slots=True)
class FolderRef:
    account_id: str
    folder_id: str
    name: str
    parent_id: str | None
    path: tuple[str, ...]
    child_count: int
    is_hidden: bool


@dataclass(frozen=True, slots=True)
class Participant:
    email: str
    name: str = ""


@dataclass(frozen=True, slots=True)
class GraphMessage:
    message_id: str
    sender: Participant | None
    to: tuple[Participant, ...] = ()
    cc: tuple[Participant, ...] = ()
    bcc: tuple[Participant, ...] = ()
    body_html: str = ""
    body_text: str = ""
    sent_at: datetime | None = None
    received_at: datetime | None = None
    is_sent: bool = False


@dataclass(frozen=True, slots=True)
class GraphMessagePage:
    messages: tuple[GraphMessage, ...]
    next_link: str | None = None
    delta_link: str | None = None


@dataclass(frozen=True, slots=True)
class SourceRef:
    account_id: str
    account_email: str
    folder_id: str
    folder_path: tuple[str, ...]


@dataclass(frozen=True, slots=True)
class RankedValue:
    value: str
    quality: int
    observed_at: datetime


@dataclass(frozen=True, slots=True)
class ContactEvidence:
    email: str
    source: SourceRef
    observed_at: datetime
    name: RankedValue | None = None
    title: RankedValue | None = None
    company: RankedValue | None = None
    mobile_numbers: tuple[str, ...] = ()
    landline_numbers: tuple[str, ...] = ()
    signature_text: str = ""


@dataclass(frozen=True, slots=True)
class Contact:
    email: str
    domain: str
    name: str = ""
    title: str = ""
    company: str = ""
    mobile_numbers: tuple[str, ...] = ()
    landline_numbers: tuple[str, ...] = ()
    sources: tuple[SourceRef, ...] = ()
    signature_texts: tuple[str, ...] = ()
    category: Category | None = None


@dataclass(frozen=True, slots=True)
class ScanSelection:
    account: AccountRef
    folders: tuple[FolderRef, ...]
    own_addresses: frozenset[str]


@dataclass(frozen=True, slots=True)
class ScanProgress:
    account_email: str = ""
    folder_path: tuple[str, ...] = ()
    processed_messages: int = 0
    unique_contacts: int = 0


@dataclass(frozen=True, slots=True)
class ScanWarning:
    account_email: str
    folder_path: tuple[str, ...]
    message: str


@dataclass(frozen=True, slots=True)
class ScanStats:
    started_at: datetime
    completed_at: datetime | None = None
    accounts: int = 0
    folders: int = 0
    messages: int = 0
    unique_contacts: int = 0


@dataclass(frozen=True, slots=True)
class ScanResult:
    contacts: tuple[Contact, ...]
    stats: ScanStats
    complete: bool
    warnings: tuple[ScanWarning, ...] = ()


@dataclass(frozen=True, slots=True)
class LegacyImportResult:
    imported: int
    skipped: int
    already_migrated: bool = False

