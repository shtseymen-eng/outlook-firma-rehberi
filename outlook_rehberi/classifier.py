from __future__ import annotations

from dataclasses import replace
from datetime import datetime
import re
import unicodedata
from typing import Iterable, Mapping, Sequence

from .models import (
    Category,
    Contact,
    ContactEvidence,
    RankedValue,
    SourceRef,
)


PERSONAL_DOMAINS = frozenset(
    {
        "gmail.com",
        "googlemail.com",
        "hotmail.com",
        "hotmail.com.tr",
        "outlook.com",
        "live.com",
        "live.com.tr",
        "windowslive.com",
        "msn.com",
        "yahoo.com",
        "yahoo.com.tr",
        "ymail.com",
        "yandex.com",
        "yandex.com.tr",
        "yandex.ru",
        "icloud.com",
        "me.com",
        "proton.me",
        "protonmail.com",
        "aol.com",
        "aim.com",
    }
)
POLISAN_DOMAINS = frozenset({"polisan.com.tr", "poliport.com"})
CUSTOMS_TERMS = (
    "gumruk",
    "gumrukleme",
    "musavirlik",
    "customs",
    "customs brokerage",
)


def normalize_email(value: str) -> str | None:
    candidate = (value or "").strip()
    if candidate.count("@") != 1:
        return None
    local, domain = candidate.rsplit("@", 1)
    local = local.strip().casefold()
    domain = domain.strip().rstrip(".").casefold()
    if not local or not domain or "." not in domain or any(char.isspace() for char in local):
        return None
    try:
        ascii_domain = domain.encode("idna").decode("ascii")
    except UnicodeError:
        return None
    if not re.fullmatch(r"[a-z0-9.!#$%&'*+/=?^_`{|}~-]+", local):
        return None
    if not re.fullmatch(r"[a-z0-9.-]+", ascii_domain):
        return None
    if any(not label or label.startswith("-") or label.endswith("-") for label in ascii_domain.split(".")):
        return None
    return f"{local}@{ascii_domain}"


def normalize_search_text(value: str) -> str:
    folded = (value or "").casefold().replace("ı", "i")
    decomposed = unicodedata.normalize("NFKD", folded)
    without_marks = "".join(char for char in decomposed if not unicodedata.combining(char))
    words = re.sub(r"[^a-z0-9]+", " ", without_marks)
    return " ".join(words.split())


def classify_contact(contact: Contact) -> Category:
    domain = contact.domain.casefold()
    if domain in POLISAN_DOMAINS:
        return Category.POLISAN_POLIPORT
    if domain in PERSONAL_DOMAINS:
        return Category.PERSONAL
    haystack = normalize_search_text(
        " ".join(
            (
                contact.email,
                contact.domain,
                contact.company,
                contact.title,
                *contact.signature_texts,
            )
        )
    )
    if re.search(r"(?:^|\s|[._-])ygm(?:$|\s|[._-])", haystack):
        return Category.CUSTOMS_YGM
    if any(term in haystack for term in CUSTOMS_TERMS):
        return Category.CUSTOMS_YGM
    return Category.CORPORATE


def company_group_key(contact: Contact) -> tuple[str, str]:
    domain = contact.domain.casefold()
    label = normalize_search_text(contact.company) or domain
    return domain, label


def _best_ranked(values: Iterable[RankedValue | None]) -> RankedValue | None:
    available = [value for value in values if value and value.value.strip()]
    if not available:
        return None
    best_quality = max(value.quality for value in available)
    by_quality = [value for value in available if value.quality == best_quality]
    newest = max(value.observed_at for value in by_quality)
    finalists = [value for value in by_quality if value.observed_at == newest]
    return min(finalists, key=lambda value: normalize_search_text(value.value))


def _source_key(source: SourceRef) -> tuple[str, tuple[str, ...], str, str]:
    return source.account_id, source.folder_path, source.folder_id, source.account_email


class ContactDirectory:
    def __init__(self, evidence: Iterable[ContactEvidence] = ()) -> None:
        self._evidence: dict[str, list[ContactEvidence]] = {}
        for item in evidence:
            self.upsert(item)

    def upsert(self, evidence: ContactEvidence) -> Contact:
        email = normalize_email(evidence.email)
        if email is None:
            raise ValueError(f"Geçersiz e-posta adresi: {evidence.email!r}")
        normalized = replace(evidence, email=email)
        self._evidence.setdefault(email, []).append(normalized)
        return self._build_contact(email)

    def _build_contact(self, email: str) -> Contact:
        items = self._evidence[email]
        domain = email.rsplit("@", 1)[1]
        name = _best_ranked(item.name for item in items)
        title = _best_ranked(item.title for item in items)
        company = _best_ranked(item.company for item in items)
        mobile_numbers = tuple(sorted({number for item in items for number in item.mobile_numbers}))
        landline_numbers = tuple(
            sorted({number for item in items for number in item.landline_numbers})
        )
        sources = tuple(sorted({item.source for item in items}, key=_source_key))
        signatures = tuple(sorted({item.signature_text for item in items if item.signature_text}))
        contact = Contact(
            email=email,
            domain=domain,
            name=name.value if name else "",
            title=title.value if title else "",
            company=company.value if company else domain.split(".")[0].upper(),
            mobile_numbers=mobile_numbers,
            landline_numbers=landline_numbers,
            sources=sources,
            signature_texts=signatures,
        )
        return replace(contact, category=classify_contact(contact))

    def contacts(self) -> tuple[Contact, ...]:
        return tuple(self._build_contact(email) for email in sorted(self._evidence))

    def to_records(self) -> tuple[dict[str, object], ...]:
        records: list[dict[str, object]] = []
        for email in sorted(self._evidence):
            items = sorted(
                self._evidence[email],
                key=lambda item: (item.observed_at, _source_key(item.source), item.email),
            )
            records.extend(_evidence_to_record(item) for item in items)
        return tuple(records)

    @classmethod
    def from_records(cls, records: Sequence[Mapping[str, object]]) -> "ContactDirectory":
        return cls(_evidence_from_record(record) for record in records)


def _ranked_to_record(value: RankedValue | None) -> dict[str, object] | None:
    if value is None:
        return None
    return {
        "value": value.value,
        "quality": value.quality,
        "observed_at": value.observed_at.isoformat(),
    }


def _ranked_from_record(record: object) -> RankedValue | None:
    if not isinstance(record, Mapping):
        return None
    return RankedValue(
        value=str(record["value"]),
        quality=int(record["quality"]),
        observed_at=datetime.fromisoformat(str(record["observed_at"])),
    )


def _evidence_to_record(evidence: ContactEvidence) -> dict[str, object]:
    return {
        "email": evidence.email,
        "observed_at": evidence.observed_at.isoformat(),
        "name": _ranked_to_record(evidence.name),
        "title": _ranked_to_record(evidence.title),
        "company": _ranked_to_record(evidence.company),
        "mobile_numbers": list(evidence.mobile_numbers),
        "landline_numbers": list(evidence.landline_numbers),
        "signature_text": evidence.signature_text,
        "source": {
            "account_id": evidence.source.account_id,
            "account_email": evidence.source.account_email,
            "folder_id": evidence.source.folder_id,
            "folder_path": list(evidence.source.folder_path),
        },
    }


def _evidence_from_record(record: Mapping[str, object]) -> ContactEvidence:
    source_record = record["source"]
    if not isinstance(source_record, Mapping):
        raise ValueError("Kaynak kaydı geçersiz.")
    return ContactEvidence(
        email=str(record["email"]),
        observed_at=datetime.fromisoformat(str(record["observed_at"])),
        source=SourceRef(
            account_id=str(source_record["account_id"]),
            account_email=str(source_record["account_email"]),
            folder_id=str(source_record["folder_id"]),
            folder_path=tuple(str(value) for value in source_record["folder_path"]),
        ),
        name=_ranked_from_record(record.get("name")),
        title=_ranked_from_record(record.get("title")),
        company=_ranked_from_record(record.get("company")),
        mobile_numbers=tuple(str(value) for value in record.get("mobile_numbers", ())),
        landline_numbers=tuple(str(value) for value in record.get("landline_numbers", ())),
        signature_text=str(record.get("signature_text", "")),
    )

