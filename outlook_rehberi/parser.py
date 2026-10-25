from __future__ import annotations

from datetime import datetime, timezone
from html.parser import HTMLParser
import re
from typing import AbstractSet

from .classifier import normalize_email
from .models import (
    ContactEvidence,
    GraphMessage,
    Participant,
    RankedValue,
    SourceRef,
)


MOBILE_RE = re.compile(
    r"(?<!\d)(?:\+?90[\s.\-]?)?0?5[0-9]{2}[\s.\-]?[0-9]{3}"
    r"[\s.\-]?[0-9]{2}[\s.\-]?[0-9]{2}(?!\d)"
)
LANDLINE_RE = re.compile(
    r"(?<!\d)(?:0\s*(?:212|216|232|242|262|312|322|332|352|362|392|"
    r"412|422|432|442|452|462|472|482)|0\s*[2-4][0-9]{2})"
    r"[\s.\-]?[0-9]{3}[\s.\-]?[0-9]{2}[\s.\-]?[0-9]{2}(?!\d)"
)

TITLE_KEYWORDS = (
    "müdür",
    "uzman",
    "şef",
    "direktör",
    "koordinatör",
    "asistan",
    "yetkili",
    "sorumlu",
    "müşavir",
    "operatör",
    "lojistik",
    "satın alma",
    "muhasebe",
    "finans",
    "insan kaynakları",
    "pazarlama",
    "satış",
    "yönetici",
    "mühendis",
    "tekniker",
    "danışman",
    "analist",
    "planlama",
    "depo",
    "sevkiyat",
    "ihracat",
    "ithalat",
    "operasyon",
    "manager",
    "director",
    "chief",
    "officer",
    "specialist",
    "coordinator",
    "assistant",
    "analyst",
    "engineer",
    "consultant",
    "supervisor",
    "executive",
    "head of",
    "logistics",
    "procurement",
    "operations",
    "sales",
    "marketing",
    "finance",
    "accounting",
    "transport",
    "shipping",
    "export",
    "import",
    "supply chain",
    "customs broker",
)

SIGNATURE_MARKER_RE = re.compile(
    r"^\s*(?:--+|iyi çalışmalar|saygılarımla|regards|best regards|thanks|"
    r"teşekkürler|kind regards|br,?|cheers|sincerely|iyi günler|good day)\s*[,.]?\s*$",
    re.IGNORECASE,
)
QUOTE_MARKER_RE = re.compile(
    r"^\s*(?:from|kimden|sent|gönderilen|to|kime|subject|konu)\s*:|"
    r"^\s*-{2,}\s*(?:original message|özgün ileti)\s*-{2,}\s*$|"
    r"^\s*on\s+.+\s+wrote\s*:\s*$",
    re.IGNORECASE,
)
COMPANY_SUFFIX_RE = re.compile(
    r"\b(?:a\.?\s*ş\.?|ltd\.?|şti\.?|inc\.?|co\.?|llc\.?|gmbh)\b",
    re.IGNORECASE,
)


class _HTMLTextExtractor(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.parts: list[str] = []
        self._skip_depth = 0
        self._quote_depth = 0
        self._stack: list[tuple[str, bool]] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        attributes = {name.casefold(): (value or "").casefold() for name, value in attrs}
        quote_hint = " ".join((attributes.get("class", ""), attributes.get("id", "")))
        starts_quote = tag == "blockquote" or any(
            marker in quote_hint
            for marker in ("gmail_quote", "yahoo_quoted", "divrplyfwdmsg", "reference-message")
        )
        if tag not in {"br", "img", "hr", "meta", "link", "input"}:
            self._stack.append((tag, starts_quote))
        if starts_quote:
            self._quote_depth += 1
        if tag in {"script", "style"}:
            self._skip_depth += 1
        elif self._skip_depth == 0 and self._quote_depth == 0 and tag in {
            "br", "p", "div", "tr", "li", "td", "th"
        }:
            self.parts.append("\n")

    def handle_endtag(self, tag: str) -> None:
        if tag in {"script", "style"}:
            self._skip_depth = max(0, self._skip_depth - 1)
        elif self._skip_depth == 0 and self._quote_depth == 0 and tag in {"p", "div", "tr", "li"}:
            self.parts.append("\n")
        for index in range(len(self._stack) - 1, -1, -1):
            open_tag, started_quote = self._stack[index]
            if open_tag == tag:
                del self._stack[index:]
                if started_quote:
                    self._quote_depth = max(0, self._quote_depth - 1)
                break

    def handle_data(self, data: str) -> None:
        if self._skip_depth == 0 and self._quote_depth == 0:
            self.parts.append(data)


def html_to_text(html: str) -> str:
    parser = _HTMLTextExtractor()
    try:
        parser.feed(html or "")
    except Exception:
        return html or ""
    lines = [re.sub(r"\s+", " ", line).strip() for line in "".join(parser.parts).splitlines()]
    return "\n".join(line for line in lines if line)


def unique_message_content(text: str) -> str:
    unique_lines: list[str] = []
    for line in (text or "").splitlines():
        if QUOTE_MARKER_RE.search(line):
            break
        unique_lines.append(line)
    return "\n".join(unique_lines).strip()


def extract_signature(text: str) -> str:
    lines = (text or "").splitlines()
    for index, line in enumerate(lines):
        if SIGNATURE_MARKER_RE.match(line):
            return "\n".join(lines[index:]).strip()
    candidate = "\n".join(lines[-20:]).strip()
    if not candidate:
        return ""
    has_identity = bool(_signature_name(candidate))
    has_detail = any(
        (
            bool(_title(candidate)),
            bool(_company(candidate)),
            bool(MOBILE_RE.search(candidate)),
            bool(LANDLINE_RE.search(candidate)),
            bool(re.search(r"\b[^\s@]+@[^\s@]+\.[^\s@]+\b", candidate)),
        )
    )
    return candidate if has_identity and has_detail else ""


def _normalize_email(value: str) -> str | None:
    return normalize_email(value)


def _observed_at(message: GraphMessage) -> datetime:
    return message.received_at or message.sent_at or datetime.min.replace(tzinfo=timezone.utc)


def _clean_display_name(value: str) -> str:
    candidate = re.sub(r"\(.*?\)", "", (value or "").replace('"', "").strip()).strip()
    if "@" in candidate or not 1 < len(candidate.split()) <= 5:
        return ""
    return candidate


def _signature_name(signature: str) -> str:
    for line in signature.splitlines():
        candidate = line.strip(" ,.-")
        if not candidate or "@" in candidate or re.search(r"\d", candidate):
            continue
        if SIGNATURE_MARKER_RE.match(candidate):
            continue
        lowered = candidate.casefold()
        if any(keyword in lowered for keyword in TITLE_KEYWORDS):
            continue
        words = candidate.split()
        if 2 <= len(words) <= 5 and not COMPANY_SUFFIX_RE.search(candidate):
            return candidate
    return ""


def _title(signature: str) -> str:
    for line in signature.splitlines():
        candidate = line.strip()
        lowered = candidate.casefold()
        if candidate and any(keyword in lowered for keyword in TITLE_KEYWORDS):
            return candidate
    return ""


def _company(signature: str) -> str:
    for line in signature.splitlines():
        candidate = line.strip()
        if not 2 < len(candidate) <= 100 or "@" in candidate:
            continue
        if re.search(r"\d{4,}", candidate):
            continue
        if COMPANY_SUFFIX_RE.search(candidate):
            return candidate
        letters = [char for char in candidate if char.isalpha()]
        if len(letters) >= 4 and all(char.isupper() for char in letters):
            if not any(keyword in candidate.casefold() for keyword in TITLE_KEYWORDS):
                return candidate
    return ""


def _phones(pattern: re.Pattern[str], signature: str) -> tuple[str, ...]:
    values = {re.sub(r"\D", "", match.group(0)) for match in pattern.finditer(signature)}
    return tuple(sorted(value for value in values if value))


def participant_evidence(
    participant: Participant,
    source: SourceRef,
    observed_at: datetime,
) -> ContactEvidence | None:
    email = _normalize_email(participant.email)
    if email is None:
        return None
    display_name = _clean_display_name(participant.name)
    return ContactEvidence(
        email=email,
        source=source,
        observed_at=observed_at,
        name=RankedValue(display_name, 20, observed_at) if display_name else None,
    )


def sender_evidence(message: GraphMessage, source: SourceRef) -> ContactEvidence | None:
    if message.sender is None:
        return None
    email = _normalize_email(message.sender.email)
    if email is None:
        return None
    observed_at = _observed_at(message)
    body = html_to_text(message.body_html) if message.body_html else message.body_text
    signature = extract_signature(unique_message_content(body))
    signature_name = _signature_name(signature)
    display_name = _clean_display_name(message.sender.name)
    name = signature_name or display_name
    name_quality = 30 if signature_name else 20
    title = _title(signature)
    company = _company(signature)
    return ContactEvidence(
        email=email,
        source=source,
        observed_at=observed_at,
        name=RankedValue(name, name_quality, observed_at) if name else None,
        title=RankedValue(title, 30, observed_at) if title else None,
        company=RankedValue(company, 30, observed_at) if company else None,
        mobile_numbers=_phones(MOBILE_RE, signature),
        landline_numbers=_phones(LANDLINE_RE, signature),
        signature_text=signature,
    )


def message_evidence(
    message: GraphMessage,
    source: SourceRef,
    own_addresses: AbstractSet[str],
) -> list[ContactEvidence]:
    own = {email for value in own_addresses if (email := _normalize_email(value))}
    evidence_by_email: dict[str, ContactEvidence] = {}
    sender = sender_evidence(message, source)
    if sender is not None and sender.email not in own:
        evidence_by_email[sender.email] = sender
    observed_at = _observed_at(message)
    for participant in (*message.to, *message.cc, *message.bcc):
        evidence = participant_evidence(participant, source, observed_at)
        if evidence is None or evidence.email in own:
            continue
        evidence_by_email.setdefault(evidence.email, evidence)
    return list(evidence_by_email.values())

