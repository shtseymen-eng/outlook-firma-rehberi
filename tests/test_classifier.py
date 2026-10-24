from datetime import datetime, timedelta, timezone

from outlook_rehberi.classifier import (
    ContactDirectory,
    classify_contact,
    company_group_key,
    normalize_email,
    normalize_search_text,
)
from outlook_rehberi.models import (
    Category,
    Contact,
    ContactEvidence,
    RankedValue,
    SourceRef,
)


NOW = datetime(2026, 10, 24, 12, 0, tzinfo=timezone.utc)
SOURCE_A = SourceRef("a", "owner-a@example.com", "inbox", ("Gelen",))
SOURCE_B = SourceRef("b", "owner-b@example.com", "sent", ("Gönderilen",))


def evidence(
    email: str,
    *,
    source: SourceRef = SOURCE_A,
    observed_at: datetime = NOW,
    name: str = "",
    name_quality: int = 20,
    company: str = "",
    title: str = "",
    signature: str = "",
    mobile: tuple[str, ...] = (),
) -> ContactEvidence:
    return ContactEvidence(
        email=email,
        source=source,
        observed_at=observed_at,
        name=RankedValue(name, name_quality, observed_at) if name else None,
        company=RankedValue(company, 30, observed_at) if company else None,
        title=RankedValue(title, 30, observed_at) if title else None,
        signature_text=signature,
        mobile_numbers=mobile,
    )


def test_normalize_email_casefolds_and_converts_unicode_domain_to_idna():
    assert normalize_email(" User@BÜCHER.DE ") == "user@xn--bcher-kva.de"
    assert normalize_email("not-an-address") is None


def test_normalize_search_text_is_turkish_character_tolerant():
    assert normalize_search_text("İSTANBUL Gümrük Müşavirliği") == (
        "istanbul gumruk musavirligi"
    )


def test_directory_globally_merges_sources_phones_and_best_values_deterministically():
    older = NOW - timedelta(days=1)
    items = [
        evidence(
            " Person@Example.COM ",
            observed_at=older,
            name="Display Name",
            name_quality=20,
            mobile=("05320000000",),
        ),
        evidence(
            "person@example.com",
            source=SOURCE_B,
            observed_at=older,
            name="Signature Name",
            name_quality=30,
            mobile=("05321111111",),
        ),
        evidence(
            "person@example.com",
            observed_at=NOW,
            name="Newest Signature",
            name_quality=30,
            mobile=("05320000000",),
        ),
    ]

    forward = ContactDirectory()
    reverse = ContactDirectory()
    for item in items:
        forward.upsert(item)
    for item in reversed(items):
        reverse.upsert(item)

    contact = forward.contacts()[0]
    assert forward.contacts() == reverse.contacts()
    assert contact.email == "person@example.com"
    assert contact.name == "Newest Signature"
    assert contact.mobile_numbers == ("05320000000", "05321111111")
    assert contact.sources == (SOURCE_A, SOURCE_B)


def test_equal_quality_and_timestamp_use_lexical_tie_break():
    directory = ContactDirectory()
    directory.upsert(evidence("tie@example.com", name="Zed Name", name_quality=30))
    directory.upsert(evidence("tie@example.com", name="Alpha Name", name_quality=30))

    assert directory.contacts()[0].name == "Alpha Name"


def test_classification_precedence_is_polisan_personal_customs_then_corporate():
    polisan = Contact(
        email="customs@polisan.com.tr",
        domain="polisan.com.tr",
        company="Gümrük Birimi",
    )
    personal = Contact(
        email="ygm.user@gmail.com",
        domain="gmail.com",
        company="Gümrük Müşavirliği",
    )
    customs = Contact(
        email="contact@acme-ygm.com",
        domain="acme-ygm.com",
        company="ACME GÜMRÜK MÜŞAVİRLİĞİ",
    )
    corporate = Contact(email="info@vendor.com", domain="vendor.com")

    assert classify_contact(polisan) is Category.POLISAN_POLIPORT
    assert classify_contact(personal) is Category.PERSONAL
    assert classify_contact(customs) is Category.CUSTOMS_YGM
    assert classify_contact(corporate) is Category.CORPORATE


def test_customs_keywords_match_signature_title_and_turkish_variants():
    by_signature = Contact(
        email="info@firm.com",
        domain="firm.com",
        signature_texts=("Yetkilendirilmiş Gümrük Müşavirliği",),
    )
    by_title = Contact(
        email="name@other.com",
        domain="other.com",
        title="Customs Broker",
    )

    assert classify_contact(by_signature) is Category.CUSTOMS_YGM
    assert classify_contact(by_title) is Category.CUSTOMS_YGM


def test_company_group_key_keeps_domains_as_the_primary_boundary():
    first = Contact("a@one.com", "one.com", company="ACME A.Ş.")
    second = Contact("b@one.com", "one.com", company="Acme Turkey")
    third = Contact("c@two.com", "two.com", company="ACME A.Ş.")

    assert company_group_key(first)[0] == company_group_key(second)[0]
    assert company_group_key(first)[0] != company_group_key(third)[0]


def test_directory_records_round_trip_without_losing_merge_metadata():
    directory = ContactDirectory()
    directory.upsert(
        evidence(
            "roundtrip@example.com",
            company="Example Ltd.",
            signature="Example Ltd.\nGümrükleme",
        )
    )

    restored = ContactDirectory.from_records(directory.to_records())
    restored.upsert(
        evidence(
            "roundtrip@example.com",
            observed_at=NOW + timedelta(hours=1),
            name="Later Name",
            name_quality=30,
        )
    )

    assert restored.contacts()[0].name == "Later Name"
    assert restored.contacts()[0].company == "Example Ltd."
    assert restored.contacts()[0].category is Category.CUSTOMS_YGM

