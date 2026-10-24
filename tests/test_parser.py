from datetime import datetime, timezone

from outlook_rehberi.models import GraphMessage, Participant
from outlook_rehberi.parser import (
    html_to_text,
    message_evidence,
    participant_evidence,
    sender_evidence,
    unique_message_content,
)
from tests.fixtures.messages import OBSERVED_AT, SOURCE, sent_message, turkish_signature_message


def test_html_to_text_preserves_visible_line_boundaries_and_skips_scripts():
    text = html_to_text(
        "<p>Merhaba<br>Firma</p><script>secret()</script><div>İmza</div>"
    )

    assert "Merhaba" in text
    assert "Firma" in text
    assert "İmza" in text
    assert "secret" not in text
    assert len([line for line in text.splitlines() if line.strip()]) == 3


def test_incoming_message_enriches_only_sender_from_turkish_signature():
    evidence = message_evidence(
        turkish_signature_message(), SOURCE, {"owner@example.com"}
    )

    by_email = {item.email: item for item in evidence}
    sender = by_email["ayse@acmegumruk.com"]
    copied = by_email["external@partner.com"]
    assert sender.name.value == "Ayşe Yılmaz"
    assert "Gümrük Müşaviri" in sender.title.value
    assert sender.company.value == "ACME Gümrük Ltd. Şti."
    assert sender.mobile_numbers == ("05321234567",)
    assert sender.landline_numbers == ("02125556677",)
    assert copied.name.value == "Partner User"
    assert copied.company is None
    assert copied.signature_text == ""
    assert "owner@example.com" not in by_email


def test_sent_message_collects_to_cc_and_bcc_without_copying_owners_signature():
    evidence = message_evidence(sent_message(), SOURCE, {"owner@example.com"})

    assert {item.email for item in evidence} == {
        "buyer@vendor.com",
        "copy@vendor.com",
        "hidden@vendor.com",
    }
    assert all(item.company is None for item in evidence)
    assert all(item.signature_text == "" for item in evidence)


def test_quoted_reply_signature_is_excluded():
    message = GraphMessage(
        message_id="quoted",
        sender=Participant("new@sender.com", "New Sender"),
        body_text=(
            "Merhaba, güncel bilgi aşağıdadır.\n\n"
            "From: Old Person <old@customs.com>\n"
            "Best regards\nOld Person\nOLD CUSTOMS LLC\n+90 532 000 00 00"
        ),
        received_at=OBSERVED_AT,
    )

    current = unique_message_content(message.body_text)
    evidence = sender_evidence(message, SOURCE)

    assert "Old Person" not in current
    assert evidence.email == "new@sender.com"
    assert evidence.company is None
    assert evidence.mobile_numbers == ()


def test_participant_rejects_invalid_address_and_keeps_graph_display_name():
    invalid = participant_evidence(
        Participant("not-an-email", "Invalid"), SOURCE, OBSERVED_AT
    )
    valid = participant_evidence(
        Participant(" USER@Example.COM ", "Graph Name"), SOURCE, OBSERVED_AT
    )

    assert invalid is None
    assert valid.email == "user@example.com"
    assert valid.name.value == "Graph Name"
    assert valid.name.quality == 20


def test_english_signature_extracts_role_and_company():
    message = GraphMessage(
        message_id="english",
        sender=Participant("jane@freight.example", "Jane Doe"),
        body_text=(
            "Hello\n\nBest regards,\nJane Doe\n"
            "Logistics Manager\nFREIGHT FORWARDING LLC\n+90 533 111 22 33"
        ),
        received_at=datetime(2026, 10, 24, tzinfo=timezone.utc),
    )

    evidence = sender_evidence(message, SOURCE)

    assert evidence.title.value == "Logistics Manager"
    assert evidence.company.value == "FREIGHT FORWARDING LLC"
    assert evidence.mobile_numbers == ("905331112233",)

