from datetime import datetime, timezone

from outlook_rehberi.models import GraphMessage, Participant, SourceRef


OBSERVED_AT = datetime(2026, 10, 24, 8, 30, tzinfo=timezone.utc)
SOURCE = SourceRef(
    account_id="account-1",
    account_email="owner@example.com",
    folder_id="inbox",
    folder_path=("Gelen Kutusu",),
)


def turkish_signature_message() -> GraphMessage:
    return GraphMessage(
        message_id="message-1",
        sender=Participant("ayse@acmegumruk.com", "Ayşe Yılmaz"),
        to=(Participant("owner@example.com", "Owner"),),
        cc=(Participant("external@partner.com", "Partner User"),),
        body_html="""
            <html><body><p>Merhaba,</p><p>Dosya ektedir.</p>
            <div>Saygılarımla</div><div>Ayşe Yılmaz</div>
            <div>Gümrük Müşaviri</div><div>ACME Gümrük Ltd. Şti.</div>
            <div>GSM: 0532 123 45 67</div><div>Tel: 0212 555 66 77</div>
            </body></html>
        """,
        received_at=OBSERVED_AT,
    )


def sent_message() -> GraphMessage:
    return GraphMessage(
        message_id="message-2",
        sender=Participant("owner@example.com", "Owner"),
        to=(Participant("buyer@vendor.com", "Buyer Name"),),
        cc=(Participant("copy@vendor.com", "Copy Name"),),
        bcc=(Participant("hidden@vendor.com", "Hidden Name"),),
        body_text="İyi çalışmalar\nOwner Name\nPOLİPORT A.Ş.",
        sent_at=OBSERVED_AT,
        is_sent=True,
    )

