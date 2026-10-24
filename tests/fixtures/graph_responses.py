GRAPH = "https://graph.microsoft.com/v1.0"


def email_recipient(address: str, name: str) -> dict:
    return {"emailAddress": {"address": address, "name": name}}


def message_payload(message_id: str, sender: str, *, html: bool = True) -> dict:
    return {
        "id": message_id,
        "from": email_recipient(sender, "Sender Name"),
        "sender": email_recipient(sender, "Sender Name"),
        "toRecipients": [email_recipient("to@example.com", "To Name")],
        "ccRecipients": [email_recipient("cc@example.com", "CC Name")],
        "bccRecipients": [email_recipient("bcc@example.com", "BCC Name")],
        "body": {
            "contentType": "html" if html else "text",
            "content": "<p>Hello</p>" if html else "Hello",
        },
        "receivedDateTime": "2026-10-24T08:00:00Z",
        "sentDateTime": "2026-10-24T07:59:00Z",
    }

