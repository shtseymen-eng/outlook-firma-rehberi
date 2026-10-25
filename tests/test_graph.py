import re
from threading import Timer
import time

import pytest
import requests
import responses

from outlook_rehberi.graph import (
    GraphClient,
    GraphPermissionError,
    GraphUnavailableError,
)
from outlook_rehberi.models import AccountRef, CancellationToken, CancelledError, FolderRef
from tests.fixtures.graph_responses import GRAPH, message_payload


class StaticTokens:
    def __init__(self, token: str = "top-secret-token") -> None:
        self.token = token

    def get_access_token(self, account_id: str) -> str:
        assert account_id == "account-1"
        return self.token


ACCOUNT = AccountRef("account-1", "owner@example.com", "Owner", "tenant-1")
INBOX = FolderRef("account-1", "inbox", "Gelen Kutusu", None, ("Gelen Kutusu",), 0, False)


@responses.activate
def test_profile_request_is_authorized_and_maps_mail_identity():
    responses.get(
        re.compile(rf"{GRAPH}/me\?.*"),
        json={
            "id": "profile-id",
            "displayName": "Test User",
            "mail": "test@example.com",
            "userPrincipalName": "upn@example.com",
            "tenantId": "tenant-id",
        },
        status=200,
    )

    profile = GraphClient(StaticTokens()).get_profile("account-1")

    assert profile == AccountRef(
        "account-1",
        "test@example.com",
        "Test User",
        "tenant-id",
        frozenset({"test@example.com", "upn@example.com"}),
    )
    assert responses.calls[0].request.headers["Authorization"] == "Bearer top-secret-token"


@responses.activate
def test_permission_error_does_not_expose_token_or_response_body():
    responses.get(
        re.compile(rf"{GRAPH}/me\?.*"),
        json={"error": {"message": "private mailbox detail"}},
        status=403,
    )

    with pytest.raises(GraphPermissionError) as caught:
        GraphClient(StaticTokens()).get_profile("account-1")

    assert "top-secret-token" not in str(caught.value)
    assert "private mailbox detail" not in str(caught.value)
    assert "403" in str(caught.value)


@responses.activate
def test_429_retry_after_is_honored_before_success():
    responses.get(re.compile(rf"{GRAPH}/me\?.*"), status=429, headers={"Retry-After": "2"})
    responses.get(
        re.compile(rf"{GRAPH}/me\?.*"),
        json={"id": "id", "displayName": "User", "mail": "u@example.com"},
        status=200,
    )
    delays: list[float] = []

    profile = GraphClient(StaticTokens(), sleep=delays.append).get_profile("account-1")

    assert profile.email == "u@example.com"
    assert delays == [2.0]
    assert len(responses.calls) == 2


@responses.activate
def test_repeated_server_failure_raises_typed_unavailable_error():
    for _ in range(3):
        responses.get(re.compile(rf"{GRAPH}/me\?.*"), status=503)

    with pytest.raises(GraphUnavailableError, match="503"):
        GraphClient(StaticTokens(), sleep=lambda _: None, max_retries=2).get_profile(
            "account-1"
        )

    assert len(responses.calls) == 3


@responses.activate
def test_cancelled_request_stops_before_network_access():
    token = CancellationToken()
    token.cancel()

    with pytest.raises(CancelledError):
        GraphClient(StaticTokens()).list_folder_tree(ACCOUNT, token)

    assert len(responses.calls) == 0


@responses.activate
def test_retry_wait_is_interrupted_promptly_by_cancellation():
    responses.get(
        re.compile(rf"{GRAPH}/me\?.*"), status=429, headers={"Retry-After": "1"}
    )
    token = CancellationToken()
    timer = Timer(0.05, token.cancel)
    timer.start()
    started = time.monotonic()

    try:
        with pytest.raises(CancelledError):
            GraphClient(StaticTokens())._request_json("account-1", f"{GRAPH}/me", token)
    finally:
        timer.cancel()

    assert time.monotonic() - started < 0.5


@responses.activate
def test_recursive_folder_tree_follows_pages_without_duplicates():
    root_page_2 = f"{GRAPH}/me/mailFolders?$skiptoken=root2"
    responses.get(
        re.compile(rf"{GRAPH}/me/mailFolders\?.*includeHiddenFolders=true.*"),
        json={
            "value": [
                {
                    "id": "level-1",
                    "displayName": "Gelen",
                    "parentFolderId": "root",
                    "childFolderCount": 1,
                    "isHidden": False,
                }
            ],
            "@odata.nextLink": root_page_2,
        },
    )
    responses.get(
        root_page_2,
        json={
            "value": [
                {
                    "id": "archive",
                    "displayName": "Arşiv",
                    "parentFolderId": "root",
                    "childFolderCount": 0,
                    "isHidden": False,
                }
            ]
        },
    )
    for parent, child, name, child_count in (
        ("level-1", "level-2", "Müşteriler", 1),
        ("level-2", "level-3", "2026", 1),
        ("level-3", "level-4", "Ekim", 0),
    ):
        responses.get(
            re.compile(rf"{GRAPH}/me/mailFolders/{parent}/childFolders\?.*"),
            json={
                "value": [
                    {
                        "id": child,
                        "displayName": name,
                        "parentFolderId": parent,
                        "childFolderCount": child_count,
                        "isHidden": False,
                    },
                    # A repeated item must be ignored rather than traversed twice.
                    {
                        "id": child,
                        "displayName": name,
                        "parentFolderId": parent,
                        "childFolderCount": child_count,
                        "isHidden": False,
                    },
                ]
            },
        )

    folders = GraphClient(StaticTokens()).list_folder_tree(ACCOUNT, CancellationToken())

    assert [folder.folder_id for folder in folders] == [
        "level-1",
        "level-2",
        "level-3",
        "level-4",
        "archive",
    ]
    assert folders[3].path == ("Gelen", "Müşteriler", "2026", "Ekim")


@responses.activate
def test_message_pages_map_participants_body_dates_and_delta_metadata():
    next_link = f"{GRAPH}/me/mailFolders/inbox/messages/delta?$skiptoken=next"
    delta_link = f"{GRAPH}/me/mailFolders/inbox/messages/delta?$deltatoken=done"
    responses.get(
        re.compile(rf"{GRAPH}/me/mailFolders/inbox/messages/delta\?.*"),
        json={
            "value": [message_payload("m1", "sender@example.com", html=True)],
            "@odata.nextLink": next_link,
        },
    )
    responses.get(
        next_link,
        json={
            "value": [message_payload("m2", "owner@example.com", html=False)],
            "@odata.deltaLink": delta_link,
        },
    )

    pages = list(
        GraphClient(StaticTokens()).iter_messages(
            ACCOUNT, INBOX, CancellationToken()
        )
    )

    assert len(pages) == 2
    first = pages[0].messages[0]
    second = pages[1].messages[0]
    assert first.message_id == "m1"
    assert first.sender.email == "sender@example.com"
    assert first.to[0].email == "to@example.com"
    assert first.cc[0].email == "cc@example.com"
    assert first.bcc[0].email == "bcc@example.com"
    assert first.body_html == "<p>Hello</p>"
    assert first.received_at.isoformat() == "2026-10-24T08:00:00+00:00"
    assert second.body_text == "Hello"
    assert second.is_sent is True
    assert pages[0].next_link == next_link
    assert pages[1].delta_link == delta_link

