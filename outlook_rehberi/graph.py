from __future__ import annotations

from collections.abc import Iterator, Mapping
from datetime import datetime
import time
from typing import Callable
from urllib.parse import quote, urlencode

import requests

from .models import (
    AccessTokenProvider,
    AccountRef,
    CancellationToken,
    FolderRef,
    GraphMessage,
    GraphMessagePage,
    Participant,
)


class GraphError(RuntimeError):
    """Base class for safe, user-displayable Graph failures."""


class GraphPermissionError(GraphError):
    pass


class GraphThrottleError(GraphError):
    pass


class GraphUnavailableError(GraphError):
    pass


class GraphProtocolError(GraphError):
    pass


FOLDER_SELECT = "id,displayName,parentFolderId,childFolderCount,isHidden"
MESSAGE_SELECT = (
    "id,body,from,sender,toRecipients,ccRecipients,bccRecipients,"
    "receivedDateTime,sentDateTime,parentFolderId"
)


class GraphClient:
    def __init__(
        self,
        token_provider: AccessTokenProvider,
        *,
        base_url: str = "https://graph.microsoft.com/v1.0",
        session: requests.Session | None = None,
        sleep: Callable[[float], None] = time.sleep,
        max_retries: int = 3,
        timeout: float = 30.0,
    ) -> None:
        self._tokens = token_provider
        self._base_url = base_url.rstrip("/")
        self._session = session or requests.Session()
        self._sleep = sleep
        self._max_retries = max(0, max_retries)
        self._timeout = timeout

    def _request_json(
        self,
        account_id: str,
        url: str,
        cancellation: CancellationToken | None = None,
    ) -> dict[str, object]:
        for attempt in range(self._max_retries + 1):
            if cancellation is not None:
                cancellation.raise_if_cancelled()
            token = self._tokens.get_access_token(account_id)
            try:
                response = self._session.get(
                    url,
                    headers={
                        "Authorization": f"Bearer {token}",
                        "Accept": "application/json",
                    },
                    timeout=self._timeout,
                )
            except requests.RequestException as exc:
                if attempt < self._max_retries:
                    self._sleep(min(2**attempt, 8))
                    continue
                raise GraphUnavailableError("Microsoft Graph bağlantısı kurulamadı.") from exc

            status = response.status_code
            if status in {401, 403}:
                raise GraphPermissionError(f"Microsoft Graph erişimi reddetti ({status}).")
            if status == 429:
                if attempt < self._max_retries:
                    self._sleep(_retry_after(response, attempt))
                    continue
                raise GraphThrottleError("Microsoft Graph istek sınırına ulaşıldı (429).")
            if 500 <= status < 600:
                if attempt < self._max_retries:
                    self._sleep(_retry_after(response, attempt))
                    continue
                raise GraphUnavailableError(f"Microsoft Graph geçici olarak kullanılamıyor ({status}).")
            if not 200 <= status < 300:
                raise GraphProtocolError(f"Microsoft Graph beklenmeyen yanıt verdi ({status}).")
            try:
                payload = response.json()
            except ValueError as exc:
                raise GraphProtocolError("Microsoft Graph geçersiz JSON döndürdü.") from exc
            if not isinstance(payload, dict):
                raise GraphProtocolError("Microsoft Graph yanıt biçimi geçersiz.")
            return payload
        raise AssertionError("unreachable")

    def get_profile(self, account_id: str) -> AccountRef:
        query = urlencode(
            {"$select": "id,displayName,mail,userPrincipalName,tenantId"}
        )
        payload = self._request_json(account_id, f"{self._base_url}/me?{query}")
        email = str(payload.get("mail") or payload.get("userPrincipalName") or "").strip()
        if not email:
            raise GraphProtocolError("Microsoft hesabının e-posta adresi bulunamadı.")
        return AccountRef(
            account_id=str(payload.get("id") or account_id),
            email=email,
            display_name=str(payload.get("displayName") or email),
            tenant_id=str(payload.get("tenantId") or ""),
        )

    def list_folder_tree(
        self,
        account: AccountRef,
        cancellation: CancellationToken,
    ) -> tuple[FolderRef, ...]:
        query = urlencode(
            {
                "includeHiddenFolders": "true",
                "$top": "200",
                "$select": FOLDER_SELECT,
            }
        )
        start_url = f"{self._base_url}/me/mailFolders?{query}"
        folders: list[FolderRef] = []
        visited_folders: set[str] = set()
        visited_pages: set[str] = set()

        def visit_collection(url: str, parent_path: tuple[str, ...]) -> None:
            current_url: str | None = url
            while current_url:
                cancellation.raise_if_cancelled()
                if current_url in visited_pages:
                    raise GraphProtocolError("Microsoft Graph klasör sayfalaması döngüye girdi.")
                visited_pages.add(current_url)
                payload = self._request_json(account.account_id, current_url, cancellation)
                values = payload.get("value", [])
                if not isinstance(values, list):
                    raise GraphProtocolError("Microsoft Graph klasör listesi geçersiz.")
                for item in values:
                    if not isinstance(item, Mapping):
                        continue
                    folder_id = str(item.get("id") or "")
                    if not folder_id or folder_id in visited_folders:
                        continue
                    visited_folders.add(folder_id)
                    name = str(item.get("displayName") or folder_id)
                    folder = FolderRef(
                        account_id=account.account_id,
                        folder_id=folder_id,
                        name=name,
                        parent_id=str(item.get("parentFolderId") or "") or None,
                        path=(*parent_path, name),
                        child_count=int(item.get("childFolderCount") or 0),
                        is_hidden=bool(item.get("isHidden", False)),
                    )
                    folders.append(folder)
                    if folder.child_count:
                        child_query = urlencode(
                            {
                                "includeHiddenFolders": "true",
                                "$top": "200",
                                "$select": FOLDER_SELECT,
                            }
                        )
                        encoded = quote(folder.folder_id, safe="")
                        visit_collection(
                            f"{self._base_url}/me/mailFolders/{encoded}/childFolders?{child_query}",
                            folder.path,
                        )
                next_link = payload.get("@odata.nextLink")
                current_url = str(next_link) if next_link else None

        visit_collection(start_url, ())
        return tuple(folders)

    def iter_messages(
        self,
        account: AccountRef,
        folder: FolderRef,
        cancellation: CancellationToken,
        delta_link: str | None = None,
    ) -> Iterator[GraphMessagePage]:
        if delta_link:
            current_url: str | None = delta_link
        else:
            query = urlencode({"$top": "100", "$select": MESSAGE_SELECT})
            encoded = quote(folder.folder_id, safe="")
            current_url = f"{self._base_url}/me/mailFolders/{encoded}/messages?{query}"
        visited_pages: set[str] = set()
        while current_url:
            cancellation.raise_if_cancelled()
            if current_url in visited_pages:
                raise GraphProtocolError("Microsoft Graph mesaj sayfalaması döngüye girdi.")
            visited_pages.add(current_url)
            payload = self._request_json(account.account_id, current_url, cancellation)
            values = payload.get("value", [])
            if not isinstance(values, list):
                raise GraphProtocolError("Microsoft Graph mesaj listesi geçersiz.")
            messages = tuple(
                message
                for item in values
                if isinstance(item, Mapping)
                if (message := _message_from_payload(item, account)) is not None
            )
            next_link = payload.get("@odata.nextLink")
            delta = payload.get("@odata.deltaLink")
            yield GraphMessagePage(
                messages=messages,
                next_link=str(next_link) if next_link else None,
                delta_link=str(delta) if delta else None,
            )
            current_url = str(next_link) if next_link else None


def _retry_after(response: requests.Response, attempt: int) -> float:
    raw = response.headers.get("Retry-After", "")
    try:
        return max(0.0, min(float(raw), 60.0))
    except ValueError:
        return float(min(2**attempt, 8))


def _participant(payload: object) -> Participant | None:
    if not isinstance(payload, Mapping):
        return None
    address_payload = payload.get("emailAddress")
    if not isinstance(address_payload, Mapping):
        return None
    address = str(address_payload.get("address") or "").strip()
    if not address:
        return None
    return Participant(email=address, name=str(address_payload.get("name") or ""))


def _participants(payload: object) -> tuple[Participant, ...]:
    if not isinstance(payload, list):
        return ()
    return tuple(
        participant
        for item in payload
        if (participant := _participant(item)) is not None
    )


def _date(value: object) -> datetime | None:
    if not value:
        return None
    raw = str(value)
    if raw.endswith("Z"):
        raw = f"{raw[:-1]}+00:00"
    try:
        return datetime.fromisoformat(raw)
    except ValueError:
        return None


def _message_from_payload(
    payload: Mapping[str, object],
    account: AccountRef,
) -> GraphMessage | None:
    message_id = str(payload.get("id") or "")
    if not message_id:
        return None
    sender = _participant(payload.get("from")) or _participant(payload.get("sender"))
    body = payload.get("body")
    body_type = ""
    body_content = ""
    if isinstance(body, Mapping):
        body_type = str(body.get("contentType") or "").casefold()
        body_content = str(body.get("content") or "")
    is_sent = bool(
        sender
        and sender.email.strip().casefold() == account.email.strip().casefold()
    )
    return GraphMessage(
        message_id=message_id,
        sender=sender,
        to=_participants(payload.get("toRecipients")),
        cc=_participants(payload.get("ccRecipients")),
        bcc=_participants(payload.get("bccRecipients")),
        body_html=body_content if body_type == "html" else "",
        body_text=body_content if body_type != "html" else "",
        received_at=_date(payload.get("receivedDateTime")),
        sent_at=_date(payload.get("sentDateTime")),
        is_sent=is_sent,
    )

