from dataclasses import dataclass
import json

import pytest

from outlook_rehberi.auth import (
    AuthenticationError,
    KeyringTokenCache,
    MicrosoftAuthenticator,
    TenantConsentRequired,
)
from outlook_rehberi.config import AppConfig


class MemoryKeyring:
    def __init__(self) -> None:
        self.values: dict[tuple[str, str], str] = {}

    def get_password(self, service: str, username: str):
        return self.values.get((service, username))

    def set_password(self, service: str, username: str, value: str):
        self.values[(service, username)] = value

    def delete_password(self, service: str, username: str):
        self.values.pop((service, username), None)


class WindowsSizedKeyring(MemoryKeyring):
    def set_password(self, service: str, username: str, value: str):
        if len(value.encode("utf-16-le")) > 2560:
            raise ValueError("credential blob exceeds Windows limit")
        super().set_password(service, username, value)


class FakeCache:
    has_state_changed = True

    def serialize(self) -> str:
        return '{"AccessToken":{"secret":"ACCESS-TOKEN"}}'


class FakeMsalApp:
    def __init__(self, accounts=(), silent_result=None, interactive_result=None) -> None:
        self._accounts = list(accounts)
        self.silent_result = silent_result
        self.interactive_result = interactive_result
        self.silent_calls = []
        self.interactive_calls = []
        self.removed = []

    def get_accounts(self):
        return list(self._accounts)

    def acquire_token_silent(self, scopes, account):
        self.silent_calls.append((tuple(scopes), account))
        return self.silent_result

    def acquire_token_interactive(self, scopes, **kwargs):
        self.interactive_calls.append((tuple(scopes), kwargs))
        return self.interactive_result

    def remove_account(self, account):
        self.removed.append(account)
        self._accounts.remove(account)


ACCOUNT_A = {
    "home_account_id": "home-a",
    "username": "a@example.com",
    "name": "Account A",
    "realm": "tenant-a",
}
ACCOUNT_B = {
    "home_account_id": "home-b",
    "username": "b@outlook.com",
    "name": "Account B",
    "realm": "consumers",
}
CONFIG = AppConfig("client-id")


def test_token_cache_is_serialized_only_into_keyring():
    backend = MemoryKeyring()
    secure_cache = KeyringTokenCache(keyring_backend=backend)

    secure_cache.save(FakeCache())

    assert backend.values[("OutlookFirmaRehberi", "msal-token-cache")].startswith("{")
    assert any(".chunk." in username for _, username in backend.values)


def test_realistic_multi_account_cache_is_chunked_and_recovers_after_restart():
    backend = WindowsSizedKeyring()
    serialized = '{"AccessToken":{"secret":"' + ("A" * 9000) + '"}}'

    class LargeCache:
        has_state_changed = True

        def serialize(self) -> str:
            return serialized

    first = KeyringTokenCache(keyring_backend=backend)
    first.save(LargeCache())

    restored = KeyringTokenCache(keyring_backend=backend).load()

    assert json.loads(restored.serialize()) == json.loads(serialized)
    assert all(len(value.encode("utf-16-le")) <= 2560 for value in backend.values.values())


def test_accounts_maps_multiple_organizational_and_personal_accounts():
    app = FakeMsalApp(accounts=(ACCOUNT_A, ACCOUNT_B))
    authenticator = MicrosoftAuthenticator(CONFIG, app=app, cache_store=MemoryKeyring())

    accounts = authenticator.accounts()

    assert [account.account_id for account in accounts] == ["home-a", "home-b"]
    assert [account.email for account in accounts] == ["a@example.com", "b@outlook.com"]
    assert accounts[0].tenant_id == "tenant-a"
    assert accounts[1].tenant_id == "consumers"


def test_access_token_uses_exact_read_only_scopes_and_silent_renewal():
    app = FakeMsalApp(accounts=(ACCOUNT_A,), silent_result={"access_token": "fresh-token"})
    authenticator = MicrosoftAuthenticator(CONFIG, app=app, cache_store=MemoryKeyring())

    assert authenticator.get_access_token("home-a") == "fresh-token"
    assert app.silent_calls == [(("Mail.Read", "User.Read"), ACCOUNT_A)]
    assert app.interactive_calls == []


def test_missing_silent_token_falls_back_to_browser_login():
    app = FakeMsalApp(
        accounts=(ACCOUNT_A,),
        silent_result=None,
        interactive_result={"access_token": "browser-token"},
    )
    authenticator = MicrosoftAuthenticator(CONFIG, app=app, cache_store=MemoryKeyring())

    assert authenticator.get_access_token("home-a") == "browser-token"
    assert app.interactive_calls == [
        (("Mail.Read", "User.Read"), {"login_hint": "a@example.com"})
    ]


def test_add_account_opens_account_selector_and_maps_claims():
    app = FakeMsalApp(
        interactive_result={
            "access_token": "token",
            "id_token_claims": {
                "oid": "object-id",
                "preferred_username": "new@example.com",
                "name": "New User",
                "tid": "tenant-id",
            },
        }
    )
    authenticator = MicrosoftAuthenticator(CONFIG, app=app, cache_store=MemoryKeyring())

    account = authenticator.add_account()

    assert account.account_id == "object-id"
    assert account.email == "new@example.com"
    assert app.interactive_calls == [
        (("Mail.Read", "User.Read"), {"prompt": "select_account"})
    ]


def test_disconnect_removes_only_selected_cached_account():
    app = FakeMsalApp(accounts=(ACCOUNT_A, ACCOUNT_B))
    authenticator = MicrosoftAuthenticator(CONFIG, app=app, cache_store=MemoryKeyring())

    authenticator.disconnect("home-a")

    assert app.removed == [ACCOUNT_A]
    assert [account.account_id for account in authenticator.accounts()] == ["home-b"]


def test_tenant_consent_failure_has_specific_user_facing_error():
    app = FakeMsalApp(
        interactive_result={
            "error": "invalid_grant",
            "error_description": "AADSTS65001: The user or administrator has not consented",
        }
    )
    authenticator = MicrosoftAuthenticator(CONFIG, app=app, cache_store=MemoryKeyring())

    with pytest.raises(TenantConsentRequired, match="yönetici onayı"):
        authenticator.add_account()


def test_other_authentication_failure_does_not_expose_server_description():
    app = FakeMsalApp(
        interactive_result={
            "error": "temporarily_unavailable",
            "error_description": "mailbox and token detail",
        }
    )
    authenticator = MicrosoftAuthenticator(CONFIG, app=app, cache_store=MemoryKeyring())

    with pytest.raises(AuthenticationError) as caught:
        authenticator.add_account()

    assert "mailbox and token detail" not in str(caught.value)

