from __future__ import annotations

from collections.abc import Mapping
from typing import Any

import keyring
import msal

from .config import AppConfig
from .models import AccountRef


class AuthenticationError(RuntimeError):
    """A safe, user-facing Microsoft authentication failure."""


class TenantConsentRequired(AuthenticationError):
    pass


class KeyringTokenCache:
    def __init__(
        self,
        *,
        service: str = "OutlookFirmaRehberi",
        username: str = "msal-token-cache",
        keyring_backend: Any = keyring,
    ) -> None:
        self.service = service
        self.username = username
        self._keyring = keyring_backend

    def load(self) -> msal.SerializableTokenCache:
        cache = msal.SerializableTokenCache()
        serialized = self._keyring.get_password(self.service, self.username)
        if serialized:
            try:
                cache.deserialize(serialized)
            except (ValueError, TypeError):
                self._keyring.delete_password(self.service, self.username)
        return cache

    def save(self, cache: msal.SerializableTokenCache) -> None:
        if getattr(cache, "has_state_changed", True):
            self._keyring.set_password(self.service, self.username, cache.serialize())


class MicrosoftAuthenticator:
    def __init__(
        self,
        config: AppConfig,
        *,
        app: Any | None = None,
        cache_store: Any = keyring,
        app_factory: Any = msal.PublicClientApplication,
    ) -> None:
        self._config = config
        self._secure_cache = KeyringTokenCache(keyring_backend=cache_store)
        self._cache = self._secure_cache.load()
        self._app = app or app_factory(
            client_id=config.client_id,
            authority=config.authority,
            token_cache=self._cache,
        )

    def _persist(self) -> None:
        self._secure_cache.save(self._cache)

    def accounts(self) -> tuple[AccountRef, ...]:
        return tuple(_account_ref(account) for account in self._app.get_accounts())

    def add_account(self) -> AccountRef:
        result = self._app.acquire_token_interactive(
            scopes=list(self._config.scopes),
            prompt="select_account",
        )
        token = _access_token(result)
        del token
        claims = result.get("id_token_claims", {}) if isinstance(result, Mapping) else {}
        if not isinstance(claims, Mapping):
            claims = {}
        email = str(
            claims.get("preferred_username")
            or claims.get("email")
            or claims.get("upn")
            or ""
        )
        account_id = str(claims.get("oid") or claims.get("sub") or "")
        if not email or not account_id:
            raise AuthenticationError("Microsoft hesabı kimliği okunamadı.")
        self._persist()
        return AccountRef(
            account_id=account_id,
            email=email,
            display_name=str(claims.get("name") or email),
            tenant_id=str(claims.get("tid") or ""),
        )

    def get_access_token(self, account_id: str) -> str:
        account = self._find_account(account_id)
        result = self._app.acquire_token_silent(
            scopes=list(self._config.scopes),
            account=account,
        )
        if not (isinstance(result, Mapping) and result.get("access_token")):
            if isinstance(result, Mapping) and result.get("error"):
                _raise_auth_error(result)
            result = self._app.acquire_token_interactive(
                scopes=list(self._config.scopes),
                login_hint=str(account.get("username") or ""),
            )
        token = _access_token(result)
        self._persist()
        return token

    def disconnect(self, account_id: str) -> None:
        account = self._find_account(account_id)
        self._app.remove_account(account)
        self._persist()

    def _find_account(self, account_id: str) -> Mapping[str, object]:
        for account in self._app.get_accounts():
            if str(account.get("home_account_id") or "") == account_id:
                return account
        raise AuthenticationError("Bağlı Microsoft hesabı bulunamadı.")


def _account_ref(account: Mapping[str, object]) -> AccountRef:
    email = str(account.get("username") or "")
    return AccountRef(
        account_id=str(account.get("home_account_id") or ""),
        email=email,
        display_name=str(account.get("name") or email),
        tenant_id=str(account.get("realm") or ""),
    )


def _access_token(result: object) -> str:
    if not isinstance(result, Mapping):
        raise AuthenticationError("Microsoft oturum açma işlemi tamamlanamadı.")
    token = result.get("access_token")
    if token:
        return str(token)
    _raise_auth_error(result)
    raise AssertionError("unreachable")


def _raise_auth_error(result: Mapping[str, object]) -> None:
    code = str(result.get("error") or "")
    description = str(result.get("error_description") or "")
    if code in {"consent_required", "interaction_required"} or "AADSTS65001" in description:
        raise TenantConsentRequired(
            "Kurumsal hesabınız Microsoft posta okuma izni için yönetici onayı istiyor."
        )
    raise AuthenticationError("Microsoft oturum açma işlemi tamamlanamadı.")

