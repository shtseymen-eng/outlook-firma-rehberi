from __future__ import annotations

from datetime import datetime, timezone
import json
from pathlib import Path
import sqlite3
from typing import Sequence

from platformdirs import user_data_dir

from .classifier import (
    ContactDirectory,
    evidence_from_record,
    evidence_to_record,
    normalize_email,
)
from .models import ContactEvidence, LegacyImportResult, RankedValue, SourceRef


SCHEMA_VERSION = 1


class LocalStore:
    def __init__(self, path: Path | None = None) -> None:
        self.path = Path(path) if path else Path(
            user_data_dir("OutlookFirmaRehberi", "S.SEYMEN")
        ) / "rehber.sqlite3"
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._initialize()

    def _connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.path)
        connection.execute("PRAGMA foreign_keys = ON")
        return connection

    def _initialize(self) -> None:
        with self._connect() as connection:
            connection.executescript(
                """
                CREATE TABLE IF NOT EXISTS evidence (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    account_id TEXT NOT NULL,
                    folder_id TEXT NOT NULL,
                    email TEXT NOT NULL,
                    payload TEXT NOT NULL
                );
                CREATE INDEX IF NOT EXISTS evidence_email_idx ON evidence(email);
                CREATE TABLE IF NOT EXISTS checkpoints (
                    account_id TEXT NOT NULL,
                    folder_id TEXT NOT NULL,
                    delta_link TEXT,
                    PRIMARY KEY(account_id, folder_id)
                );
                CREATE TABLE IF NOT EXISTS aliases (
                    account_id TEXT NOT NULL,
                    email TEXT NOT NULL,
                    PRIMARY KEY(account_id, email)
                );
                CREATE TABLE IF NOT EXISTS metadata (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                );
                """
            )
            connection.execute(f"PRAGMA user_version = {SCHEMA_VERSION}")

    def load_directory(self) -> ContactDirectory:
        records: list[dict[str, object]] = []
        with self._connect() as connection:
            rows = connection.execute("SELECT payload FROM evidence ORDER BY id")
            for (payload,) in rows:
                try:
                    record = json.loads(payload)
                    evidence_from_record(record)
                except (ValueError, TypeError, KeyError, json.JSONDecodeError):
                    continue
                records.append(record)
        return ContactDirectory.from_records(records)

    def commit_folder(
        self,
        account_id: str,
        folder_id: str,
        delta_link: str | None,
        evidence: Sequence[ContactEvidence],
    ) -> None:
        serialized: list[tuple[str, str]] = []
        for item in evidence:
            if not isinstance(item, ContactEvidence):
                raise TypeError("Klasör kanıtı ContactEvidence olmalıdır.")
            record = evidence_to_record(item)
            serialized.append(
                (
                    item.email,
                    json.dumps(record, ensure_ascii=False, sort_keys=True),
                )
            )
        with self._connect() as connection:
            connection.executemany(
                "INSERT INTO evidence(account_id, folder_id, email, payload) VALUES (?, ?, ?, ?)",
                [
                    (account_id, folder_id, email, payload)
                    for email, payload in serialized
                ],
            )
            connection.execute(
                """
                INSERT INTO checkpoints(account_id, folder_id, delta_link)
                VALUES (?, ?, ?)
                ON CONFLICT(account_id, folder_id)
                DO UPDATE SET delta_link = excluded.delta_link
                """,
                (account_id, folder_id, delta_link),
            )

    def checkpoint(self, account_id: str, folder_id: str) -> str | None:
        with self._connect() as connection:
            row = connection.execute(
                "SELECT delta_link FROM checkpoints WHERE account_id = ? AND folder_id = ?",
                (account_id, folder_id),
            ).fetchone()
        return str(row[0]) if row and row[0] is not None else None

    def aliases(self, account_id: str) -> frozenset[str]:
        with self._connect() as connection:
            rows = connection.execute(
                "SELECT email FROM aliases WHERE account_id = ? ORDER BY email",
                (account_id,),
            )
            return frozenset(str(row[0]) for row in rows)

    def set_aliases(self, account_id: str, aliases: set[str] | frozenset[str]) -> None:
        normalized = sorted(
            email for alias in aliases if (email := normalize_email(alias)) is not None
        )
        with self._connect() as connection:
            connection.execute("DELETE FROM aliases WHERE account_id = ?", (account_id,))
            connection.executemany(
                "INSERT INTO aliases(account_id, email) VALUES (?, ?)",
                [(account_id, email) for email in normalized],
            )

    def clear_directory(self) -> None:
        with self._connect() as connection:
            connection.execute("DELETE FROM evidence")
            connection.execute("DELETE FROM checkpoints")

    def import_legacy_json(self, path: Path) -> LegacyImportResult:
        path = Path(path)
        migration_key = f"legacy-import:{path.resolve()}"
        with self._connect() as connection:
            if connection.execute(
                "SELECT 1 FROM metadata WHERE key = ?", (migration_key,)
            ).fetchone():
                return LegacyImportResult(0, 0, already_migrated=True)

        payload = json.loads(path.read_text(encoding="utf-8"))
        observed_at = _legacy_date(payload.get("last_scan"), path)
        contacts = payload.get("contacts", {})
        if not isinstance(contacts, dict):
            raise ValueError("Eski rehber contacts alanı geçersiz.")
        source = SourceRef(
            account_id="legacy-v4",
            account_email="legacy-v4",
            folder_id="legacy-import",
            folder_path=("v4 içe aktarım",),
        )
        imported: list[ContactEvidence] = []
        skipped = 0
        for key, raw in contacts.items():
            if not isinstance(raw, dict):
                skipped += 1
                continue
            email = normalize_email(str(raw.get("email") or key))
            if email is None:
                skipped += 1
                continue
            name = str(raw.get("name") or "").strip()
            title = str(raw.get("title") or "").strip()
            company = str(raw.get("firma_display") or raw.get("firma") or "").strip()
            imported.append(
                ContactEvidence(
                    email=email,
                    source=source,
                    observed_at=observed_at,
                    name=RankedValue(name, 25, observed_at) if name else None,
                    title=RankedValue(title, 25, observed_at) if title else None,
                    company=RankedValue(company, 25, observed_at) if company else None,
                    mobile_numbers=tuple(sorted({str(value) for value in raw.get("gsm", [])})),
                    landline_numbers=tuple(
                        sorted({str(value) for value in raw.get("sabit", [])})
                    ),
                )
            )

        records = [
            (
                item.email,
                json.dumps(evidence_to_record(item), ensure_ascii=False, sort_keys=True),
            )
            for item in imported
        ]
        with self._connect() as connection:
            connection.executemany(
                "INSERT INTO evidence(account_id, folder_id, email, payload) VALUES (?, ?, ?, ?)",
                [("legacy-v4", "legacy-import", email, raw) for email, raw in records],
            )
            connection.execute(
                "INSERT INTO metadata(key, value) VALUES (?, ?)",
                (migration_key, datetime.now(timezone.utc).isoformat()),
            )
        return LegacyImportResult(len(imported), skipped, already_migrated=False)


def _legacy_date(value: object, path: Path) -> datetime:
    if value:
        try:
            parsed = datetime.fromisoformat(str(value))
            return parsed if parsed.tzinfo else parsed.replace(tzinfo=timezone.utc)
        except ValueError:
            pass
    return datetime.fromtimestamp(path.stat().st_mtime, tz=timezone.utc)

