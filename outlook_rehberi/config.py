from __future__ import annotations

from dataclasses import dataclass
import json
from pathlib import Path
from typing import Mapping


class ConfigurationError(RuntimeError):
    """Raised when release configuration cannot support Microsoft sign-in."""


@dataclass(frozen=True, slots=True)
class AppConfig:
    client_id: str
    authority: str = "https://login.microsoftonline.com/common"
    graph_base_url: str = "https://graph.microsoft.com/v1.0"
    scopes: tuple[str, ...] = ("Mail.Read", "User.Read")

    @classmethod
    def load(
        cls,
        env: Mapping[str, str],
        resource_path: Path | None = None,
    ) -> "AppConfig":
        client_id = env.get("OUTLOOK_REHBERI_CLIENT_ID", "").strip()
        resource = resource_path or Path(__file__).with_name("app_config.json")
        if not client_id and resource.is_file():
            try:
                payload = json.loads(resource.read_text(encoding="utf-8"))
            except (OSError, ValueError, TypeError) as exc:
                raise ConfigurationError(
                    "Microsoft uygulama kimliği yapılandırması okunamadı."
                ) from exc
            client_id = str(payload.get("client_id", "")).strip()
        if not client_id:
            raise ConfigurationError(
                "Microsoft uygulama kimliği bulunamadı; uygulama yöneticisi "
                "OUTLOOK_REHBERI_CLIENT_ID değerini yapılandırmalıdır."
            )
        return cls(client_id=client_id)

