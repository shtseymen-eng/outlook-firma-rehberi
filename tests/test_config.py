import json

import pytest

from outlook_rehberi.config import AppConfig, ConfigurationError


def test_environment_client_id_overrides_packaged_resource(tmp_path):
    resource = tmp_path / "app_config.json"
    resource.write_text(json.dumps({"client_id": "resource-client"}), encoding="utf-8")

    config = AppConfig.load(
        {"OUTLOOK_REHBERI_CLIENT_ID": "environment-client"},
        resource_path=resource,
    )

    assert config.client_id == "environment-client"
    assert config.authority == "https://login.microsoftonline.com/common"
    assert config.graph_base_url == "https://graph.microsoft.com/v1.0"
    assert config.scopes == ("Mail.Read", "User.Read")


def test_packaged_resource_supplies_client_id_when_environment_is_empty(tmp_path):
    resource = tmp_path / "app_config.json"
    resource.write_text(json.dumps({"client_id": "resource-client"}), encoding="utf-8")

    config = AppConfig.load({}, resource_path=resource)

    assert config.client_id == "resource-client"


def test_missing_client_id_raises_user_facing_configuration_error(tmp_path):
    resource = tmp_path / "missing.json"

    with pytest.raises(ConfigurationError, match="Microsoft uygulama kimliği"):
        AppConfig.load({}, resource_path=resource)

