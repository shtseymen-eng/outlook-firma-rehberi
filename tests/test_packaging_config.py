from __future__ import annotations

from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def read(relative: str) -> str:
    return (ROOT / relative).read_text(encoding="utf-8")


def test_platform_specs_share_entry_point_dependencies_and_windowed_mode() -> None:
    windows = read("packaging/OutlookFirmaRehberi.windows.spec")
    macos = read("packaging/OutlookFirmaRehberi.macos.spec")

    for spec in (windows, macos):
        assert "outlook_rehberi/__main__.py" in spec
        assert "msal" in spec
        assert "keyring" in spec
        assert "PySide6.QtWidgets" in spec
        assert "console=False" in spec
        assert "uygulama.py" not in spec
        assert "collect_all('win32com')" not in spec
        assert '"win32com"' in spec or "'win32com'" in spec

    assert "assets/icon.ico" in windows
    assert "assets/icon.icns" in macos
    assert "BUNDLE(" in macos


def test_old_windows_only_spec_is_removed_and_installer_name_is_versioned() -> None:
    assert not (ROOT / "OutlookFirmaRehberi.spec").exists()
    setup = read("setup.iss")

    assert 'DefaultDirName          = {localappdata}\\Programs\\{#AppDir}' in setup
    assert "OutlookFirmaRehberi_v{#AppVersion}_Windows_x64_Setup" in setup
    assert "rehber_data.json" not in setup


def test_workflow_builds_three_platform_artifacts_after_tests() -> None:
    workflow = read(".github/workflows/build.yml")

    assert "windows-latest" in workflow
    assert "macos-15-intel" in workflow
    assert "macos-15" in workflow
    assert "pytest" in workflow
    assert "scripts/smoke_launch.py" in workflow
    assert "OutlookFirmaRehberi.windows.spec" in workflow
    assert "OutlookFirmaRehberi.macos.spec" in workflow
    assert "Windows_x64_Setup.exe" in workflow
    assert "arch: AppleSilicon" in workflow
    assert "arch: Intel" in workflow
    assert "macOS_${ARCH}.dmg" in workflow
    assert ".sha256" in workflow
    assert "vars.OUTLOOK_REHBERI_CLIENT_ID" in workflow
    assert "Mail.ReadWrite" not in workflow
    assert "Mail.Send" not in workflow


def test_macos_icon_builder_and_source_smoke_mode_exist() -> None:
    icon_script = read("scripts/build_macos_icon.sh")
    smoke = read("scripts/smoke_launch.py")
    app = read("outlook_rehberi/app.py")

    assert "iconutil" in icon_script
    assert "--source" in smoke
    assert "--executable" in smoke
    assert "OUTLOOK_REHBERI_SMOKE_TEST" in app


def test_legacy_script_is_only_a_v5_compatibility_launcher() -> None:
    legacy = read("uygulama.py")

    assert "outlook_rehberi.__main__" in legacy
    assert "win32com" not in legacy
    assert "tkinter" not in legacy
