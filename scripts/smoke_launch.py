from __future__ import annotations

import argparse
import os
from pathlib import Path
import subprocess
import sys


def command_line() -> list[str]:
    parser = argparse.ArgumentParser(description="Uygulamayı pencere açmadan başlatır.")
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--source", action="store_true", help="Python paketini sınar.")
    mode.add_argument("--executable", type=Path, help="Paketlenmiş uygulamayı sınar.")
    arguments = parser.parse_args()
    if arguments.source:
        return [sys.executable, "-m", "outlook_rehberi"]
    executable = arguments.executable.resolve()
    if not executable.is_file():
        parser.error(f"Çalıştırılabilir dosya bulunamadı: {executable}")
    return [str(executable)]


def main() -> int:
    environment = os.environ.copy()
    environment["OUTLOOK_REHBERI_SMOKE_TEST"] = "1"
    environment.setdefault("OUTLOOK_REHBERI_CLIENT_ID", "smoke-test-client-id")
    environment.setdefault("QT_QPA_PLATFORM", "offscreen")
    completed = subprocess.run(
        command_line(),
        env=environment,
        check=False,
        timeout=45,
    )
    return completed.returncode


if __name__ == "__main__":
    raise SystemExit(main())
