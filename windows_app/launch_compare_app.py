#!/usr/bin/env python3
from __future__ import annotations

import socket
import subprocess
import sys
import time
import urllib.request
import webbrowser
from pathlib import Path


ROOT = Path(__file__).resolve().parent
HOST = "127.0.0.1"
PORT = 8765
URL = f"http://{HOST}:{PORT}"
HEALTH_URL = f"{URL}/api/health"
SERVER = ROOT / "compare_ui_server.py"
LOG = ROOT / "ui_server.log"


def preferred_python() -> Path:
    candidates = [
        ROOT / ".venv" / "Scripts" / "python.exe",
        ROOT / ".venv" / "bin" / "python",
        ROOT / ".venv" / "bin" / "python3",
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return Path(sys.executable)


def is_port_open(host: str, port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.settimeout(0.2)
        return sock.connect_ex((host, port)) == 0


def health_ok() -> bool:
    try:
        with urllib.request.urlopen(HEALTH_URL, timeout=1.5) as response:
            return response.status == 200
    except Exception:
        return False


def start_server() -> None:
    python = preferred_python()
    with LOG.open("ab") as handle:
        subprocess.Popen(
            [str(python), str(SERVER)],
            cwd=str(ROOT),
            stdout=handle,
            stderr=subprocess.STDOUT,
            start_new_session=True,
        )


def ensure_server() -> None:
    if health_ok():
        return
    if not is_port_open(HOST, PORT):
        start_server()
    for _ in range(40):
        if health_ok():
            return
        time.sleep(0.25)
    raise RuntimeError(f"Compare UI server did not start. Check {LOG}.")


def main() -> int:
    ensure_server()
    webbrowser.open(URL)
    print(f"Opened {URL}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
