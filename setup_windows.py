#!/usr/bin/env python3
from __future__ import annotations

import os
import platform
import shutil
import subprocess
import sys
import json
from datetime import datetime
from pathlib import Path


ROOT = Path(__file__).resolve().parent
LOG_PATH = ROOT / "setup_windows.log"
VENV_DIR = ROOT / ".venv"
VENV_PYTHON = VENV_DIR / "Scripts" / "python.exe"
REQUIREMENTS = ROOT / "requirements.txt"
MIN_PYTHON = (3, 11)
OCR_CONFIG = ROOT / "ocr_runtime.json"
TESSERACT_CANDIDATES = [
    Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
    Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"),
]


def write_log(message: str) -> None:
    with LOG_PATH.open("a", encoding="utf-8", newline="\n") as handle:
        handle.write(message)
        if not message.endswith("\n"):
            handle.write("\n")


def emit(message: str = "") -> None:
    print(message)
    write_log(message)


def fail(message: str, code: int = 1) -> int:
    emit(f"ERROR: {message}")
    emit(f"Setup log: {LOG_PATH}")
    return code


def run(cmd: list[str], *, description: str) -> None:
    emit("")
    emit(f"[STEP] {description}")
    emit(f"[CMD] {' '.join(cmd)}")
    process = subprocess.Popen(
        cmd,
        cwd=str(ROOT),
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    assert process.stdout is not None
    for line in process.stdout:
        text = line.rstrip("\n")
        print(text)
        write_log(text)
    return_code = process.wait()
    if return_code != 0:
        raise RuntimeError(f"Command failed with exit code {return_code}: {' '.join(cmd)}")


def find_tesseract() -> Path | None:
    command_path = shutil.which("tesseract")
    if command_path:
        return Path(command_path)
    for candidate in TESSERACT_CANDIDATES:
        if candidate.exists():
            return candidate
    return None


def write_ocr_config(tesseract_path: Path) -> None:
    OCR_CONFIG.write_text(
        json.dumps({"tesseract_path": str(tesseract_path)}, indent=2),
        encoding="utf-8",
    )


def install_tesseract() -> None:
    existing = find_tesseract()
    if existing is not None:
        emit("")
        emit(f"[STEP] Reusing existing Tesseract OCR at {existing}")
        write_ocr_config(existing)
        emit(f"[INFO] Wrote OCR runtime config: {OCR_CONFIG}")
        return

    winget_path = shutil.which("winget")
    if winget_path is None:
        raise RuntimeError(
            "Tesseract OCR is not installed. Install it with "
            "`winget install -e --id UB-Mannheim.TesseractOCR` "
            "or install the UB Mannheim Tesseract package manually, then rerun setup."
        )

    run(
        [
            winget_path,
            "install",
            "-e",
            "--id",
            "UB-Mannheim.TesseractOCR",
            "--accept-package-agreements",
            "--accept-source-agreements",
        ],
        description="Install Tesseract OCR for DOCX-vs-PDF compare mode",
    )

    installed = find_tesseract()
    if installed is None:
        raise RuntimeError(
            "Tesseract OCR installation completed but `tesseract.exe` was not found. "
            "Open a new Command Prompt, run `tesseract --version`, and rerun setup if needed."
        )
    emit(f"[INFO] Tesseract OCR detected at {installed}")
    write_ocr_config(installed)
    emit(f"[INFO] Wrote OCR runtime config: {OCR_CONFIG}")


def check_environment() -> None:
    emit("DOCX/HTML Compare App Windows setup")
    emit(f"Started: {datetime.now().isoformat(timespec='seconds')}")
    emit(f"Platform: {platform.platform()}")
    emit(f"Python executable: {sys.executable}")
    emit(f"Python version: {platform.python_version()}")
    emit(f"Working directory: {ROOT}")
    emit(f"PATH: {os.environ.get('PATH', '')}")
    if sys.version_info < MIN_PYTHON:
        required = ".".join(str(part) for part in MIN_PYTHON)
        current = platform.python_version()
        raise RuntimeError(f"Python {required}+ is required. Current version is {current}.")
    if not REQUIREMENTS.exists():
        raise RuntimeError(f"Missing requirements file: {REQUIREMENTS}")


def create_venv() -> None:
    if VENV_PYTHON.exists():
        emit("")
        emit(f"[STEP] Reusing existing virtual environment at {VENV_DIR}")
        return
    run([sys.executable, "-m", "venv", str(VENV_DIR)], description="Create virtual environment")
    if not VENV_PYTHON.exists():
        raise RuntimeError(f"Virtual environment was created but {VENV_PYTHON} was not found.")


def install_dependencies() -> None:
    run([str(VENV_PYTHON), "-m", "pip", "install", "--upgrade", "pip"], description="Upgrade pip")
    run([str(VENV_PYTHON), "-m", "pip", "install", "-r", str(REQUIREMENTS)], description="Install Python packages")
    run([str(VENV_PYTHON), "-m", "playwright", "install", "chromium"], description="Install Chromium for Playwright")


def smoke_test() -> None:
    run(
        [
            str(VENV_PYTHON),
            "-c",
            (
                "import fitz, playwright, pypdf; "
                "print('fitz', getattr(fitz, '__doc__', 'ok').splitlines()[0] if getattr(fitz, '__doc__', None) else 'ok'); "
                "print('playwright', getattr(playwright, '__version__', 'ok')); "
                "print('pypdf', getattr(pypdf, '__version__', 'ok'))"
            ),
        ],
        description="Verify installed Python modules",
    )
    tesseract_path = find_tesseract()
    if tesseract_path is None:
        raise RuntimeError(
            "Tesseract OCR is still missing after setup. "
            "Run `winget install -e --id UB-Mannheim.TesseractOCR`, then rerun setup."
        )
    run([str(tesseract_path), "--version"], description="Verify Tesseract OCR")
    if not OCR_CONFIG.exists():
        raise RuntimeError(
            f"Tesseract OCR was found at {tesseract_path}, but runtime config was not written: {OCR_CONFIG}"
        )
    emit(f"[INFO] OCR runtime config ready: {OCR_CONFIG}")


def main() -> int:
    LOG_PATH.write_text("", encoding="utf-8")
    try:
        check_environment()
        create_venv()
        install_dependencies()
        install_tesseract()
        smoke_test()
    except Exception as exc:
        return fail(str(exc))

    emit("")
    emit("Setup completed successfully.")
    emit(f"Virtual environment: {VENV_DIR}")
    emit(f"Setup log: {LOG_PATH}")
    emit("Next step: run 'Start DOCX Compare App.bat'")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
