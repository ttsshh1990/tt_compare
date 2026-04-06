#!/usr/bin/env python3
from __future__ import annotations

import argparse
import shutil
import tempfile
import struct
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parent
WINDOWS_APP_DIR = ROOT / "windows_app"
WINDOWS_ZIP = ROOT / "windows_app.zip"

WINDOWS_APP_FILES = [
    "README_WINDOWS.txt",
    "Setup DOCX Compare App.bat",
    "Start DOCX Compare App.bat",
    "compare_ui.html",
    "compare_ui_redesign_demo.html",
    "compare_ui_server.py",
    "generate_diff_pdf.py",
    "launch_compare_app.py",
    "quality_check_pdf_compare.py",
    "requirements.txt",
    "setup_windows.py",
    "ui_runs/.gitkeep",
]


def windows_zipinfo(name: str) -> zipfile.ZipInfo:
    info = zipfile.ZipInfo(name)
    info.compress_type = zipfile.ZIP_STORED
    # Stable timestamp within DOS ZIP limits.
    info.date_time = (2026, 4, 6, 9, 45, 0)
    info.external_attr = 0x20
    return info


def patch_zip_for_windows(zip_path: Path) -> None:
    data = bytearray(zip_path.read_bytes())
    central_sig = b"PK\x01\x02"
    index = 0
    while True:
        index = data.find(central_sig, index)
        if index == -1:
            break
        version_made_by_offset = index + 4
        version_needed_offset = index + 6
        filename_length = struct.unpack_from("<H", data, index + 28)[0]
        extra_length = struct.unpack_from("<H", data, index + 30)[0]
        comment_length = struct.unpack_from("<H", data, index + 32)[0]
        filename = bytes(data[index + 46:index + 46 + filename_length]).decode("utf-8")
        data[version_made_by_offset:version_made_by_offset + 2] = struct.pack("<H", 20)
        data[version_needed_offset:version_needed_offset + 2] = struct.pack("<H", 20)
        external_attr = 0x10 if filename.endswith("/") else 0x20
        struct.pack_into("<I", data, index + 38, external_attr)
        index += 46 + filename_length + extra_length + comment_length
    zip_path.write_bytes(data)


def build_windows_zip(zip_path: Path = WINDOWS_ZIP) -> None:
    missing = [rel for rel in WINDOWS_APP_FILES if not (WINDOWS_APP_DIR / rel).exists()]
    if missing:
        raise FileNotFoundError(f"Missing Windows app files: {missing}")

    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_STORED) as zf:
        for rel in WINDOWS_APP_FILES:
            src = WINDOWS_APP_DIR / rel
            arcname = f"windows_app/{rel.replace('\\', '/')}"
            info = windows_zipinfo(arcname)
            info.file_size = src.stat().st_size
            with src.open("rb") as handle:
                zf.writestr(info, handle.read())
    patch_zip_for_windows(zip_path)


def validate_windows_zip(zip_path: Path = WINDOWS_ZIP) -> None:
    with zipfile.ZipFile(zip_path) as zf:
        names = zf.namelist()
        if "windows_app/compare_ui_redesign_demo.html" not in names:
            raise RuntimeError("Redesign demo file is missing from Windows ZIP.")
        if any("__pycache__" in name for name in names):
            raise RuntimeError("Windows ZIP still contains __pycache__ content.")
        for info in zf.infolist():
            if info.create_system != 0:
                raise RuntimeError(f"Non-Windows ZIP metadata found for {info.filename}: create_system={info.create_system}")

    temp_dir = Path(tempfile.mkdtemp(prefix="windows_zip_verify_"))
    try:
        with zipfile.ZipFile(zip_path) as zf:
            zf.extractall(temp_dir)
        extracted_demo = temp_dir / "windows_app" / "compare_ui_redesign_demo.html"
        if not extracted_demo.exists():
            raise RuntimeError("Validation extract did not produce the redesign demo page.")
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build a Windows-friendly ZIP bundle for the compare app.")
    parser.add_argument("--output", type=Path, default=WINDOWS_ZIP, help="Path to the output zip file.")
    parser.add_argument("--skip-validate", action="store_true", help="Skip post-build validation.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    build_windows_zip(args.output)
    if not args.skip_validate:
        validate_windows_zip(args.output)
    print(f"Built Windows ZIP: {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
