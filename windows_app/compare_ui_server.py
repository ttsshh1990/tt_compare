#!/usr/bin/env python3
from __future__ import annotations

import base64
import json
import re
import traceback
import uuid
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote, urlparse

from generate_diff_pdf import run_compare, run_compare_pdf


ROOT = Path(__file__).resolve().parent
INDEX_HTML = ROOT / "compare_ui.html"
REDESIGN_DEMO_HTML = ROOT / "compare_ui_redesign_demo.html"
RUNS_DIR = ROOT / "ui_runs"
RUNS_DIR.mkdir(exist_ok=True)


def safe_filename(name: str, fallback: str) -> str:
    cleaned = Path(name or fallback).name
    cleaned = re.sub(r"[^A-Za-z0-9._'()\- ]+", "_", cleaned).strip()
    return cleaned or fallback


def json_bytes(payload: dict[str, object]) -> bytes:
    return json.dumps(payload).encode("utf-8")


class CompareHandler(BaseHTTPRequestHandler):
    server_version = "DocxCompareUI/1.0"

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/":
            self._serve_file(INDEX_HTML, "text/html; charset=utf-8")
            return
        if parsed.path == "/demo-redesign":
            self._serve_file(REDESIGN_DEMO_HTML, "text/html; charset=utf-8")
            return
        if parsed.path == "/api/health":
            self._send_json({"ok": True})
            return
        if parsed.path.startswith("/downloads/"):
            parts = [unquote(part) for part in parsed.path.split("/") if part]
            if len(parts) < 3:
                self.send_error(HTTPStatus.NOT_FOUND)
                return
            run_id, filename = parts[1], parts[2]
            file_path = RUNS_DIR / run_id / filename
            if not file_path.exists() or not file_path.is_file():
                self.send_error(HTTPStatus.NOT_FOUND)
                return
            self._serve_file(file_path, self._content_type_for(file_path), download=True)
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path != "/api/compare":
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        try:
            content_length = int(self.headers.get("Content-Length", "0"))
            body = self.rfile.read(content_length)
            payload = json.loads(body.decode("utf-8"))

            mode = str(payload.get("mode", "html")).strip().lower() or "html"
            proofread = bool(payload.get("proofread", False))
            if mode not in {"html", "pdf"}:
                self._send_json({"error": "Unsupported compare mode."}, status=HTTPStatus.BAD_REQUEST)
                return

            docx_name = safe_filename(str(payload.get("docx_name", "")), "input.docx")
            target_name = safe_filename(
                str(payload.get("target_name", "")),
                "input.html" if mode == "html" else "input.pdf",
            )
            docx_b64 = str(payload.get("docx_b64", ""))
            target_b64 = str(payload.get("target_b64", ""))

            if not docx_b64 or not target_b64:
                self._send_json({"error": "Both files are required."}, status=HTTPStatus.BAD_REQUEST)
                return

            run_id = uuid.uuid4().hex
            run_dir = RUNS_DIR / run_id
            run_dir.mkdir(parents=True, exist_ok=True)

            docx_path = run_dir / docx_name
            target_path = run_dir / target_name
            output_name = f"{Path(target_name).stem}__docx_diff_comments.pdf"
            output_path = run_dir / output_name
            summary_path = run_dir / "comparison_summary.json"

            docx_path.write_bytes(base64.b64decode(docx_b64))
            target_path.write_bytes(base64.b64decode(target_b64))

            if mode == "pdf":
                summary = run_compare_pdf(
                    docx_path=docx_path,
                    pdf_path=target_path,
                    output_path=output_path,
                    summary_json_path=summary_path,
                    proofread_mode=proofread,
                )
            else:
                summary = run_compare(
                    docx_path=docx_path,
                    html_path=target_path,
                    output_path=output_path,
                    summary_json_path=summary_path,
                    renderer="playwright",
                    proofread_mode=proofread,
                )

            self._send_json(
                {
                    "ok": True,
                    "mode": mode,
                    "proofread": proofread,
                    "run_id": run_id,
                    "summary": summary,
                    "output_name": output_name,
                    "pdf_url": f"/downloads/{run_id}/{output_name}",
                    "summary_url": f"/downloads/{run_id}/{summary_path.name}",
                }
            )
        except Exception as exc:  # pragma: no cover - runtime path
            self._send_json(
                {
                    "error": str(exc),
                    "traceback": traceback.format_exc(),
                },
                status=HTTPStatus.INTERNAL_SERVER_ERROR,
            )

    def log_message(self, fmt: str, *args: object) -> None:
        print(f"{self.address_string()} - {fmt % args}")

    def _send_json(self, payload: dict[str, object], status: int = HTTPStatus.OK) -> None:
        data = json_bytes(payload)
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)

    def _serve_file(self, path: Path, content_type: str, download: bool = False) -> None:
        data = path.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        if download:
            self.send_header("Content-Disposition", f'attachment; filename="{path.name}"')
        self.end_headers()
        self.wfile.write(data)

    @staticmethod
    def _content_type_for(path: Path) -> str:
        if path.suffix.lower() == ".pdf":
            return "application/pdf"
        if path.suffix.lower() == ".json":
            return "application/json; charset=utf-8"
        return "application/octet-stream"


def main() -> int:
    host = "127.0.0.1"
    port = 8765
    server = ThreadingHTTPServer((host, port), CompareHandler)
    print(f"Docx compare UI running at http://{host}:{port}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
