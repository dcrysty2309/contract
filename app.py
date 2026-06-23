from __future__ import annotations

import json
import base64
import traceback
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from generator import extract_offer_data, generate_contract_docx, generate_process_verbal_docx

ROOT = Path(__file__).resolve().parent


class ContractHandler(SimpleHTTPRequestHandler):
    def end_headers(self) -> None:
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        super().end_headers()

    def translate_path(self, path: str) -> str:
        if path == "/":
            path = "/index.html"
        return str(ROOT / path.lstrip("/"))

    def do_OPTIONS(self) -> None:
        self.send_response(HTTPStatus.NO_CONTENT)
        self.end_headers()

    def do_POST(self) -> None:
        try:
            length = int(self.headers.get("Content-Length", "0"))
            payload = json.loads(self.rfile.read(length) or b"{}")
            if self.path == "/api/generate":
                filename = str(payload.get("contract_no") or payload.get("type", "produse"))
                filename = filename.replace("/", "_").replace("\\", "_").replace(" ", "_")
                docx_bytes = generate_contract_docx(payload)

                self.send_response(HTTPStatus.OK)
                self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                self.send_header("Content-Disposition", f'attachment; filename="contract_{filename}.docx"')
                self.send_header("Content-Length", str(len(docx_bytes)))
                self.end_headers()
                self.wfile.write(docx_bytes)
                return

            if self.path == "/api/generate-pv":
                filename = str(payload.get("pv_no") or payload.get("contract_no") or "pv")
                filename = filename.replace("/", "_").replace("\\", "_").replace(" ", "_")
                docx_bytes = generate_process_verbal_docx(payload)

                self.send_response(HTTPStatus.OK)
                self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                self.send_header("Content-Disposition", f'attachment; filename="proces_verbal_{filename}.docx"')
                self.send_header("Content-Length", str(len(docx_bytes)))
                self.end_headers()
                self.wfile.write(docx_bytes)
                return

            if self.path == "/api/extract":
                file_name = str(payload.get("file_name") or "offer")
                mime_type = str(payload.get("mime_type") or "")
                file_data_b64 = payload.get("file_data") or ""
                extra_files = payload.get("files") or []
                file_data = base64.b64decode(file_data_b64) if file_data_b64 else b""
                offer = extract_offer_data(
                    file_name,
                    mime_type,
                    file_data,
                    str(payload.get("type") or "produse"),
                    extra_files=extra_files if isinstance(extra_files, list) else [],
                )

                message = json.dumps(offer, ensure_ascii=False).encode("utf-8")
                self.send_response(HTTPStatus.OK)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Content-Length", str(len(message)))
                self.end_headers()
                self.wfile.write(message)
                return

            self.send_error(HTTPStatus.NOT_FOUND)
        except Exception as exc:
            traceback.print_exc()
            message = json.dumps({
                "error": str(exc),
                "traceback": traceback.format_exc(),
            }, ensure_ascii=False).encode("utf-8")
            status = HTTPStatus.INTERNAL_SERVER_ERROR
            if isinstance(exc, (RuntimeError, ValueError)) and "OPENAI_API_KEY" in str(exc):
                status = HTTPStatus.BAD_REQUEST
            if "Nu pot ajunge la OpenAI API" in str(exc):
                status = HTTPStatus.BAD_GATEWAY
            if "OpenAI API a raspuns cu 429" in str(exc):
                status = HTTPStatus.TOO_MANY_REQUESTS
            if isinstance(exc, ValueError) and "Vision" in str(exc):
                status = HTTPStatus.BAD_REQUEST
            self.send_response(status)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(message)))
            self.end_headers()
            self.wfile.write(message)


def main() -> None:
    server = ThreadingHTTPServer(("127.0.0.1", 8000), ContractHandler)
    print("Serving on http://127.0.0.1:8000")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
