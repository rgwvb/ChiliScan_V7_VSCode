#!/usr/bin/env python3
"""Minimal authenticated API scaffold for ChiliScan and Jetson."""

import hmac
import json
import math
import os
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlsplit


HOST = os.environ.get("CHILISCAN_HOST", "127.0.0.1")
PORT = int(os.environ.get("CHILISCAN_PORT", "8765"))
TOKEN = os.environ.get("CHILISCAN_TOKEN", "")
ALLOWED_ORIGIN = os.environ.get("CHILISCAN_ORIGIN", "https://rgwvb.github.io")
MAX_BODY_BYTES = 2_000_000
MAX_SPECTRA = 100
MAX_POINTS_PER_SPECTRUM = 20_000


class RequestError(Exception):
    def __init__(self, status, message):
        super().__init__(message)
        self.status = status


def validate_request(payload):
    if not isinstance(payload, dict) or payload.get("api_version") != 1:
        raise RequestError(400, "Expected api_version 1.")

    spectra = payload.get("spectra")
    if not isinstance(spectra, list) or not spectra or len(spectra) > MAX_SPECTRA:
        raise RequestError(400, f"spectra must contain 1 to {MAX_SPECTRA} samples.")

    for spectrum in spectra:
        if not isinstance(spectrum, dict):
            raise RequestError(400, "Each spectrum must be an object.")
        name = spectrum.get("sample_name")
        wavelengths = spectrum.get("wavelengths_nm")
        intensities = spectrum.get("intensities")
        replicates = spectrum.get("replicate_count")
        if not isinstance(name, str) or not name.strip() or len(name) > 200:
            raise RequestError(400, "Each spectrum needs a valid sample_name.")
        if not isinstance(wavelengths, list) or not isinstance(intensities, list):
            raise RequestError(400, f"{name}: wavelengths_nm and intensities must be arrays.")
        if not wavelengths or len(wavelengths) != len(intensities):
            raise RequestError(400, f"{name}: wavelength and intensity arrays must have equal non-zero lengths.")
        if len(wavelengths) > MAX_POINTS_PER_SPECTRUM:
            raise RequestError(413, f"{name}: too many measurement points.")
        if type(replicates) is not int or replicates < 1:
            raise RequestError(400, f"{name}: replicate_count must be a positive integer.")
        for value in wavelengths + intensities:
            try:
                finite = type(value) in (int, float) and math.isfinite(value)
            except OverflowError:
                finite = False
            if not finite:
                raise RequestError(400, f"{name}: measurements must contain finite numbers.")


class ChiliScanHandler(BaseHTTPRequestHandler):
    server_version = "ChiliScanJetsonAPI/1.0"
    sys_version = ""

    def _send_json(self, status, payload):
        body = json.dumps(payload, ensure_ascii=False, allow_nan=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        origin = self.headers.get("Origin")
        if origin == ALLOWED_ORIGIN:
            self.send_header("Access-Control-Allow-Origin", ALLOWED_ORIGIN)
            self.send_header("Vary", "Origin")
        self.end_headers()
        self.wfile.write(body)

    def _origin_allowed(self):
        origin = self.headers.get("Origin")
        if origin and origin != ALLOWED_ORIGIN:
            self._send_json(403, {"error": "origin_not_allowed"})
            return False
        return True

    def _authorized(self):
        scheme, separator, supplied = self.headers.get("Authorization", "").partition(" ")
        return bool(
            TOKEN
            and separator
            and scheme.lower() == "bearer"
            and hmac.compare_digest(supplied.encode("utf-8"), TOKEN.encode("utf-8"))
        )

    def _require_auth(self):
        if not self._authorized():
            self._send_json(401, {"error": "unauthorized"})
            return False
        return True

    def do_OPTIONS(self):
        if not self._origin_allowed():
            return
        if urlsplit(self.path).path not in ("/api/v1/health", "/api/v1/predict"):
            self._send_json(404, {"error": "not_found"})
            return
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", ALLOWED_ORIGIN)
        self.send_header("Vary", "Origin")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Authorization, Content-Type")
        self.send_header("Access-Control-Max-Age", "600")
        self.end_headers()

    def do_GET(self):
        if not self._origin_allowed():
            return
        if urlsplit(self.path).path != "/api/v1/health":
            self._send_json(404, {"error": "not_found"})
            return
        if not self._require_auth():
            return
        self._send_json(200, {
            "api_version": 1,
            "model_ready": False,
            "model_name": None,
        })

    def do_POST(self):
        if not self._origin_allowed():
            return
        if urlsplit(self.path).path != "/api/v1/predict":
            self._send_json(404, {"error": "not_found"})
            return
        if not self._require_auth():
            return

        try:
            content_length = int(self.headers.get("Content-Length", ""))
        except ValueError:
            self._send_json(400, {"error": "invalid_content_length"})
            return
        if content_length <= 0 or content_length > MAX_BODY_BYTES:
            self._send_json(413, {"error": "request_size_not_allowed"})
            return

        try:
            payload = json.loads(self.rfile.read(content_length))
            validate_request(payload)
        except (json.JSONDecodeError, UnicodeDecodeError):
            self._send_json(400, {"error": "invalid_json"})
            return
        except RequestError as error:
            self._send_json(error.status, {"error": str(error)})
            return

        self._send_json(503, {
            "error": "model_not_ready",
            "message": "No trained classification model is configured.",
        })


def main():
    if not TOKEN:
        raise SystemExit("Set CHILISCAN_TOKEN to a private, random access token before starting the API.")
    if not ALLOWED_ORIGIN:
        raise SystemExit("CHILISCAN_ORIGIN must be the exact ChiliScan website origin.")
    server = ThreadingHTTPServer((HOST, PORT), ChiliScanHandler)
    print(f"ChiliScan Jetson API listening on {HOST}:{PORT}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
