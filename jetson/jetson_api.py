#!/usr/bin/env python3
"""Minimal authenticated API scaffold for ChiliScan and Jetson."""

import hmac
import json
import math
import os
import urllib.error
import urllib.request
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlsplit


HOST = os.environ.get("CHILISCAN_HOST", "127.0.0.1")
PORT = int(os.environ.get("CHILISCAN_PORT", os.environ.get("PORT", "8765")))
TOKEN = os.environ.get("CHILISCAN_TOKEN", "")
ALLOWED_ORIGIN = os.environ.get("CHILISCAN_ORIGIN", "https://rgwvb.github.io")
NVIDIA_API_KEY = os.environ.get("NVIDIA_API_KEY", "")
NVIDIA_KUMO_PREDICT_URL = os.environ.get("NVIDIA_KUMO_PREDICT_URL", "")
MAX_BODY_BYTES = 2_000_000
MAX_SPECTRA = 100
MAX_POINTS_PER_SPECTRUM = 20_000
MAX_KUMO_FEATURE_COLUMNS = 498
LABELS = frozenset({"fresh", "pest", "rotten"})


class RequestError(Exception):
    def __init__(self, status, message):
        super().__init__(message)
        self.status = status


def validate_spectra(spectra, *, labelled):
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
        if any(value < 200 or value > 1100 for value in wavelengths):
            raise RequestError(400, f"{name}: wavelengths must be between 200 and 1100 nm.")
        if any(current <= previous for previous, current in zip(wavelengths, wavelengths[1:])):
            raise RequestError(400, f"{name}: wavelengths must be strictly increasing.")
        if labelled and spectrum.get("label") not in LABELS:
            raise RequestError(400, f"{name}: label must be fresh, pest, or rotten.")


def build_kumo_payload(context, spectra):
    reference_wavelengths = context[0]["wavelengths_nm"]
    for sample in context + spectra:
        if sample["wavelengths_nm"] != reference_wavelengths:
            raise RequestError(400, "All context and query spectra must use the same wavelength positions.")

    sample_ids = [sample["sample_name"] for sample in context + spectra]
    if len(sample_ids) != len(set(sample_ids)):
        raise RequestError(400, "Sample names must be unique across context and query spectra.")

    context_labels = {sample["label"] for sample in context}
    if context_labels != LABELS:
        raise RequestError(400, "The labelled context must include fresh, pest, and rotten samples.")

    id_column = "sample_id"
    feature_indexes = list(range(len(reference_wavelengths)))
    if len(feature_indexes) > MAX_KUMO_FEATURE_COLUMNS:
        last_index = len(feature_indexes) - 1
        feature_indexes = [
            round(index * last_index / (MAX_KUMO_FEATURE_COLUMNS - 1))
            for index in range(MAX_KUMO_FEATURE_COLUMNS)
        ]
    feature_columns = [f"wavelength_{index:04d}" for index in feature_indexes]
    target_column = "quality"
    columns = [id_column] + feature_columns + [target_column]
    values = {
        column: {"dtype": "float64", "stype": "numerical"}
        for column in feature_columns
    }
    values[id_column] = {"dtype": "string", "stype": "ID", "nullable": False}
    values[target_column] = {"dtype": "string", "stype": "categorical"}

    return {
        "model": "kumo-tabular",
        "task": {
            "kind": "classification",
            "target": {
                "column_name": target_column,
                "dtype": "string",
                "classes": sorted(LABELS),
            },
        },
        "schema": {
            "instance_table": {"columns": values, "primary_key": id_column},
            "related_tables": {},
            "relationships": [],
        },
        "context": {
            "instance_table": {
                "format": "arrays",
                "columns": columns,
                "rows": [
                    [
                        sample["sample_name"],
                        *(sample["intensities"][index] for index in feature_indexes),
                        sample["label"],
                    ]
                    for sample in context
                ],
            },
            "related_tables": {},
        },
        "predict": {
            "instance_table": {
                "format": "arrays",
                "columns": [id_column] + feature_columns,
                "rows": [
                    [
                        sample["sample_name"],
                        *(sample["intensities"][index] for index in feature_indexes),
                    ]
                    for sample in spectra
                ],
            },
            "related_tables": {},
        },
        "output": {"fields": ["prediction", "probabilities"]},
    }


class NoRedirectHandler(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, request, response, code, message, headers, new_url):
        return None


def request_kumo(payload):
    parsed = urlsplit(NVIDIA_KUMO_PREDICT_URL)
    if parsed.scheme != "https" or parsed.hostname != "ai.api.nvidia.com" or not parsed.path:
        raise RequestError(503, "Configure NVIDIA_KUMO_PREDICT_URL with the HTTPS NVIDIA API Catalog prediction URL.")

    request = urllib.request.Request(
        NVIDIA_KUMO_PREDICT_URL,
        data=json.dumps(payload, allow_nan=False).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {NVIDIA_API_KEY}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        },
        method="POST",
    )
    opener = urllib.request.build_opener(NoRedirectHandler())
    try:
        with opener.open(request, timeout=60) as response:
            body = response.read(MAX_BODY_BYTES + 1)
            if len(body) > MAX_BODY_BYTES:
                raise RequestError(502, "NVIDIA returned an oversized response.")
    except urllib.error.HTTPError as error:
        if 300 <= error.code < 400:
            raise RequestError(502, "NVIDIA returned an unexpected redirect.") from None
        raise RequestError(502, f"NVIDIA Kumo request failed (HTTP {error.code}).") from None
    except (urllib.error.URLError, TimeoutError, OSError):
        raise RequestError(502, "Could not reach the NVIDIA Kumo endpoint.") from None

    try:
        result = json.loads(body)
    except (json.JSONDecodeError, UnicodeDecodeError):
        raise RequestError(502, "NVIDIA returned invalid JSON.") from None

    predictions = result.get("predictions") if isinstance(result, dict) else None
    if not isinstance(predictions, list) or len(predictions) != len(payload["predict"]["instance_table"]["rows"]):
        raise RequestError(502, "NVIDIA returned an unexpected prediction response.")

    normalized = []
    seen = set()
    for item in predictions:
        if not isinstance(item, dict):
            raise RequestError(502, "NVIDIA returned an invalid prediction item.")
        index = item.get("row_index")
        label = item.get("prediction")
        probabilities = item.get("probabilities")
        if type(index) is not int or index < 0 or index >= len(predictions) or index in seen:
            raise RequestError(502, "NVIDIA returned invalid prediction row indexes.")
        if not isinstance(label, str) or label not in LABELS or not isinstance(probabilities, dict):
            raise RequestError(502, "NVIDIA returned an invalid classification.")
        scores = [value for value in probabilities.values() if type(value) in (float, int) and math.isfinite(value)]
        if len(scores) != len(probabilities) or not scores or any(value < 0 or value > 1 for value in scores):
            raise RequestError(502, "NVIDIA returned invalid class probabilities.")
        if label not in probabilities:
            raise RequestError(502, "NVIDIA omitted the predicted class probability.")
        seen.add(index)
        normalized.append((index, label, probabilities[label]))

    if len(seen) != len(predictions):
        raise RequestError(502, "NVIDIA omitted one or more prediction rows.")
    return normalized


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
        configured = bool(NVIDIA_API_KEY and NVIDIA_KUMO_PREDICT_URL)
        self._send_json(200, {
            "api_version": 1,
            "model_ready": configured,
            "model_name": "NVIDIA Kumo-Tabular" if configured else None,
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
            if not isinstance(payload, dict) or payload.get("api_version") != 1:
                raise RequestError(400, "Expected api_version 1.")
            if not NVIDIA_API_KEY or not NVIDIA_KUMO_PREDICT_URL:
                raise RequestError(503, "Configure the NVIDIA API key and prediction URL on the server.")
            context = payload.get("context_spectra")
            spectra = payload.get("spectra")
            validate_spectra(context, labelled=True)
            validate_spectra(spectra, labelled=False)
            kumo_payload = build_kumo_payload(context, spectra)
            predictions = request_kumo(kumo_payload)
        except (json.JSONDecodeError, UnicodeDecodeError):
            self._send_json(400, {"error": "invalid_json"})
            return
        except RequestError as error:
            self._send_json(error.status, {"error": str(error)})
            return

        self._send_json(200, {
            "api_version": 1,
            "predictions": [
                {
                    "sample_name": spectra[index]["sample_name"],
                    "label": label,
                    "confidence": confidence,
                }
                for index, label, confidence in predictions
            ],
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
