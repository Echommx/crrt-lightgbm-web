#!/usr/bin/env python3
"""Loopback-only LightGBM research UI.

This launcher deliberately lives outside the source project.  It imports the
project's frozen LightGBM_trend bundle at runtime and never trains, writes to,
or edits project files. Excel and the documented long-form tabular upload
format are normalized to the project's existing ``static`` + ``observations``
input contract before inference. CSV and JSON remain compatibility/API paths;
the visible browser workflow uses a simple one-sheet XLSX workbook.
"""
from __future__ import annotations

import argparse
import csv
import io
import json
import math
import os
import secrets
import sys
import threading
import webbrowser
from datetime import date, datetime, time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import urlsplit

MODEL_NAME = "LightGBM_trend"
MAX_BODY_BYTES = 2_000_000
MAX_OBSERVATIONS = 10_000
STATIC_KEYS = (
    "event_id", "patient_id", "encounter_id", "t0", "start_time", "end_time",
    "age", "sex", "history_diabetes", "history_ckd", "history_chronic_liver",
    "history_malignancy", "history_chronic_lung", "history_heart_failure",
    "specimen_site", "antibiotic_exposure_7d", "antibiotic_classes_7d",
    "carbapenem_exposure_7d", "invasive_ventilation_t0",
)
OBSERVATION_KEYS = ("measured_at", "feature", "value", "unit", "available_at")
CSV_COLUMNS = (*STATIC_KEYS, *OBSERVATION_KEYS)
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
NUMERIC_KEYS = {
    "age", "antibiotic_classes_7d", "history_diabetes", "history_ckd",
    "history_chronic_liver", "history_malignancy", "history_chronic_lung",
    "history_heart_failure", "antibiotic_exposure_7d", "carbapenem_exposure_7d",
    "invasive_ventilation_t0", "value",
}


class InputError(ValueError):
    """User-facing malformed input error."""


def _no_constant(value: str) -> None:
    raise InputError(f"Non-finite JSON number is not allowed: {value}")


def _blank(value: Any) -> bool:
    return value is None or str(value).strip() == ""


def _convert_csv_value(key: str, value: Any) -> Any:
    if _blank(value):
        return None
    value = str(value).strip()
    if key in NUMERIC_KEYS:
        try:
            number = float(value)
        except ValueError as exc:
            raise InputError(f"CSV column {key!r} must be numeric; got {value!r}.") from exc
        if not math.isfinite(number):
            raise InputError(f"CSV column {key!r} must be finite.")
        return int(number) if number.is_integer() else number
    return value


def payload_to_csv(payload: dict[str, Any]) -> str:
    """Serialize the contract payload to the documented long-form CSV."""
    static = payload.get("static")
    observations = payload.get("observations")
    if not isinstance(static, dict) or not isinstance(observations, list):
        raise InputError("Payload must contain static and observations fields.")
    output = io.StringIO(newline="")
    writer = csv.DictWriter(output, fieldnames=CSV_COLUMNS, lineterminator="\n")
    writer.writeheader()
    for row in observations:
        record = {key: static.get(key, "") for key in STATIC_KEYS}
        record.update({key: row.get(key, "") for key in OBSERVATION_KEYS})
        writer.writerow(record)
    return output.getvalue()


def csv_to_payload(text: str) -> dict[str, Any]:
    """Parse one row per observation, with static fields on the first row.

    Static columns may be repeated on each row or left blank after the first
    row.  The parser takes the first non-empty value for each static field,
    then converts numeric columns without guessing units or timestamps.
    """
    if len(text.encode("utf-8")) > MAX_BODY_BYTES:
        raise InputError("CSV input must not exceed 2 MB.")
    try:
        rows = list(csv.DictReader(io.StringIO(text.lstrip("\ufeff"))))
    except csv.Error as exc:
        raise InputError(f"Invalid CSV: {exc}") from exc
    if not rows:
        raise InputError("CSV must contain a header and at least one observation row.")
    actual = set(rows[0])
    missing = sorted(set(OBSERVATION_KEYS[:4]) - actual)
    if missing:
        raise InputError("CSV is missing required observation columns: " + ", ".join(missing))
    # Ignore harmless extra columns, but reject duplicate/empty headers because
    # csv.DictReader would silently discard data in those cases.
    headers = list(rows[0])
    if any(not h or h is None for h in headers):
        raise InputError("CSV headers must be non-empty.")
    static: dict[str, Any] = {}
    for key in STATIC_KEYS:
        value = next((row.get(key) for row in rows if not _blank(row.get(key))), None)
        if value is not None:
            static[key] = _convert_csv_value(key, value)
    if _blank(static.get("specimen_site")):
        static["specimen_site"] = "unknown"
    observations: list[dict[str, Any]] = []
    for index, row in enumerate(rows, start=2):
        if all(_blank(row.get(key)) for key in OBSERVATION_KEYS):
            continue
        missing_row = [key for key in OBSERVATION_KEYS[:4] if _blank(row.get(key))]
        if missing_row:
            raise InputError(f"CSV row {index} is missing: {', '.join(missing_row)}.")
        observation = {
            key: _convert_csv_value(key, row.get(key))
            for key in OBSERVATION_KEYS
            if not _blank(row.get(key))
        }
        observations.append(observation)
    if not observations:
        raise InputError("CSV contains no non-empty observations.")
    return {"static": static, "observations": observations}


def _xlsx_cell_text(value: Any) -> Any:
    """Convert openpyxl date cells to the ISO strings used by the contract."""
    if isinstance(value, datetime):
        return value.isoformat(sep=" ")
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, time):
        return value.isoformat()
    return value


def xlsx_to_payload(raw: bytes) -> dict[str, Any]:
    """Parse the first worksheet of a simple long-table XLSX example.

    Excel is used only as an editable transport format. The resulting payload
    is the same static-plus-observations contract used by JSON and CSV, and the
    workbook bytes are not written to disk.
    """
    if len(raw) > MAX_BODY_BYTES:
        raise InputError("Excel input must not exceed 2 MB.")
    try:
        from openpyxl import load_workbook
    except ImportError as exc:
        raise InputError("Excel upload requires openpyxl. Install requirements-inference.txt.") from exc
    try:
        workbook = load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
        sheet = workbook[workbook.sheetnames[0]]
        iterator = sheet.iter_rows(values_only=True)
        header_row = next(iterator, None)
        if not header_row:
            raise InputError("Excel must contain a header row and at least one observation row.")
        headers = [str(_xlsx_cell_text(value)).strip() if value is not None else "" for value in header_row]
        if any(not header for header in headers) or len(set(headers)) != len(headers):
            raise InputError("Excel headers must be unique and non-empty.")
        rows = []
        for values in iterator:
            row = {}
            for index, header in enumerate(headers):
                value = values[index] if index < len(values) else None
                row[header] = _xlsx_cell_text(value)
            rows.append(row)
    except InputError:
        raise
    except Exception as exc:
        raise InputError(f"Invalid Excel workbook: {exc}") from exc
    if not rows:
        raise InputError("Excel must contain at least one observation row.")
    # Reuse the exact CSV contract conversion and validation. DictWriter is not
    # used here, so dates and numbers remain easy to inspect in the workbook.
    output = io.StringIO(newline="")
    writer = csv.DictWriter(output, fieldnames=headers, extrasaction="ignore", lineterminator="\n")
    writer.writeheader()
    writer.writerows(rows)
    return csv_to_payload(output.getvalue())


def _minimal_contract_check(payload: Any) -> None:
    """Cheap check used before project validation for clearer errors."""
    if not isinstance(payload, dict):
        raise InputError("JSON input must be an object.")
    if not isinstance(payload.get("static"), dict):
        raise InputError("JSON input must contain a static object.")
    if not isinstance(payload.get("observations"), list):
        raise InputError("JSON input must contain an observations array.")
    if not 1 <= len(payload["observations"]) <= MAX_OBSERVATIONS:
        raise InputError(f"Provide 1–{MAX_OBSERVATIONS} observations.")


class LightGBMEngine:
    def __init__(self, project_root: Path):
        self.root = Path(project_root).resolve()
        if not (self.root / "configs" / "models.json").is_file():
            raise RuntimeError(f"Not a CRRT project root (missing configs/models.json): {self.root}")
        # The minimal GitHub release keeps the inference modules in
        # runtime_bundle.zip. Add it before importing the same project modules
        # and expose the real release root so config/model paths still resolve
        # correctly when the package is imported from a ZIP archive.
        os.environ["CRRT_WEB_ROOT"] = str(self.root)
        runtime_bundle = Path(__file__).resolve().parent / "runtime_bundle.zip"
        if runtime_bundle.is_file() and str(runtime_bundle) not in sys.path:
            sys.path.insert(0, str(runtime_bundle))
        # The source project's imports use crrt_clean.storage.ROOT, which is
        # derived from this explicit release root rather than the ZIP path.
        if str(self.root) not in sys.path:
            sys.path.insert(0, str(self.root))
        from crrt_clean.common import load_json  # imported after sys.path setup
        from paperlib.web_predict import example_payload, predict_payload
        from crrt_clean.models import load_bundle
        from crrt_clean.storage import require_finished
        self._load_json = load_json
        self._example_payload = example_payload
        self._predict_payload = predict_payload
        self._load_bundle = load_bundle
        self._require_finished = require_finished
        # A source project keeps models under outputs/models; the sanitized
        # public package keeps only the inference artifact under models/.
        packaged_dir = self.root / "models" / MODEL_NAME
        source_dir = self.root / "outputs" / "models" / MODEL_NAME
        self.model_dir = packaged_dir if (packaged_dir / "model.joblib").is_file() else source_dir
        registry = load_json(self.root / "configs" / "models.json")
        if MODEL_NAME not in registry:
            raise RuntimeError(f"{MODEL_NAME} is not registered in configs/models.json.")
        self.registry = registry[MODEL_NAME]
        self.bundle: Any = None
        self.lock = threading.Lock()
        self.available = (self.model_dir / "COMPLETE.json").is_file() and (self.model_dir / "model.joblib").is_file()
        self.threshold = 0.5
        self.threshold_origin = "0.50 research threshold; not clinically validated"
        threshold_file = self.model_dir / "thresholds.json"
        if threshold_file.is_file():
            try:
                saved = load_json(threshold_file).get("mimic_apparent_sensitivity_target")
                if isinstance(saved, (int, float)) and math.isfinite(float(saved)) and 0.0 < float(saved) < 1.0:
                    self.threshold = float(saved)
                    self.threshold_origin = "Saved MIMIC training-derived research threshold; not independently validated"
            except Exception:
                # A malformed optional threshold must not prevent the score from
                # being shown; the explicit 0.50 research fallback remains.
                pass
        # Prefer the project validator; this preserves the established leakage,
        # unit, timestamp and range checks exactly for valid project installs.
        try:
            from webapp.input_contract import validate_payload
            self._validate_payload = validate_payload
        except Exception:
            self._validate_payload = None

    def config(self) -> dict[str, Any]:
        example = self._example_payload()
        return {
            "model": MODEL_NAME,
            "parameter_id": self.registry.get("parameter_id"),
            "available": bool(self.available),
            "example": example,
            "example_csv": payload_to_csv(example),
            "example_xlsx_url": "/api/example.xlsx",
            "csv_columns": list(CSV_COLUMNS),
            "upload_format": "xlsx",
            "threshold": self.threshold,
            "threshold_origin": self.threshold_origin,
            "max_body_bytes": MAX_BODY_BYTES,
            "max_observations": MAX_OBSERVATIONS,
            "disclaimer": "Research use only. The result is a study-defined high/low risk classification, not a diagnosis or treatment recommendation.",
        }

    def predict(self, payload: dict[str, Any]) -> dict[str, Any]:
        _minimal_contract_check(payload)
        if self._validate_payload is not None:
            self._validate_payload(payload, MAX_OBSERVATIONS)
        if not self.available:
            raise RuntimeError(
                f"Frozen LightGBM bundle is unavailable at {self.model_dir}. "
                "Complete/export LightGBM_trend before running inference."
            )
        with self.lock:
            if self.bundle is None:
                self._require_finished(self.model_dir)
                self.bundle = self._load_bundle(self.model_dir)
            score, audit = self._predict_payload(self.bundle, payload)
        return {
            "model": MODEL_NAME,
            "parameter_id": self.bundle.get("parameter_id", self.registry.get("parameter_id")),
            "score": float(score),
            "binary_prediction": int(float(score) >= self.threshold),
            "risk_label": "High risk" if float(score) >= self.threshold else "Low risk",
            "threshold": self.threshold,
            "threshold_origin": self.threshold_origin,
            "audit": audit,
            "research_only": True,
        }


def create_server(project_root: Path, static_root: Path, port: int = 8501) -> ThreadingHTTPServer:
    engine = LightGBMEngine(project_root)
    token = secrets.token_urlsafe(32)
    static_root = Path(static_root).resolve()
    example_xlsx = static_root.parent / "examples" / "simulated_input.xlsx"
    allowed_files = {
        "/": ("index.html", "text/html; charset=utf-8"),
        "/index.html": ("index.html", "text/html; charset=utf-8"),
        "/app.js": ("app.js", "application/javascript; charset=utf-8"),
        "/style.css": ("style.css", "text/css; charset=utf-8"),
    }

    class Handler(BaseHTTPRequestHandler):
        server_version = "LightGBMResearchUI/1.0"

        def log_message(self, *_args: Any) -> None:
            return

        def _safe_host(self) -> bool:
            value = self.headers.get("Host", "")
            p = self.server.server_port
            return value in {f"127.0.0.1:{p}", f"localhost:{p}"}

        def _send(self, code: int, body: Any, content_type: str = "application/json") -> None:
            if isinstance(body, bytes):
                raw = body
            elif isinstance(body, str):
                raw = body.encode("utf-8")
            else:
                raw = json.dumps(body, ensure_ascii=False, allow_nan=False).encode("utf-8")
            self.send_response(code)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(raw)))
            self.send_header("Cache-Control", "no-store")
            self.send_header("X-Content-Type-Options", "nosniff")
            self.send_header("X-Frame-Options", "DENY")
            self.send_header("Content-Security-Policy", "default-src 'self'; style-src 'self'; script-src 'self'; connect-src 'self'; img-src 'self' data:; frame-ancestors 'none'; base-uri 'none'")
            self.end_headers()
            self.wfile.write(raw)

        def do_GET(self) -> None:
            if not self._safe_host():
                return self._send(403, {"error": "Local host only."})
            path = urlsplit(self.path).path
            if path in allowed_files:
                name, mime = allowed_files[path]
                file = static_root / name
                if not file.is_file():
                    return self._send(404, {"error": "Static file missing."})
                return self._send(200, file.read_bytes(), mime)
            if path == "/api/example.xlsx":
                if not example_xlsx.is_file():
                    return self._send(404, {"error": "Excel example missing."})
                raw = example_xlsx.read_bytes()
                self.send_response(200)
                self.send_header("Content-Type", XLSX_MIME)
                self.send_header("Content-Disposition", 'attachment; filename="simulated_input.xlsx"')
                self.send_header("Content-Length", str(len(raw)))
                self.send_header("Cache-Control", "no-store")
                self.send_header("X-Content-Type-Options", "nosniff")
                self.end_headers()
                self.wfile.write(raw)
                return
            if path == "/api/config":
                return self._send(200, {**engine.config(), "token": token})
            if path == "/health":
                return self._send(200, {"status": "ok", "model": MODEL_NAME, "available": engine.available})
            return self._send(404, {"error": "Not found."})

        def do_POST(self) -> None:
            if not self._safe_host():
                return self._send(403, {"error": "Local host only."})
            if self.path not in {"/api/predict", "/api/parse-xlsx"}:
                return self._send(404, {"error": "Not found."})
            origin = self.headers.get("Origin")
            allowed = {f"http://127.0.0.1:{self.server.server_port}", f"http://localhost:{self.server.server_port}"}
            if origin and origin not in allowed:
                return self._send(403, {"error": "Cross-origin requests are disabled."})
            if self.headers.get("X-CRRT-Token") != token:
                return self._send(403, {"error": "Reload this local page before predicting."})
            try:
                length = int(self.headers.get("Content-Length", "0"))
                if not 0 < length <= MAX_BODY_BYTES:
                    raise InputError("Input must be between 1 byte and 2 MB.")
                content_type = self.headers.get("Content-Type", "")
                raw = self.rfile.read(length)
                if self.path == "/api/parse-xlsx":
                    if not content_type.startswith(XLSX_MIME):
                        raise InputError("Use an .xlsx workbook for Excel upload.")
                    return self._send(200, {"payload": xlsx_to_payload(raw)})
                if not content_type.startswith("application/json"):
                    raise InputError("Use application/json for prediction requests.")
                data = json.loads(raw, parse_constant=_no_constant)
                payload = data.get("payload", data) if isinstance(data, dict) else data
                out = engine.predict(payload)
                return self._send(200, out)
            except Exception as exc:
                return self._send(400, {"error": str(exc)})

    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    server.engine = engine  # type: ignore[attr-defined]
    return server


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--project-root", type=Path, required=False, default=Path.cwd(), help="CRRT project root containing configs/ and outputs/")
    parser.add_argument("--port", type=int, default=8501)
    parser.add_argument("--no-browser", action="store_true")
    parser.add_argument("--self-test", action="store_true", help="Validate package/config/CSV normalization without loading weights")
    args = parser.parse_args(argv)
    if not 1 <= args.port <= 65535:
        parser.error("--port must be between 1 and 65535")
    package_root = Path(__file__).resolve().parent
    engine = LightGBMEngine(args.project_root)
    cfg = engine.config()
    parsed = csv_to_payload(cfg["example_csv"])
    if args.self_test:
        _minimal_contract_check(parsed)
        if parsed["observations"] != cfg["example"]["observations"]:
            raise RuntimeError("CSV round-trip changed illustrative observations")
        xlsx_path = package_root / "examples" / "simulated_input.xlsx"
        xlsx_rows = None
        if xlsx_path.is_file():
            parsed_xlsx = xlsx_to_payload(xlsx_path.read_bytes())
            _minimal_contract_check(parsed_xlsx)
            if parsed_xlsx["observations"] != cfg["example"]["observations"]:
                raise RuntimeError("Excel round-trip changed illustrative observations")
            xlsx_rows = len(parsed_xlsx["observations"])
        print(json.dumps({"self_test": "PASS", "model": MODEL_NAME, "bundle_available": cfg["available"], "csv_rows": len(parsed["observations"]), "xlsx_rows": xlsx_rows}, ensure_ascii=False))
        return 0
    server = create_server(args.project_root, package_root / "static", args.port)
    url = f"http://127.0.0.1:{server.server_port}"
    print(f"LightGBM-only research UI: {url}", flush=True)
    print("Excel uploads remain local; no training or patient data are written.", flush=True)
    if not args.no_browser:
        timer = threading.Timer(0.6, lambda: webbrowser.open(url))
        timer.daemon = True
        timer.start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        return 0
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
