"""HTTP API checks for the hebdo endpoints (and the untouched converter ones).

Run while the server is up:  python api_check_hebdo.py
"""

from __future__ import annotations

import json
import os
import time
import urllib.error
import urllib.request
import uuid
import zipfile
from io import BytesIO
from pathlib import Path

# Override with e.g. HEBDO_API_BASE=http://127.0.0.1:8010 when 8000 is taken.
BASE = os.environ.get("HEBDO_API_BASE", "http://127.0.0.1:8000")
PROJECT = Path(__file__).resolve().parent     # relocatable: follows the folder
BRISE = (PROJECT / "sample-data" / "brise_smc_week39.csv").read_bytes()
GNOC = (PROJECT / "sample-data" / "GNOC_week39.csv").read_bytes()
SAMPLE_CSV = (PROJECT / "sample-data" / "employees.csv").read_bytes()

FAILURES: list[str] = []
CHECKS = 0


def check(condition, message: str) -> None:
    global CHECKS
    CHECKS += 1
    if not condition:
        FAILURES.append(message)


def get(path: str):
    req = urllib.request.Request(BASE + path)
    try:
        with urllib.request.urlopen(req, timeout=30) as res:
            return res.status, {k.lower(): v for k, v in res.headers.items()}, res.read()
    except urllib.error.HTTPError as exc:
        return exc.code, {k.lower(): v for k, v in exc.headers.items()}, exc.read()


def post_multipart(path: str, fields: dict, files: dict):
    boundary = uuid.uuid4().hex
    body = BytesIO()
    for key, value in fields.items():
        body.write(f"--{boundary}\r\nContent-Disposition: form-data; name=\"{key}\"\r\n\r\n"
                   f"{value}\r\n".encode())
    for name, (filename, data) in files.items():
        body.write(f"--{boundary}\r\n"
                   f"Content-Disposition: form-data; name=\"{name}\"; filename=\"{filename}\"\r\n"
                   f"Content-Type: text/csv\r\n\r\n".encode())
        body.write(data)
        body.write(b"\r\n")
    body.write(f"--{boundary}--\r\n".encode())
    req = urllib.request.Request(
        BASE + path, data=body.getvalue(),
        headers={"Content-Type": f"multipart/form-data; boundary={boundary}"})
    try:
        with urllib.request.urlopen(req, timeout=120) as res:
            return res.status, {k.lower(): v for k, v in res.headers.items()}, res.read()
    except urllib.error.HTTPError as exc:
        return exc.code, {k.lower(): v for k, v in exc.headers.items()}, exc.read()


def wait_for_server(timeout: float = 30.0) -> bool:
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            status, _, _ = get("/api/health")
            if status == 200:
                return True
        except Exception:  # noqa: BLE001
            pass
        time.sleep(0.4)
    return False


def main() -> None:
    if not wait_for_server():
        print("FAILED: server not reachable at", BASE)
        raise SystemExit(1)

    # ---- pages ------------------------------------------------------------
    status, headers, body = get("/")
    check(status == 200 and b"Weekly report" in body, "index page + nav link to /hebdo")
    status, headers, body = get("/hebdo")
    check(status == 200 and b"hebdo_SXX" in body, "hebdo page served")
    check("text/html" in headers.get("content-type", ""), "hebdo page content-type")

    # ---- converter endpoints untouched ------------------------------------
    status, _, body = get("/api/health")
    check(status == 200, "legacy /api/health")
    status, _, body = get("/api/options")
    check(status == 200 and json.loads(body)["options"]["encoding"] == "auto",
          "legacy /api/options")
    status, headers, body = post_multipart(
        "/api/preview", {"options": "{}"}, {"file": ("employees.csv", SAMPLE_CSV)})
    check(status == 200 and json.loads(body)["rows"] > 0, "legacy /api/preview")

    # ---- hebdo preview -----------------------------------------------------
    status, headers, body = post_multipart(
        "/api/hebdo/preview",
        {"week": "S39"},
        {"brise": ("brise_smc_week39.csv", BRISE), "gnoc": ("GNOC_week39.csv", GNOC)})
    check(status == 200, f"hebdo preview status {status}")
    metrics = json.loads(body) if status == 200 else {}
    check(metrics.get("week") == 39 and metrics.get("week_label") == "S39", "preview week S39")
    check(metrics.get("rows", {}).get("IT", {}).get("tt") == 9, "preview IT tt = 9")
    check(metrics.get("rows", {}).get("Core PS", {}).get("mttr_ok", {}).get("P1") is False,
          "preview Core PS P1 over target")
    check(metrics.get("rows", {}).get("IT", {}).get("mttr", {}).get("P1") == "04:10",
          "preview MTTR as HH:MM")
    check(metrics.get("rows", {}).get("IT", {}).get("mttr", {}).get("P2") == "33:00",
          "preview MTTR beyond 24 h (33:00)")
    core_ip = metrics.get("rows", {}).get("Core IP", {}).get("rates", {})
    check(core_ip.get("P3") == 0.0, "preview: evaluated zero stays 0.00%")
    check(core_ip.get("P4") is None, "preview: unevaluated SLA stays empty (null, not 0)")
    check(len(metrics.get("row_order", [])) == 5, "preview row order")

    # preview resolving from a date instead of a week
    status, headers, body = post_multipart(
        "/api/hebdo/preview", {"date": "2026-09-22"},
        {"brise": ("b.csv", BRISE), "gnoc": ("g.csv", GNOC)})
    check(status == 200 and json.loads(body).get("week") == 39, "preview date -> ISO week")

    # column overrides
    status, headers, body = post_multipart(
        "/api/hebdo/preview", {"week": "S39", "columns": json.dumps({"priority": "prio"})},
        {"brise": ("b.csv", BRISE), "gnoc": ("g.csv", GNOC)})
    check(status == 400 and "prio" in json.loads(body)["detail"],
          f"bad column override -> 400 ({status})")

    # ---- error paths -------------------------------------------------------
    status, headers, body = post_multipart(
        "/api/hebdo/preview", {"week": "S99"},
        {"brise": ("b.csv", BRISE), "gnoc": ("g.csv", GNOC)})
    check(status == 400 and "week" in json.loads(body)["detail"].lower(),
          f"bad week -> 400 ({status}): {body[:120]!r}")

    status, headers, body = post_multipart(
        "/api/hebdo/preview", {"week": "S39"},
        {"brise": ("b.csv", BRISE)})
    check(status == 422, f"missing gnoc file -> 422 ({status})")

    status, headers, body = post_multipart(
        "/api/hebdo/preview", {"week": "S39"},
        {"brise": ("b.csv", b""), "gnoc": ("g.csv", GNOC)})
    check(status == 400 and "empty" in json.loads(body)["detail"],
          f"empty file -> 400 ({status})")

    # ---- generate: xlsx / pptx / zip ---------------------------------------
    status, headers, body = post_multipart(
        "/api/hebdo/generate", {"week": "S39", "output": "xlsx"},
        {"brise": ("b.csv", BRISE), "gnoc": ("g.csv", GNOC)})
    cd = headers.get("content-disposition", "")
    check(status == 200 and 'filename="hebdo_S39.xlsx"' in cd, f"generate xlsx ({status}) {cd}")
    check("spreadsheetml" in headers.get("content-type", ""), "xlsx media type")
    if status == 200:
        import openpyxl
        wb = openpyxl.load_workbook(BytesIO(body))
        check(wb["Tickets"]["B6"].value == 9, "generated xlsx IT tt = 9")
        check(wb["MTTR P1"]["B6"].value == "04:10", "generated xlsx MTTR P1 IT (HH:MM)")
        check(wb["MTTR P2"]["B6"].value == "33:00", "generated xlsx MTTR P2 IT (>24 h)")

    status, headers, body = post_multipart(
        "/api/hebdo/generate", {"week": "S39", "output": "pptx"},
        {"brise": ("b.csv", BRISE), "gnoc": ("g.csv", GNOC)})
    cd = headers.get("content-disposition", "")
    check(status == 200 and 'filename="hebdo_S39.pptx"' in cd, f"generate pptx ({status}) {cd}")
    check("presentationml" in headers.get("content-type", ""), "pptx media type")

    status, headers, body = post_multipart(
        "/api/hebdo/generate", {"week": "S39", "output": "both"},
        {"brise": ("b.csv", BRISE), "gnoc": ("g.csv", GNOC)})
    cd = headers.get("content-disposition", "")
    check(status == 200 and 'filename="hebdo_S39.zip"' in cd, f"generate zip ({status}) {cd}")
    check("zip" in headers.get("content-type", ""), "zip media type")
    if status == 200:
        zf = zipfile.ZipFile(BytesIO(body))
        names = zf.namelist()
        check(names == ["hebdo_S39.xlsx", "hebdo_S39.pptx", "S39-metrics.json"],
              f"zip contents {names}")

    status, headers, body = post_multipart(
        "/api/hebdo/generate", {"week": "S39", "output": "pdf"},
        {"brise": ("b.csv", BRISE), "gnoc": ("g.csv", GNOC)})
    check(status == 400, f"bad output -> 400 ({status})")

    print(f"{CHECKS} checks run")
    if FAILURES:
        print(f"FAILED ({len(FAILURES)}):")
        for failure in FAILURES:
            print("  -", failure)
        raise SystemExit(1)
    print("OK: API checks passed")


if __name__ == "__main__":
    main()
