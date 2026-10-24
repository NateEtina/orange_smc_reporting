"""FastAPI app: web CSV -> Excel converter.

Run with:  python main.py     (or: uvicorn main:app --reload)
"""

from __future__ import annotations

import asyncio
import io
import json
import re
import zipfile
from dataclasses import asdict
from pathlib import Path
from typing import Any, Dict, List
from urllib.parse import quote

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse, Response

from converter import (
    ConversionError,
    ENCODINGS,
    Options,
    convert_many,
    preview_file,
)
from hebdo import HebdoError, analyze as hebdo_analyze, generate as hebdo_build

BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"

MAX_FILE_BYTES = 300 * 1024 * 1024      # 300 MB per file
MAX_FILES = 50
MAX_REPORT_TYPES_IN_HEADER = 200        # keep the X-Report header reasonably sized

XLSX_MEDIA = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
PPTX_MEDIA = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
ZIP_MEDIA = "application/zip"
HEBDO_OUTPUTS = ("xlsx", "pptx", "both")

app = FastAPI(
    title="CSV -> Excel Converter",
    description="Convert CSV files to richly formatted Excel workbooks, right in your browser.",
    version="1.0.0",
)


@app.get("/", include_in_schema=False)
def index() -> FileResponse:
    return FileResponse(STATIC_DIR / "index.html")


@app.get("/api/health", include_in_schema=False)
@app.get("/health", include_in_schema=False)
def health() -> Dict[str, Any]:
    return {"status": "ok"}


@app.get("/api/options", include_in_schema=False)
def default_options() -> Dict[str, Any]:
    return {"options": asdict(Options()), "encodings": ENCODINGS}


def _parse_options(raw: str) -> Options:
    try:
        data = json.loads(raw) if raw else {}
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=400, detail=f"Invalid options JSON: {exc}") from exc
    if not isinstance(data, dict):
        raise HTTPException(status_code=400, detail="Options must be a JSON object.")
    return Options.from_dict(data)


def _safe_filename(name: str, fallback: str) -> str:
    stem = re.sub(r'[\\/:*?"<>|]+', "_", (name or fallback)).strip().strip(".")
    return (stem or fallback)[:100]


def _report_header(reports: Any) -> Dict[str, str]:
    payload = json.dumps(reports, ensure_ascii=False, default=str)
    return {"X-Report": quote(payload[:60_000])}


def _unique(name: str, used: set) -> str:
    candidate, k = name, 2
    while candidate.lower() in used:
        stem, dot, ext = name.rpartition(".")
        if not dot:
            stem, ext = name, ""
        candidate = f"{stem} ({k}).{ext}" if dot else f"{name} ({k})"
        k += 1
    used.add(candidate.lower())
    return candidate


@app.post("/api/preview")
async def preview(file: UploadFile = File(...), options: str = Form("{}")) -> JSONResponse:
    opts = _parse_options(options)
    raw = await file.read()
    if len(raw) > MAX_FILE_BYTES:
        raise HTTPException(status_code=413, detail="File is larger than 300 MB.")
    try:
        result = await asyncio.to_thread(preview_file, file.filename or "data.csv", raw, opts)
    except ConversionError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:  # pragma: no cover - defensive
        raise HTTPException(status_code=400, detail=f"Could not read the file: {exc}") from exc
    return JSONResponse(result)


@app.post("/api/convert")
async def convert(files: List[UploadFile] = File(...), options: str = Form("{}")) -> Response:
    opts = _parse_options(options)
    if not files:
        raise HTTPException(status_code=400, detail="No files were uploaded.")
    if len(files) > MAX_FILES:
        raise HTTPException(status_code=400, detail=f"At most {MAX_FILES} files per conversion.")

    items: List[tuple] = []
    for upload in files:
        name = upload.filename or "data.csv"
        raw = await upload.read()
        if not raw:
            raise HTTPException(status_code=400, detail=f"'{name}' is empty.")
        if len(raw) > MAX_FILE_BYTES:
            raise HTTPException(status_code=413, detail=f"'{name}' is larger than 300 MB.")
        items.append((name, raw))

    try:
        outputs, reports = await asyncio.to_thread(convert_many, items, opts)
    except ConversionError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:  # pragma: no cover - defensive
        raise HTTPException(status_code=500, detail=f"Conversion failed: {exc}") from exc

    report_data: Any = reports
    if len(reports) == 1 and len(reports[0].get("types", [])) > MAX_REPORT_TYPES_IN_HEADER:
        trimmed = dict(reports[0])
        trimmed["types"] = trimmed["types"][:MAX_REPORT_TYPES_IN_HEADER]
        trimmed["types_truncated"] = True
        report_data = [trimmed]

    if len(outputs) == 1:
        filename, data = outputs[0]
        headers = {
            "Content-Disposition": f'attachment; filename="{_safe_filename(filename, "converted.xlsx")}"',
            **_report_header(report_data),
        }
        return Response(data, media_type=XLSX_MEDIA, headers=headers)

    buffer = io.BytesIO()
    used: set = set()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as bundle:
        for name, data in outputs:
            bundle.writestr(_unique(_safe_filename(name, "converted.xlsx"), used), data)
        bundle.writestr(
            "conversion-report.json",
            json.dumps(reports, indent=2, ensure_ascii=False, default=str),
        )
    headers = {
        "Content-Disposition": 'attachment; filename="converted-csv.zip"',
        **_report_header(report_data),
    }
    return Response(buffer.getvalue(), media_type="application/zip", headers=headers)


@app.get("/hebdo", include_in_schema=False)
def hebdo_page() -> FileResponse:
    return FileResponse(STATIC_DIR / "hebdo.html")


# --------------------------------------------------------------------------- #
# Feature 1 - weekly Excel / PowerPoint report (brise_smc + GNOC -> hebdo_SXX)
# --------------------------------------------------------------------------- #

async def _hebdo_source(upload: UploadFile, label: str) -> tuple:
    raw = await upload.read()
    name = upload.filename or f"{label}.csv"
    if not raw:
        raise HTTPException(status_code=400, detail=f"'{name}' is empty.")
    if len(raw) > MAX_FILE_BYTES:
        raise HTTPException(status_code=413, detail=f"'{name}' is larger than 300 MB.")
    return name, raw


def _parse_columns(raw: str) -> Dict[str, str]:
    try:
        data = json.loads(raw) if raw else {}
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=400, detail=f"Invalid columns JSON: {exc}") from exc
    if not isinstance(data, dict):
        raise HTTPException(status_code=400, detail="columns must be a JSON object.")
    return {str(k): str(v) for k, v in data.items() if v}


async def _hebdo_inputs(brise: UploadFile, gnoc: UploadFile, week: str, date: str,
                        columns: str):
    pair_b = await _hebdo_source(brise, "brise_smc")
    pair_g = await _hebdo_source(gnoc, "GNOC")
    return pair_b, pair_g, (week or None), (date or None), _parse_columns(columns)


async def _hebdo_run(func, args, kwargs, generic_status: int = 500):
    try:
        return await asyncio.to_thread(func, *args, **kwargs)
    except HebdoError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:  # noqa: BLE001 - defensive
        raise HTTPException(status_code=generic_status,
                            detail=f"Report generation failed: {exc}") from exc


@app.post("/api/hebdo/preview")
async def hebdo_preview(brise: UploadFile = File(...), gnoc: UploadFile = File(...),
                        week: str = Form(""), date: str = Form(""),
                        columns: str = Form("{}")) -> JSONResponse:
    b, g, wk, day, cols = await _hebdo_inputs(brise, gnoc, week, date, columns)
    metrics = await _hebdo_run(hebdo_analyze, (b, g),
                               {"week": wk, "date": day, "columns": cols},
                               generic_status=400)
    return JSONResponse(metrics)


@app.post("/api/hebdo/generate")
async def hebdo_generate(brise: UploadFile = File(...), gnoc: UploadFile = File(...),
                         week: str = Form(""), date: str = Form(""),
                         columns: str = Form("{}"),
                         output: str = Form("both")) -> Response:
    if output not in HEBDO_OUTPUTS:
        raise HTTPException(status_code=400,
                            detail=f"output must be one of: {', '.join(HEBDO_OUTPUTS)}.")
    b, g, wk, day, cols = await _hebdo_inputs(brise, gnoc, week, date, columns)
    outputs = ("xlsx", "pptx") if output == "both" else (output,)
    files, metrics = await _hebdo_run(hebdo_build, (b, g),
                                      {"week": wk, "date": day, "columns": cols,
                                       "outputs": outputs})
    label = metrics["week_label"]

    if len(files) == 1:
        name, data = files[0]
        media = XLSX_MEDIA if name.endswith(".xlsx") else PPTX_MEDIA
        return Response(data, media_type=media, headers={
            "Content-Disposition": f'attachment; filename="{_safe_filename(name, "hebdo.xlsx")}"',
        })

    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as bundle:
        for name, data in files:
            bundle.writestr(_safe_filename(name, "hebdo.bin"), data)
        bundle.writestr(f"{label}-metrics.json",
                        json.dumps(metrics, indent=2, ensure_ascii=False, default=str))
    return Response(buffer.getvalue(), media_type=ZIP_MEDIA, headers={
        "Content-Disposition": f'attachment; filename="hebdo_{label}.zip"',
    })


@app.exception_handler(HTTPException)
async def http_exception_handler(request, exc: HTTPException):
    """Return JSON errors with a consistent shape for the UI."""
    return JSONResponse(status_code=exc.status_code, content={"detail": exc.detail})


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="127.0.0.1", port=8000, log_level="info")
