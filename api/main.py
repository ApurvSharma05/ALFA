"""
api/main.py

ALFA Headless REST API — FastAPI serving layer.

Endpoints
---------
GET  /health            Liveness and readiness probe.
POST /v1/extract        Upload a PDF + company name → extracted CompanyData JSON.
POST /v1/fill           Upload a template .xlsx + CompanyData JSON → populated .xlsx stream.

Run locally (requires `pip install fastapi uvicorn`):
    uvicorn api.main:app --reload --port 8000

Run with Docker Compose (add the service to docker-compose.yml):
    command: uvicorn api.main:app --host 0.0.0.0 --port 8000

Environment variables mirror the Streamlit app (.env is loaded automatically).
"""

import io
import json
import os
import shutil
import tempfile
from typing import Optional

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import StreamingResponse

from api.schemas import (
    ErrorResponse,
    ExtractionResponse,
    FillRequest,
    FillResponse,
    HealthResponse,
)
from config.settings import load_settings
from core.excel_engine import expand_dynamic_tables, fill_factsheet
from core.extractor import GeminiExtractor
from core.pdf_processor import extract_relevant_pages, is_scanned_pdf, validate_pdf
from config.template_mapping import FactSheetMapping


# ---------------------------------------------------------------------------
# Application factory
# ---------------------------------------------------------------------------

_settings = load_settings()

app = FastAPI(
    title="ALFA API",
    description=(
        "Headless REST interface for ALFA — Automated Lead & Financial Analysis. "
        "Provides PDF extraction and Excel template population endpoints decoupled "
        "from the Streamlit UI."
    ),
    version="1.0.0",
    docs_url="/docs",
    redoc_url="/redoc",
)


# ---------------------------------------------------------------------------
# Helper: build a GeminiExtractor from settings or a per-request API key
# ---------------------------------------------------------------------------

def _get_extractor(api_key: Optional[str] = None) -> GeminiExtractor:
    """Return a GeminiExtractor using the provided key or the environment default."""
    key = api_key or _settings.gemini_api_key
    if not key:
        raise HTTPException(
            status_code=401,
            detail=(
                "Gemini API key is required. Pass it as the 'x-gemini-api-key' "
                "header or set GEMINI_API_KEY in the environment."
            ),
        )
    return GeminiExtractor(
        api_key=key,
        model=_settings.gemini_model,
        prompt_path=_settings.skill_prompt_path,
    )


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.get(
    "/health",
    response_model=HealthResponse,
    summary="Liveness / readiness probe",
    tags=["System"],
)
def health() -> HealthResponse:
    """Returns HTTP 200 when the service is up and the settings are loaded."""
    return HealthResponse(status="healthy", model=_settings.gemini_model)


@app.post(
    "/v1/extract",
    response_model=ExtractionResponse,
    summary="Extract structured data from an Annual Report PDF",
    tags=["Extraction"],
    responses={
        400: {"model": ErrorResponse, "description": "Invalid or unreadable PDF."},
        401: {"model": ErrorResponse, "description": "Missing or invalid Gemini API key."},
        502: {"model": ErrorResponse, "description": "Gemini API returned an error after retries."},
    },
)
async def extract(
    file: UploadFile = File(..., description="Annual Report PDF file."),
    company_name: str = Form(..., description="Company name used in the extraction prompt."),
    api_key: Optional[str] = Form(None, description="Gemini API key (overrides GEMINI_API_KEY env var)."),
) -> ExtractionResponse:
    """
    Accept a PDF upload and a company name, run the dual-mode extraction
    pipeline (text → Gemini text request; scanned → Gemini File/Vision API),
    validate the result against the CompanyData Pydantic schema, and return
    the structured JSON.

    The returned ``data`` dict can be passed directly to POST /v1/fill.
    """
    extractor = _get_extractor(api_key)

    # Write the uploaded bytes to a named temp file so pdfplumber can open it.
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        tmp_path = tmp.name
        content = await file.read()
        tmp.write(content)

    try:
        # Validate the PDF before spending API quota.
        pdf_info = validate_pdf(tmp_path)
        if not pdf_info.is_valid:
            raise HTTPException(
                status_code=400,
                detail=f"Invalid PDF: {pdf_info.error or 'file is empty or encrypted'}",
            )

        # Extract and route.
        text = extract_relevant_pages(tmp_path, max_pages=_settings.max_relevant_pages)
        if is_scanned_pdf(text, threshold=_settings.vision_fallback_threshold):
            result = extractor.extract_from_pdf(tmp_path, company_name)
            mode = "vision"
        else:
            result = extractor.extract_from_text(text, company_name)
            mode = "text"

        if not result.success:
            raise HTTPException(status_code=502, detail=result.error)

        return ExtractionResponse(
            entity=result.data.get("entity", company_name),
            mode=mode,
            data=result.data,
            warnings=[],
        )

    finally:
        # Always remove the temp file.
        if os.path.exists(tmp_path):
            os.remove(tmp_path)


@app.post(
    "/v1/fill",
    summary="Populate a BD Fact Sheet Excel template with extracted data",
    tags=["Excel Generation"],
    responses={
        200: {
            "content": {
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": {}
            },
            "description": "Populated .xlsx file streamed as binary response.",
        },
        400: {"model": ErrorResponse, "description": "Invalid template or malformed data."},
    },
)
async def fill(
    template: UploadFile = File(..., description="BD Fact Sheet Excel template (.xlsx)."),
    payload: str = Form(..., description="JSON string matching the FillRequest schema."),
) -> StreamingResponse:
    """
    Accept a BD Fact Sheet Excel template and the CompanyData JSON produced by
    /v1/extract (or the Streamlit review editor).

    Processing steps
    ----------------
    1. Parse and validate the FillRequest payload.
    2. Pre-expand the shareholding, RPT, and Litigation tables as needed via
       expand_dynamic_tables().
    3. Populate all cells using fill_factsheet(), attaching source comments.
    4. Stream the resulting .xlsx back to the caller.

    Returns the populated workbook as a binary file download with content-type
    ``application/vnd.openxmlformats-officedocument.spreadsheetml.sheet``.
    """
    # Parse the JSON payload.
    try:
        body = FillRequest(**json.loads(payload))
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Invalid payload: {exc}")

    data = body.data

    with tempfile.TemporaryDirectory() as tmp_dir:
        # Save uploaded template.
        template_path = os.path.join(tmp_dir, "template.xlsx")
        template_bytes = await template.read()
        with open(template_path, "wb") as f:
            f.write(template_bytes)

        # Determine worksheet name.
        from openpyxl import load_workbook as _lw
        wb_check = _lw(template_path)
        sheet_name = body.sheet_name or wb_check.sheetnames[0]
        wb_check.close()

        mapping = FactSheetMapping()
        template_to_use = template_path

        # ── Shareholding expansion ──────────────────────────────────────────
        from core.safe_sheet import insert_rows_safe
        required_sh_rows = len(data.get("fields", {}).get("shareholding", {}).get("rows", []))
        if required_sh_rows > mapping.shareholding_default_capacity:
            expanded_sh = os.path.join(tmp_dir, "expanded_sh.xlsx")
            delta = required_sh_rows - mapping.shareholding_default_capacity
            insert_rows_safe(
                src_path=template_path,
                out_path=expanded_sh,
                sheet_name=sheet_name,
                threshold=mapping.shareholding_insert_threshold,
                delta=delta,
            )
            template_to_use = expanded_sh

        # ── RPT and Litigation expansion ────────────────────────────────────
        rpt_lit_path = os.path.join(tmp_dir, "expanded_rpt_lit.xlsx")
        template_to_use, _ = expand_dynamic_tables(
            data=data,
            template_path=template_to_use,
            out_path=rpt_lit_path,
            sheet_name=sheet_name,
            mapping=mapping,
        )

        # ── Populate ────────────────────────────────────────────────────────
        out_path = os.path.join(tmp_dir, "populated.xlsx")
        try:
            fill_result = fill_factsheet(
                data=data,
                template_path=template_to_use,
                out_path=out_path,
                sheet_name=sheet_name,
                mapping=mapping,
            )
        except Exception as exc:
            raise HTTPException(status_code=400, detail=f"Excel generation error: {exc}")

        # Read output bytes into memory before tmp_dir is deleted.
        with open(out_path, "rb") as f:
            workbook_bytes = f.read()

    company_name = data.get("entity", "Company")
    filename = f"{company_name}_Populated_BD.xlsx"

    return StreamingResponse(
        io.BytesIO(workbook_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"',
            "X-ALFA-RPT-Rows-Written": str(fill_result.rpt_rows_written),
            "X-ALFA-Lit-Rows-Written": str(fill_result.lit_rows_written),
            "X-ALFA-Warnings": "; ".join(fill_result.warnings) if fill_result.warnings else "none",
        },
    )
