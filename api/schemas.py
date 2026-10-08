"""
api/schemas.py

Pydantic request and response schemas for the ALFA FastAPI serving layer.

These are kept separate from core/schemas.py to:
  - Avoid importing Streamlit dependencies through the core package.
  - Define API-specific response envelopes (status, error, metadata).
  - Allow the API contract to evolve independently of the internal data model.
"""

from __future__ import annotations

from typing import Any, Dict, List, Optional
from pydantic import BaseModel, Field


# ---------------------------------------------------------------------------
# Shared envelope
# ---------------------------------------------------------------------------

class ErrorResponse(BaseModel):
    """Standard error envelope returned on 4xx / 5xx responses."""
    detail: str = Field(description="Human-readable error description.")
    code: Optional[str] = Field(None, description="Machine-readable error code.")


class HealthResponse(BaseModel):
    """Response from the /health endpoint."""
    status: str = Field("healthy", description="Always 'healthy' when the service is up.")
    model: str = Field(description="Configured Gemini model identifier.")
    version: str = Field("1.0.0", description="API version string.")


# ---------------------------------------------------------------------------
# /v1/extract
# ---------------------------------------------------------------------------

class ExtractionResponse(BaseModel):
    """
    Response from POST /v1/extract.

    ``data`` contains the validated CompanyData JSON dict — the same shape
    produced by GeminiExtractor and stored in Streamlit session state.
    ``mode`` is either 'text' or 'vision' indicating the extraction path taken.
    ``warnings`` carries non-fatal issues (e.g., ambiguous figures flagged by
    the model).
    """
    entity: str = Field(description="Name of the company extracted from the PDF.")
    mode: str = Field(description="Extraction mode: 'text' or 'vision'.")
    data: Dict[str, Any] = Field(description="Validated CompanyData JSON dict.")
    warnings: List[str] = Field(default_factory=list)


# ---------------------------------------------------------------------------
# /v1/fill
# ---------------------------------------------------------------------------

class FillRequest(BaseModel):
    """
    Body for POST /v1/fill.

    ``data`` must be a dict matching the CompanyData schema produced by
    /v1/extract (or by Streamlit's review editor).  ``sheet_name`` identifies
    which worksheet inside the uploaded template to populate.
    """
    data: Dict[str, Any] = Field(
        description="CompanyData JSON dict from /v1/extract or the review editor."
    )
    sheet_name: Optional[str] = Field(
        None,
        description="Worksheet to populate.  Defaults to the first sheet in the template.",
    )


class FillResponse(BaseModel):
    """Response from POST /v1/fill."""
    company_name: str
    shareholding_rows_written: int
    rpt_rows_written: int
    lit_rows_written: int
    rpt_rows_dropped: int
    lit_rows_dropped: int
    warnings: List[str] = Field(default_factory=list)
    # The populated .xlsx is returned as a StreamingResponse binary, not in this body.
