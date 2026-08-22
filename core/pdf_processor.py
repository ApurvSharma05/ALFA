"""
core/pdf_processor.py

PDF validation, text extraction, and intelligent page filtering.

Key capabilities:
  - validate_pdf: checks for encrypted, empty, or corrupt files
  - extract_text: full pdfplumber text extraction
  - extract_relevant_pages: keyword-scored page filtering to reduce token cost
  - is_scanned_pdf: heuristic check for image-only PDFs
"""

import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional

import pdfplumber

logger = logging.getLogger("alfa.pdf")

# Keywords scored by relevance to TP BD extraction.
# Higher weight = more likely to contain critical data.
SECTION_KEYWORDS = {
    # Financial statements
    "statement of profit and loss": 10,
    "profit and loss": 10,
    "revenue from operations": 10,
    "total expenses": 8,
    "balance sheet": 6,
    "other income": 5,

    # Related party
    "related party": 10,
    "related parties": 10,
    "associated enterprise": 8,
    "transactions with related": 8,

    # Shareholding
    "shareholding pattern": 9,
    "promoter": 7,
    "beneficial ownership": 6,

    # Contingent liabilities / litigation
    "contingent liabilities": 10,
    "contingent liability": 10,
    "litigation": 8,
    "claims against the company": 7,
    "disputed": 6,

    # General corporate
    "registered office": 6,
    "statutory auditor": 5,
    "auditor": 4,
    "subsidiary": 5,
    "country": 3,
}


@dataclass
class PDFInfo:
    """Metadata about a validated PDF."""
    path: str
    page_count: int
    is_encrypted: bool = False
    is_empty: bool = False
    error: Optional[str] = None

    @property
    def is_valid(self) -> bool:
        return not self.is_encrypted and not self.is_empty and self.error is None


def validate_pdf(path: str) -> PDFInfo:
    """
    Open a PDF and check for common problems before spending API quota.

    Returns a PDFInfo with diagnostics.  Callers should check ``info.is_valid``
    before proceeding.
    """
    resolved = Path(path)
    if not resolved.exists():
        return PDFInfo(path=path, page_count=0, error=f"File not found: {path}")

    try:
        with pdfplumber.open(path) as pdf:
            page_count = len(pdf.pages)
            if page_count == 0:
                return PDFInfo(path=path, page_count=0, is_empty=True)
            return PDFInfo(path=path, page_count=page_count)
    except Exception as exc:
        # pdfplumber raises various exceptions for encrypted / corrupt files
        err_msg = str(exc).lower()
        is_enc = "encrypt" in err_msg or "password" in err_msg
        return PDFInfo(
            path=path,
            page_count=0,
            is_encrypted=is_enc,
            error=str(exc),
        )


def extract_text(path: str) -> str:
    """Extract text from every page of a PDF using pdfplumber."""
    pages: List[str] = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                pages.append(text)
    full_text = "\n".join(pages)
    logger.info("Extracted %d chars from %d pages of %s", len(full_text), len(pages), path)
    return full_text


def _score_page(text: str) -> int:
    """Score a single page's relevance based on keyword hits."""
    lower = text.lower()
    score = 0
    for keyword, weight in SECTION_KEYWORDS.items():
        if keyword in lower:
            score += weight
    return score


def extract_relevant_pages(
    path: str,
    max_pages: int = 40,
    min_chars_fallback: int = 5000,
) -> str:
    """
    Extract text from the most relevant pages of a PDF, ranked by keyword
    scoring.  Falls back to full text if the filtered result is too small.

    Args:
        path: PDF file path
        max_pages: maximum number of pages to include
        min_chars_fallback: if filtered text is below this threshold, return
            full text instead (guards against over-aggressive filtering)

    Returns:
        Extracted text string
    """
    scored_pages: List[tuple] = []  # (page_index, score, text)

    with pdfplumber.open(path) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            score = _score_page(text)
            scored_pages.append((i, score, text))

    # If the PDF is small enough, just return everything
    if len(scored_pages) <= max_pages:
        full = "\n".join(t for _, _, t in scored_pages if t)
        logger.info(
            "PDF has %d pages (<= %d max), using full text (%d chars)",
            len(scored_pages), max_pages, len(full),
        )
        return full

    # Sort by score descending, take top N, then re-sort by page order
    scored_pages.sort(key=lambda x: x[1], reverse=True)
    selected = scored_pages[:max_pages]
    selected.sort(key=lambda x: x[0])  # restore page order

    filtered_text = "\n".join(t for _, _, t in selected if t)

    # Fallback: if filtering was too aggressive, use full text
    if len(filtered_text) < min_chars_fallback:
        full = "\n".join(t for _, _, t in scored_pages if t)
        logger.warning(
            "Filtered text too short (%d chars < %d threshold), falling back to full text",
            len(filtered_text), min_chars_fallback,
        )
        return full

    logger.info(
        "Filtered to %d/%d pages (%d chars) based on keyword scoring",
        len(selected), len(scored_pages), len(filtered_text),
    )
    return filtered_text


def is_scanned_pdf(text: str, threshold: int = 2000) -> bool:
    """
    Heuristic: if extracted text is below threshold characters, the PDF is
    likely scanned/image-based and needs the Vision API.
    """
    return len(text.strip()) < threshold
