"""
core/extractor.py

Gemini LLM extraction client with exponential-backoff retries (via tenacity)
for transient errors only (429 / 5xx / network), schema validation of the
returned JSON against CompanyData,
structured logging, and clean resource management.

Two extraction modes:
  1. Text-based: sends extracted PDF text as a prompt (fast, cheap)
  2. Vision-based: uploads the PDF file for multimodal analysis (scanned PDFs)
"""

import json
import logging
from dataclasses import dataclass
from typing import Optional

import httpx
from google import genai
from google.genai import errors as genai_errors
from google.genai import types
from pydantic import ValidationError
from tenacity import (
    retry,
    stop_after_attempt,
    wait_exponential,
    before_sleep_log,
    retry_if_exception,
)

from core.schemas import CompanyData
from config.settings import get_system_prompt

logger = logging.getLogger("alfa.extractor")


# ---------------------------------------------------------------------------
# Retry policy
# ---------------------------------------------------------------------------

def _is_transient(exc: BaseException) -> bool:
    """
    True only for errors worth retrying: rate limits (429), request timeouts
    (408), server errors (5xx) and network-level failures. Auth (401/403),
    bad-request (400) and other 4xx errors fail immediately.
    """
    if isinstance(exc, genai_errors.APIError):
        code = getattr(exc, "code", None)
        return code in (408, 429) or (isinstance(code, int) and code >= 500)
    return isinstance(exc, (httpx.TransportError, ConnectionError, TimeoutError))


def _validate_payload(raw: str) -> dict:
    """
    Parse the model's JSON and validate it against CompanyData.

    Returns the validated, coerced dict (string numbers converted to numbers,
    missing optional keys filled with None). Raises json.JSONDecodeError for
    malformed JSON or pydantic.ValidationError for schema violations.
    """
    return CompanyData.model_validate(json.loads(raw)).model_dump()


def _format_validation_error(exc: ValidationError) -> str:
    """Compact, human-readable summary of a pydantic ValidationError."""
    errs = exc.errors()
    lines = [f"{'.'.join(str(p) for p in e['loc'])}: {e['msg']}" for e in errs[:5]]
    more = f" (+{len(errs) - 5} more)" if len(errs) > 5 else ""
    return f"Schema validation failed ({len(errs)} error(s)): " + "; ".join(lines) + more


# ---------------------------------------------------------------------------
# Extraction result wrapper
# ---------------------------------------------------------------------------

@dataclass
class ExtractionResult:
    """Wraps a successful or failed extraction attempt."""
    company_name: str
    data: Optional[dict] = None
    error: Optional[str] = None
    mode: str = "text"  # "text" or "vision"

    @property
    def success(self) -> bool:
        return self.data is not None and self.error is None


# ---------------------------------------------------------------------------
# Gemini Extractor
# ---------------------------------------------------------------------------

class GeminiExtractor:
    """
    Stateful extractor wrapping Google GenAI client.

    Usage::

        extractor = GeminiExtractor(api_key="...", model="gemini-flash-latest")
        result = extractor.extract_from_text(text, "Acme Corp")
        if result.success:
            print(result.data)
    """

    def __init__(self, api_key: str, model: str = "gemini-flash-latest", prompt_path: str = "SKILL.md"):
        self.client = genai.Client(api_key=api_key)
        self.model = model
        self.system_prompt = get_system_prompt(prompt_path)
        self._config = types.GenerateContentConfig(
            system_instruction=self.system_prompt,
            temperature=0.0,
            response_mime_type="application/json",
            response_schema=CompanyData,
        )

    # ---- Text-based extraction ----

    @retry(
        retry=retry_if_exception(_is_transient),
        stop=stop_after_attempt(5),
        wait=wait_exponential(multiplier=1, min=2, max=30),
        before_sleep=before_sleep_log(logger, logging.WARNING),
        reraise=True,
    )
    def _call_text(self, prompt: str) -> str:
        """Call Gemini with a text prompt. Retried only on transient failures (429/5xx/network)."""
        response = self.client.models.generate_content(
            model=self.model,
            contents=prompt,
            config=self._config,
        )
        return response.text

    def extract_from_text(self, text: str, company_name: str) -> ExtractionResult:
        """
        Extract BD data from pre-extracted PDF text.

        Args:
            text: full or filtered text from the annual report
            company_name: company name for the prompt

        Returns:
            ExtractionResult with parsed data or error details
        """
        prompt = (
            f"Please extract the Business Description data for {company_name} "
            f"from the following Annual Report text:\n\n{text}"
        )
        try:
            raw = self._call_text(prompt)
            data = _validate_payload(raw)
            logger.info("[%s] Text extraction succeeded", company_name)
            return ExtractionResult(company_name=company_name, data=data, mode="text")
        except ValidationError as exc:
            msg = _format_validation_error(exc)
            logger.error("[%s] %s", company_name, msg)
            return ExtractionResult(company_name=company_name, error=msg, mode="text")
        except json.JSONDecodeError as exc:
            logger.error("[%s] JSON parse error: %s", company_name, exc)
            return ExtractionResult(company_name=company_name, error=f"JSON parse error: {exc}", mode="text")
        except Exception as exc:
            logger.error("[%s] Text extraction failed: %s", company_name, exc, exc_info=True)
            return ExtractionResult(company_name=company_name, error=str(exc), mode="text")

    # ---- Vision-based extraction (scanned PDFs) ----

    @retry(
        retry=retry_if_exception(_is_transient),
        stop=stop_after_attempt(5),
        wait=wait_exponential(multiplier=1, min=2, max=30),
        before_sleep=before_sleep_log(logger, logging.WARNING),
        reraise=True,
    )
    def _call_vision(self, uploaded_file, prompt: str) -> str:
        """Call Gemini with an uploaded file + prompt. Retried only on transient failures (429/5xx/network)."""
        response = self.client.models.generate_content(
            model=self.model,
            contents=[uploaded_file, prompt],
            config=self._config,
        )
        return response.text

    def extract_from_pdf(self, pdf_path: str, company_name: str) -> ExtractionResult:
        """
        Extract BD data by uploading the PDF to Gemini File API (vision mode).

        The uploaded file is always cleaned up, even on failure.

        Args:
            pdf_path: path to the PDF file
            company_name: company name for the prompt

        Returns:
            ExtractionResult with parsed data or error details
        """
        uploaded_file = None
        try:
            uploaded_file = self.client.files.upload(file=pdf_path)
            prompt = (
                f"Please extract the Business Description data for {company_name} "
                f"from the attached Annual Report PDF."
            )
            raw = self._call_vision(uploaded_file, prompt)
            data = _validate_payload(raw)
            logger.info("[%s] Vision extraction succeeded", company_name)
            return ExtractionResult(company_name=company_name, data=data, mode="vision")

        except ValidationError as exc:
            msg = _format_validation_error(exc)
            logger.error("[%s] %s (vision)", company_name, msg)
            return ExtractionResult(company_name=company_name, error=msg, mode="vision")
        except json.JSONDecodeError as exc:
            logger.error("[%s] JSON parse error (vision): %s", company_name, exc)
            return ExtractionResult(company_name=company_name, error=f"JSON parse error: {exc}", mode="vision")
        except Exception as exc:
            logger.error("[%s] Vision extraction failed: %s", company_name, exc, exc_info=True)
            return ExtractionResult(company_name=company_name, error=str(exc), mode="vision")
        finally:
            if uploaded_file is not None:
                try:
                    self.client.files.delete(name=uploaded_file.name)
                    logger.debug("[%s] Cleaned up uploaded file", company_name)
                except Exception:
                    logger.warning("[%s] Failed to clean up uploaded file", company_name, exc_info=True)
