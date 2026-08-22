"""
core/extractor.py

Gemini LLM extraction client with exponential-backoff retries (via tenacity),
structured logging, and clean resource management.

Two extraction modes:
  1. Text-based: sends extracted PDF text as a prompt (fast, cheap)
  2. Vision-based: uploads the PDF file for multimodal analysis (scanned PDFs)
"""

import json
import logging
from dataclasses import dataclass
from typing import Optional

from google import genai
from google.genai import types
from tenacity import (
    retry,
    stop_after_attempt,
    wait_exponential,
    before_sleep_log,
    retry_if_exception_type,
)

from core.schemas import CompanyData
from config.settings import get_system_prompt

logger = logging.getLogger("alfa.extractor")


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
        stop=stop_after_attempt(5),
        wait=wait_exponential(multiplier=1, min=2, max=30),
        before_sleep=before_sleep_log(logger, logging.WARNING),
        reraise=True,
    )
    def _call_text(self, prompt: str) -> str:
        """Call Gemini with a text prompt. Retried on transient failures."""
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
            data = json.loads(raw)
            logger.info("[%s] Text extraction succeeded", company_name)
            return ExtractionResult(company_name=company_name, data=data, mode="text")
        except json.JSONDecodeError as exc:
            logger.error("[%s] JSON parse error: %s", company_name, exc)
            return ExtractionResult(company_name=company_name, error=f"JSON parse error: {exc}", mode="text")
        except Exception as exc:
            logger.error("[%s] Text extraction failed after retries: %s", company_name, exc, exc_info=True)
            return ExtractionResult(company_name=company_name, error=str(exc), mode="text")

    # ---- Vision-based extraction (scanned PDFs) ----

    @retry(
        stop=stop_after_attempt(5),
        wait=wait_exponential(multiplier=1, min=2, max=30),
        before_sleep=before_sleep_log(logger, logging.WARNING),
        reraise=True,
    )
    def _call_vision(self, uploaded_file, prompt: str) -> str:
        """Call Gemini with an uploaded file + prompt. Retried on transient failures."""
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
            data = json.loads(raw)
            logger.info("[%s] Vision extraction succeeded", company_name)
            return ExtractionResult(company_name=company_name, data=data, mode="vision")

        except json.JSONDecodeError as exc:
            logger.error("[%s] JSON parse error (vision): %s", company_name, exc)
            return ExtractionResult(company_name=company_name, error=f"JSON parse error: {exc}", mode="vision")
        except Exception as exc:
            logger.error("[%s] Vision extraction failed after retries: %s", company_name, exc, exc_info=True)
            return ExtractionResult(company_name=company_name, error=str(exc), mode="vision")
        finally:
            if uploaded_file is not None:
                try:
                    self.client.files.delete(name=uploaded_file.name)
                    logger.debug("[%s] Cleaned up uploaded file", company_name)
                except Exception:
                    logger.warning("[%s] Failed to clean up uploaded file", company_name, exc_info=True)
