"""
config/settings.py

Centralized application settings loaded from environment variables / .env file.
All configurable values live here — nothing is hardcoded in the core modules.
"""

import os
import logging
from dataclasses import dataclass, field
from functools import lru_cache
from pathlib import Path

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# ---------------------------------------------------------------------------
# Logging setup
# ---------------------------------------------------------------------------

LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO").upper()
logging.basicConfig(
    level=getattr(logging, LOG_LEVEL, logging.INFO),
    format="%(asctime)s | %(name)s | %(levelname)s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger("alfa")


# ---------------------------------------------------------------------------
# Application Settings
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class Settings:
    """Immutable application-wide settings."""

    # --- Gemini API ---
    gemini_api_key: str = ""
    gemini_model: str = "gemini-flash-latest"

    # --- Processing ---
    max_workers: int = 3  # concurrent PDF extractions
    vision_fallback_threshold: int = 2000  # chars below which we use Vision API

    # --- Excel template defaults ---
    max_rpt_rows: int = 15  # generous default RPT capacity
    max_lit_rows: int = 10  # generous default Litigation capacity

    # --- PDF filtering ---
    max_relevant_pages: int = 40  # max pages to send to LLM after filtering

    # --- Paths ---
    skill_prompt_path: str = "SKILL.md"


def load_settings(**overrides) -> Settings:
    """
    Build a Settings instance from environment variables, with optional
    keyword overrides (useful for testing or UI-driven config).
    """
    env_values = {
        "gemini_api_key": os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY", ""),
        "gemini_model": os.getenv("GEMINI_MODEL", "gemini-flash-latest"),
        "max_workers": int(os.getenv("MAX_WORKERS", "3")),
        "vision_fallback_threshold": int(os.getenv("VISION_FALLBACK_THRESHOLD", "2000")),
        "max_rpt_rows": int(os.getenv("MAX_RPT_ROWS", "15")),
        "max_lit_rows": int(os.getenv("MAX_LIT_ROWS", "10")),
        "max_relevant_pages": int(os.getenv("MAX_RELEVANT_PAGES", "40")),
        "skill_prompt_path": os.getenv("SKILL_PROMPT_PATH", "SKILL.md"),
    }
    env_values.update(overrides)
    return Settings(**env_values)


@lru_cache(maxsize=1)
def get_system_prompt(path: str = "SKILL.md") -> str:
    """Read and cache the SKILL.md system prompt. Called once per process."""
    resolved = Path(path)
    if not resolved.exists():
        logger.warning("System prompt file not found at %s", resolved)
        return ""
    text = resolved.read_text(encoding="utf-8")
    logger.info("Loaded system prompt from %s (%d chars)", resolved, len(text))
    return text
