#!/usr/bin/env python
"""
scripts/run_eval.py

ALFA Automated Evaluation Harness.

Compares ALFA's LLM extraction output against verified ground-truth values in
tests/data/golden_dataset.json, then prints a field-level accuracy report.

Usage
-----
    # Standard run (uses GEMINI_API_KEY from .env)
    python scripts/run_eval.py

    # Specify a different golden dataset
    python scripts/run_eval.py --dataset path/to/my_dataset.json

    # Output full JSON results to a file
    python scripts/run_eval.py --out eval_results.json

    # Skip extraction (re-use cached results from a previous run)
    python scripts/run_eval.py --from-cache eval_results.json

Evaluation Metrics
------------------
For numeric fields:      Relative tolerance of ±2 % (configurable with --tol).
For string fields:       Exact match (case-insensitive strip).
For count fields:        Exact integer match.
For list/label fields:   Order-insensitive set containment.

Exit Codes
----------
    0  All tested fields pass within tolerance.
    1  At least one field fails or the dataset has unresolved TODOs.
    2  Import / configuration error.
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import sys
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

# ---------------------------------------------------------------------------
# Path setup so we can import core/ from the project root.
# ---------------------------------------------------------------------------
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from dotenv import load_dotenv
load_dotenv(ROOT / ".env")

from core.extractor import GeminiExtractor
from core.pdf_processor import extract_relevant_pages, is_scanned_pdf, validate_pdf
from config.settings import load_settings

logging.basicConfig(
    level=os.getenv("LOG_LEVEL", "INFO"),
    format="%(asctime)s  %(levelname)-8s  %(name)s — %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger("alfa.eval")

GOLDEN_DATASET_PATH = ROOT / "tests" / "data" / "golden_dataset.json"
DEFAULT_TOL = 0.02  # ± 2 % relative tolerance for numeric fields


# ---------------------------------------------------------------------------
# Numeric comparison helpers
# ---------------------------------------------------------------------------

def _numeric_match(expected: float, actual: Any, tol: float = DEFAULT_TOL) -> Tuple[bool, str]:
    """
    Return (passed, reason) for a numeric field comparison.
    Tolerates small floating-point differences to account for rounding in LLM output.
    """
    try:
        actual_f = float(actual)
    except (TypeError, ValueError):
        return False, f"actual value '{actual}' is not numeric"

    if expected == 0:
        passed = actual_f == 0
        return passed, f"expected 0, got {actual_f}"

    relative_err = abs((actual_f - expected) / expected)
    passed = relative_err <= tol
    return passed, f"expected {expected}, got {actual_f}, rel_err={relative_err:.2%}"


def _string_match(expected: str, actual: Any) -> Tuple[bool, str]:
    actual_s = str(actual or "").strip().lower()
    expected_s = str(expected or "").strip().lower()
    return actual_s == expected_s, f"expected '{expected}', got '{actual}'"


def _set_containment(expected_labels: List[str], actual_labels: List[Any]) -> Tuple[bool, str]:
    expected_set = {s.strip().lower() for s in expected_labels}
    actual_set = {str(s).strip().lower() for s in (actual_labels or [])}
    missing = expected_set - actual_set
    if missing:
        return False, f"missing labels: {missing}"
    return True, "all expected labels present"


# ---------------------------------------------------------------------------
# Field-level evaluation
# ---------------------------------------------------------------------------

def evaluate_case(case: dict, extracted: dict, tol: float) -> List[dict]:
    """
    Compare extracted data against the ground-truth expected_fields in one case.
    Returns a list of result dicts, one per evaluated field.
    """
    expected = case.get("expected_fields", {})
    fields = extracted.get("fields", {})
    results = []

    def _record(field_name: str, passed: bool, reason: str):
        results.append({
            "company": case["company"],
            "field": field_name,
            "passed": passed,
            "reason": reason,
        })

    # ── Numeric fields ────────────────────────────────────────────────────────
    numeric_map = {
        "standalone_turnover_fy25_lakhs":      ("standalone_summary_of_operations", "value", "turnover_fy25"),
        "standalone_total_cost_fy25_lakhs":    ("standalone_summary_of_operations", "value", "total_cost_fy25"),
        "cash_fy25_lakhs":                     ("cash_and_ae_trade_receivables",    "value", "cash_fy25"),
        "ae_trade_receivables_fy25_lakhs":     ("cash_and_ae_trade_receivables",    "value", "ae_receivables_fy25"),
    }
    for key, path in numeric_map.items():
        if expected.get(key) is None:
            continue  # skip unverified fields
        section_key, _, sub_key = path
        actual = (fields.get(section_key) or {}).get("value", {}).get(sub_key)
        passed, reason = _numeric_match(expected[key], actual, tol)
        _record(key, passed, reason)

    # ── Shareholding ─────────────────────────────────────────────────────────
    sh_expected = expected.get("shareholding", {})
    sh_actual = fields.get("shareholding", {})
    if sh_expected.get("expected_rows_count") is not None:
        rows = sh_actual.get("rows", [])
        passed = len(rows) == sh_expected["expected_rows_count"]
        _record(
            "shareholding.rows_count",
            passed,
            f"expected {sh_expected['expected_rows_count']}, got {len(rows)}",
        )
    if sh_expected.get("expected_total_shares_fy25") is not None:
        total = sh_actual.get("total_shares_fy25")
        passed, reason = _numeric_match(sh_expected["expected_total_shares_fy25"], total, tol)
        _record("shareholding.total_shares_fy25", passed, reason)

    # ── RPT ──────────────────────────────────────────────────────────────────
    rpt_expected = expected.get("rpt", {})
    rpt_actual = fields.get("related_party_transactions_lakhs", {})
    if rpt_expected.get("expected_items_count") is not None:
        items = rpt_actual.get("items", [])
        passed = len(items) == rpt_expected["expected_items_count"]
        _record(
            "rpt.items_count",
            passed,
            f"expected {rpt_expected['expected_items_count']}, got {len(items)}",
        )
    if rpt_expected.get("expected_labels"):
        actual_labels = [it.get("label") for it in rpt_actual.get("items", [])]
        passed, reason = _set_containment(rpt_expected["expected_labels"], actual_labels)
        _record("rpt.labels", passed, reason)

    # ── Litigation ────────────────────────────────────────────────────────────
    lit_expected = expected.get("litigation", {})
    lit_actual = fields.get("litigation", {})
    if lit_expected.get("expected_items_count") is not None:
        items = lit_actual.get("items", [])
        passed = len(items) == lit_expected["expected_items_count"]
        _record(
            "litigation.items_count",
            passed,
            f"expected {lit_expected['expected_items_count']}, got {len(items)}",
        )

    return results


# ---------------------------------------------------------------------------
# Extraction runner
# ---------------------------------------------------------------------------

def run_extraction(case: dict, settings) -> Optional[dict]:
    """Run the full extraction pipeline for one golden dataset case."""
    pdf_path_rel = case.get("pdf_path")
    if not pdf_path_rel:
        logger.warning("[%s] No pdf_path, skipping extraction.", case["company"])
        return None

    pdf_path = ROOT / pdf_path_rel
    if not pdf_path.exists():
        logger.warning("[%s] PDF not found at %s — skipping.", case["company"], pdf_path)
        return None

    logger.info("[%s] Extracting from %s …", case["company"], pdf_path.name)
    t0 = time.perf_counter()

    info = validate_pdf(str(pdf_path))
    if not info.is_valid:
        logger.error("[%s] PDF invalid: %s", case["company"], info.error)
        return None

    extractor = GeminiExtractor(
        api_key=settings.gemini_api_key,
        model=settings.gemini_model,
        prompt_path=settings.skill_prompt_path,
    )

    text = extract_relevant_pages(str(pdf_path), max_pages=settings.max_relevant_pages)
    if is_scanned_pdf(text, threshold=settings.vision_fallback_threshold):
        result = extractor.extract_from_pdf(str(pdf_path), case["company"])
    else:
        result = extractor.extract_from_text(text, case["company"])

    elapsed = time.perf_counter() - t0
    logger.info("[%s] Extraction done in %.1fs (success=%s)", case["company"], elapsed, result.success)

    if not result.success:
        logger.error("[%s] Extraction failed: %s", case["company"], result.error)
        return None

    return result.data


# ---------------------------------------------------------------------------
# Report printer
# ---------------------------------------------------------------------------

def _print_report(all_results: List[dict]) -> bool:
    """Print a coloured field-level report. Returns True if all fields pass."""
    GREEN = "\033[92m"
    RED = "\033[91m"
    RESET = "\033[0m"
    BOLD = "\033[1m"

    pass_count = sum(1 for r in all_results if r["passed"])
    fail_count = len(all_results) - pass_count

    print(f"\n{'=' * 70}")
    print(f"  {BOLD}ALFA Evaluation Report{RESET}  —  {pass_count}/{len(all_results)} fields pass")
    print(f"{'=' * 70}")

    current_company = None
    for r in sorted(all_results, key=lambda x: (x["company"], x["field"])):
        if r["company"] != current_company:
            current_company = r["company"]
            print(f"\n  {BOLD}{current_company}{RESET}")

        status_icon = f"{GREEN}✓{RESET}" if r["passed"] else f"{RED}✗{RESET}"
        print(f"    {status_icon}  {r['field']:<48}  {r['reason']}")

    print(f"\n{'=' * 70}")
    result_line = f"  PASS: {pass_count}   FAIL: {fail_count}"
    if fail_count == 0:
        print(f"{GREEN}{BOLD}{result_line}{RESET}")
    else:
        print(f"{RED}{BOLD}{result_line}{RESET}")
    print(f"{'=' * 70}\n")

    return fail_count == 0


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------

def main() -> int:
    parser = argparse.ArgumentParser(description="ALFA evaluation harness.")
    parser.add_argument(
        "--dataset",
        type=str,
        default=str(GOLDEN_DATASET_PATH),
        help="Path to the golden dataset JSON file.",
    )
    parser.add_argument(
        "--out",
        type=str,
        default=None,
        help="If set, write the full per-field result JSON to this file.",
    )
    parser.add_argument(
        "--from-cache",
        type=str,
        default=None,
        metavar="FILE",
        dest="from_cache",
        help="Re-run field evaluation from a previous --out JSON (no LLM calls).",
    )
    parser.add_argument(
        "--tol",
        type=float,
        default=DEFAULT_TOL,
        help=f"Numeric relative tolerance, default {DEFAULT_TOL:.0%}.",
    )
    parser.add_argument(
        "--only",
        type=str,
        nargs="+",
        default=None,
        metavar="COMPANY",
        help="Only evaluate cases whose company name contains any of these strings.",
    )
    args = parser.parse_args()

    # ── Load golden dataset ──────────────────────────────────────────────────
    dataset_path = Path(args.dataset)
    if not dataset_path.exists():
        logger.error("Golden dataset not found at %s", dataset_path)
        return 2
    with open(dataset_path) as f:
        dataset = json.load(f)
    cases = dataset.get("cases", [])

    # ── Filter by --only ─────────────────────────────────────────────────────
    if args.only:
        cases = [
            c for c in cases
            if any(kw.lower() in c["company"].lower() for kw in args.only)
        ]
    if not cases:
        logger.error("No matching cases in dataset.")
        return 2

    # ── Check for unresolved TODOs ───────────────────────────────────────────
    todo_cases = [c for c in cases if "TODO" in (c.get("status") or "")]
    if todo_cases:
        logger.warning(
            "%d case(s) have status=TODO and cannot be evaluated accurately:\n  %s",
            len(todo_cases),
            "\n  ".join(c["company"] for c in todo_cases),
        )
        if len(todo_cases) == len(cases):
            logger.error(
                "All cases are in TODO state. Populate 'expected_fields' in %s first.",
                args.dataset,
            )
            return 1

    # ── Load or run extractions ──────────────────────────────────────────────
    if args.from_cache:
        cache_path = Path(args.from_cache)
        logger.info("Loading cached extraction results from %s", cache_path)
        with open(cache_path) as f:
            cached = json.load(f)
        extraction_map: Dict[str, Optional[dict]] = {
            e["company"]: e.get("extracted_data") for e in cached.get("extractions", [])
        }
    else:
        try:
            settings = load_settings()
        except Exception as exc:
            logger.error("Failed to load settings: %s", exc)
            return 2

        extraction_map: Dict[str, Optional[dict]] = {}
        for case in cases:
            if "TODO" in (case.get("status") or ""):
                extraction_map[case["company"]] = None
                continue
            extraction_map[case["company"]] = run_extraction(case, settings)

    # ── Evaluate ─────────────────────────────────────────────────────────────
    all_results: List[dict] = []
    for case in cases:
        company = case["company"]
        extracted = extraction_map.get(company)
        if extracted is None:
            logger.warning("[%s] No extraction data — skipping evaluation.", company)
            continue
        field_results = evaluate_case(case, extracted, args.tol)
        all_results.extend(field_results)

    if not all_results:
        logger.warning("No fields evaluated. Check that golden_dataset.json has verified expected_fields.")
        return 1

    # ── Optional JSON output ──────────────────────────────────────────────────
    if args.out:
        out_doc = {
            "extractions": [
                {"company": c, "extracted_data": extraction_map.get(c)}
                for c in extraction_map
            ],
            "results": all_results,
            "summary": {
                "total": len(all_results),
                "pass": sum(1 for r in all_results if r["passed"]),
                "fail": sum(1 for r in all_results if not r["passed"]),
            },
        }
        with open(args.out, "w") as f:
            json.dump(out_doc, f, indent=2, default=str)
        logger.info("Full results written to %s", args.out)

    # ── Print report ──────────────────────────────────────────────────────────
    all_pass = _print_report(all_results)
    return 0 if all_pass else 1


if __name__ == "__main__":
    sys.exit(main())
