# Comprehensive Refactoring & Engineering Specification: ALFA (TP BD Automation)

> **Context for Claude / Refactoring Engineer:**
> You are tasked with refactoring and upgrading **ALFA (Automated Lead & Financial Analysis)**, an enterprise-grade automation tool built with Python, Streamlit, Google Gemini GenAI SDK, and OpenPyXL. 
>
> The system automates extracting corporate financials, shareholding, related-party transactions (RPT), and litigations from 100+ page Annual Report (AR) PDFs to populate Transfer Pricing (TP) Business Development Fact Sheet Excel templates.
>
> Your goal is to transform this prototype into a **production-ready, fail-safe, modular, efficient, visually impressive, and easily deployable** enterprise solution.

---

## 📑 Table of Contents
1. [Executive Summary & Current Architecture Audit](#1-executive-summary--current-architecture-audit)
2. [Key Problem Areas & Vulnerabilities](#2-key-problem-areas--vulnerabilities)
3. [Target Architecture & Module Organization](#3-target-architecture--module-organization)
4. [Refactoring Workstreams & Implementation Guidelines](#4-refactoring-workstreams--implementation-guidelines)
   - [Workstream 1: Efficiency & Token Optimization](#workstream-1-efficiency--token-optimization)
   - [Workstream 2: Fail-Safe Reliability & Resilience](#workstream-2-fail-safe-reliability--resilience)
   - [Workstream 3: Template Engine & Dynamic Excel Handling](#workstream-3-template-engine--dynamic-excel-handling)
   - [Workstream 4: Enterprise UI/UX & Interactivity](#workstream-4-enterprise-uiux--interactivity)
   - [Workstream 5: Deployment, Containerization & CI/CD](#workstream-5-deployment-containerization--cicd)
5. [Proposed Directory Structure](#5-proposed-directory-structure)
6. [Step-by-Step Refactoring Plan for Claude](#6-step-by-step-refactoring-plan-for-claude)
7. [Verification & Test Strategy](#7-verification--test-strategy)

---

## 1. Executive Summary & Current Architecture Audit

### Current Workflow
1. **User Uploads**: One Excel template (`.xlsx`) and multiple Annual Report PDFs (`.pdf`).
2. **Text Extraction**: Uses `pdfplumber` to extract raw text from all pages sequentially. If extracted text is under 2,000 characters, it falls back to Gemini Multimodal File API.
3. **LLM Querying**: Passes the entire text dump (or uploaded PDF) into Gemini using `SKILL.md` as system instructions and Pydantic `CompanyData` schema.
4. **Excel Injection**:
   - `safe_insert_rows.py`: Shifts rows if shareholding has > 2 rows by clearing and recreating all sheet cells.
   - `fill_factsheet.py`: Writes values and comments to hardcoded cell coordinates (`C3`, `D12`, `F12`, etc.), caps RPT items at 7, caps Litigations at 2, and hides empty rows.
5. **Download**: Provides individual download buttons for each processed workbook.

---

## 2. Key Problem Areas & Vulnerabilities

### A. Performance & Token Inefficiency
- **Full Text Flooding**: Passing entire 100–200 page annual reports (50k–150k+ tokens) into Gemini for each document creates high latency, excessive API cost, and risk of context truncation or attention degradation.
- **Sequential Processing**: Annual reports are processed one by one synchronously in a blocking `for` loop. No multi-threading / asynchronous execution.
- **Redundant File Reads**: `SKILL.md` is re-read from disk on every iteration.

### B. Fragile Error Handling & Silent Failures
- **Broad Exception Swallowing**: `try-except` blocks print a generic `st.error(e)` and execute `continue`, without saving partial results, logging tracebacks, or cleaning up resources.
- **No Rate Limit / Retry Strategy**: No exponential backoff for Gemini API rate limits (`429 Too Many Requests`, `503 Service Unavailable`).
- **Encrypted/Corrupt PDF Handling**: Fails abruptly if a PDF is password-protected, encrypted, or corrupted.

### C. Rigid Excel Template Assumptions
- **Hardcoded Coordinates**: Fixed mappings (e.g. `C3`, `D12`, `F12`) will break silently if a client provides a template with slightly different header rows.
- **Hardcoded Truncation Capacities**: RPT rows are strictly limited to `rpt_capacity = 7` and Litigation to `lit_capacity = 2`. Any excess items are dropped via `break` without notifying the user.
- **Destructive Sheet Recreation**: `safe_insert_rows.py` snapshots and clears the entire worksheet, which can discard workbook-level metadata, drawing elements, charts, and conditional formatting.

### D. UI/UX Limitations
- **No Pre-Download Inspection**: Users cannot review, edit, or override the extracted JSON before it is written to Excel.
- **No Bulk Export**: When uploading 10 ARs, the user must click 10 separate download buttons rather than having a single "Download All as ZIP" option.
- **Basic State Management**: Reloading the Streamlit page or changing a widget can re-trigger extractions and waste API quota.

---

## 3. Target Architecture & Module Organization

Refactor the monolithic script into a clean, layered architecture:

```
┌────────────────────────────────────────────────────────┐
│               Streamlit Web App (UI/UX)                │
│   (Uploads, Live Progress, Editable Preview, Zip DL)   │
└──────────────────────────┬─────────────────────────────┘
                           │
┌──────────────────────────▼─────────────────────────────┐
│           Orchestration & Task Queue Engine            │
│  (Concurrent Workers, Session State, Batch Pipeline)   │
└──────────────┬──────────────────────────┬──────────────┘
               │                          │
┌──────────────▼──────────┐   ┌───────────▼──────────────┐
│  Document Preprocessor  │   │  LLM Extraction Engine   │
│ - PDF Validation        │   │ - Gemini Client Manager  │
│ - Section Segmentation  │   │ - Backoff & Rate Limiter │
│ - OCR / Vision Fallback │   │ - Pydantic Validators    │
└─────────────────────────┘   └───────────┬──────────────┘
                                          │
┌─────────────────────────────────────────▼──────────────┐
│             Excel Fact Sheet Synthesizer               │
│  - Configurable Template Coordinate Mapper             │
│  - Dynamic Non-Destructive Table Insertion             │
│  - Formula-Safe Shifting & Auto-Comments Injection     │
└────────────────────────────────────────────────────────┘
```

---

## 4. Refactoring Workstreams & Implementation Guidelines

### Workstream 1: Efficiency & Token Optimization

1. **Targeted PDF Chunking / Section Filtering**:
   - Implement a fast pre-filter scanning PDF bookmarks/TOC or page headings for critical keywords:
     - `Balance Sheet`, `Profit and Loss`, `Statement of Accounts`
     - `Related Party Disclosures` / `RPT`
     - `Contingent Liabilities` / `Litigation` / `Claims against Company`
     - `Shareholding Pattern` / `Significant Beneficial Ownership`
   - Extract only the relevant page ranges (typically 15–30 pages instead of 150 pages) to reduce token consumption by **70–80%** and speed up response times.
2. **Concurrent Batch Processing**:
   - Use `concurrent.futures.ThreadPoolExecutor` or `asyncio` to process multiple AR PDFs simultaneously up to a configurable concurrency limit (e.g. `max_workers=3`).
3. **Prompt & Schema Caching**:
   - Cache `SKILL.md` system prompt and Pydantic schemas in memory (e.g., using `@st.cache_data` or a singleton configuration class).

---

### Workstream 2: Fail-Safe Reliability & Resilience

1. **API Resilience with Tenacity**:
   - Wrap Gemini API calls with `tenacity` retry decorators:
     ```python
     from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type
     
     @retry(
         stop=stop_after_attempt(5),
         wait=wait_exponential(multiplier=1, min=2, max=30),
         reraise=True
     )
     def call_gemini_with_retry(...):
         ...
     ```
2. **Structured Logging**:
   - Replace raw `print` statements with Python's standard `logging` library configured for console and rotating file logs.
3. **Comprehensive PDF & Template Validation**:
   - Check if PDF is encrypted or empty before starting.
   - Validate that the uploaded Excel template has the expected structure and sheet names before making API calls.
4. **Graceful Fallbacks & Partial Recovery**:
   - If an extraction for an individual company fails, log the full traceback, mark that entity as failed with a user-friendly error card in the UI, and allow remaining files to continue.

---

### Workstream 3: Template Engine & Dynamic Excel Handling

1. **Configurable Mapping (`TemplateConfig`)**:
   - Decouple coordinate addresses into a clean dataclass/dictionary schema rather than scattering hardcoded cell strings across the code:
     ```python
     @dataclass
     class FactSheetMapping:
         hq_india_cell: str = "C3"
         turnover_fy25_cell: str = "D12"
         shareholding_start_row: int = 42
         rpt_start_row_offset: int = 2
         # ...
     ```
2. **True Dynamic Capacity for RPT and Litigation**:
   - Instead of truncating RPT at 7 and Litigation at 2, calculate required rows dynamically and expand both tables using safe row insertion.
3. **Non-Destructive Row Insertion**:
   - Optimize `safe_insert_rows.py` to insert rows without destroying worksheet properties, retaining conditional formatting, data validations, and named styles.

---

### Workstream 4: Enterprise UI/UX & Interactivity

1. **Modern Layout & Theming**:
   - Implement clean metrics cards (e.g., Total Processed, Succeeded, Failed, Tokens Used).
   - Add status badges (`Pending`, `Processing`, `Success`, `Failed`).
2. **Interactive Preview & Human-in-the-Loop Overrides**:
   - Allow users to inspect extracted financial figures, shareholding tables, and RPTs in an interactive `st.data_editor` before generating the final Excel file.
3. **Bulk Download as ZIP**:
   - Provide a single **"Download All Workbooks (.zip)"** button using `io.BytesIO` and Python's `zipfile` module.
4. **Session State Persistence**:
   - Store results in `st.session_state` so widget interactions or tab switches do not lose extracted data.

---

### Workstream 5: Deployment, Containerization & CI/CD

1. **Docker Support**:
   - Provide a multi-stage `Dockerfile` and `docker-compose.yml`.
2. **Environment Variable Configuration**:
   - Support `GEMINI_API_KEY`, `GEMINI_MODEL`, `LOG_LEVEL`, `MAX_WORKERS`, `CONCURRENCY_LIMIT` via `.env` and OS environment variables.
3. **Automated Testing Suite**:
   - Add `pytest` test cases for:
     - Schema validation.
     - Formula reference shifting in Excel.
     - Text vs Vision fallback heuristics.
     - Missing field fallbacks.

---

## 5. Proposed Directory Structure

```
BD_automation/
├── .env.example
├── .gitignore
├── Dockerfile
├── docker-compose.yml
├── requirements.txt
├── README.md
├── config/
│   ├── __init__.py
│   ├── settings.py            # Global settings, environment variables, LLM configs
│   └── template_mapping.py    # Template coordinate schemas and field definitions
├── core/
│   ├── __init__.py
│   ├── extractor.py           # Gemini API client, retry logic, prompt builders
│   ├── pdf_processor.py       # PDF parsing, section extraction, image/vision routing
│   ├── schemas.py             # Pydantic data models with robust validators
│   ├── excel_engine.py        # Cell population, formula auditing, source comments
│   └── safe_sheet.py          # Non-destructive row insertion & formula rewriting
├── ui/
│   ├── __init__.py
│   ├── components.py          # UI cards, preview tables, download handlers
│   └── styles.py              # Custom CSS and theming
├── tests/
│   ├── __init__.py
│   ├── test_excel_engine.py   # Test cell writing and comment formatting
│   ├── test_formula_shift.py  # Test regex formula shifts and table expansion
│   └── test_schemas.py        # Test schema parsing and edge-case validation
├── SKILL.md                   # Extraction system prompt & TP rules
└── app.py                     # Streamlit entry point
```

---

## 6. Step-by-Step Refactoring Plan for Claude

When refactoring this project, execute the following steps in order:

### Phase 1: Core & Schema Refactoring
- [ ] Move Pydantic models from `app.py` to `core/schemas.py` with custom field validators (handling nulls, string-to-float conversions, and negative sign conventions).
- [ ] Move environment variable parsing and LLM configuration to `config/settings.py`.
- [ ] Implement `core/pdf_processor.py` with safe PDF opening, page count checks, encrypted PDF detection, and intelligent section filtering.

### Phase 2: LLM Client & Resilience
- [ ] Create `core/extractor.py` utilizing `google-genai` with `tenacity` exponential backoff retries.
- [ ] Implement structured error catching and logging for API failures, token overages, and schema deserialization errors.

### Phase 3: Excel Engine Modernization
- [ ] Refactor `safe_insert_rows.py` into `core/safe_sheet.py` with safer formula regex, boundary validation, and style cloning.
- [ ] Refactor `fill_factsheet.py` into `core/excel_engine.py` using `FactSheetMapping` configuration to support dynamic RPT/Litigation capacities without hardcoded cuts.

### Phase 4: UI/UX & Workflow Enhancement
- [ ] Refactor `app.py` to use `st.session_state`.
- [ ] Add editable review tables (`st.data_editor`) for extracted figures.
- [ ] Add ZIP packaging for bulk downloads (`zipfile`).
- [ ] Implement clear progress bars with company-specific status indicators.

### Phase 5: Containerization & Tests
- [ ] Create `Dockerfile` and `docker-compose.yml`.
- [ ] Write unit tests under `tests/` and verify with `pytest`.
- [ ] Update `README.md` to reflect the new modular architecture.

---

## 7. Verification & Test Strategy

| Test Domain | Target Scenario | Verification Method |
| :--- | :--- | :--- |
| **PDF Extraction** | Digital PDF (normal) vs Scanned PDF (<2,000 chars) | Mock PDF input and verify correct routing (Text vs Vision API) |
| **Schema Validation** | Missing fields, null values, malformed data | Pytest passing edge-case JSONs into Pydantic models |
| **Formula Shifting** | Insert 5 rows into shareholding table | Verify Excel formulas (`=SUM(D42:D48)`) shift accurately |
| **RPT & Litigation** | Company with 12 RPT items and 5 litigations | Confirm no rows are dropped and template expands dynamically |
| **API Failure** | Simulated 429 / 503 HTTP status | Verify exponential backoff and retry mechanism |
| **Bulk Zip Export** | 5 simultaneous AR PDFs | Confirm zip file contains 5 correctly named and populated `.xlsx` files |

---

*This specification serves as the comprehensive engineering blueprint for refactoring ALFA into an enterprise-grade solution.*
