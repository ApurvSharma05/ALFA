# ALFA: Automated Lead & Financial Analysis 📊

### Transfer Pricing Business Development Fact Sheet Automation

[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![Streamlit](https://img.shields.io/badge/UI-Streamlit-FF4B4B.svg)](https://streamlit.io/)
[![Google Gemini](https://img.shields.io/badge/AI-Google%20Gemini-4285F4.svg)](https://ai.google.dev/)
[![OpenPyXL](https://img.shields.io/badge/Excel-openpyxl-217346.svg)](https://openpyxl.readthedocs.io/)
[![Pydantic](https://img.shields.io/badge/Schema-Pydantic%20v2-E92063.svg)](https://docs.pydantic.dev/)
[![Docker](https://img.shields.io/badge/Deploy-Docker-2496ED.svg)](https://www.docker.com/)

---

## Overview

**ALFA (Automated Lead & Financial Analysis)** is an AI-powered document processing system that automates the extraction of financial, tax, and corporate information from **Annual Reports and Financial Statement PDFs** and populates standardized **Transfer Pricing (TP) Business Development Fact Sheet** Excel templates.

The system was developed during an **EY internship by Harshiv and Apurv** to reduce the manual effort involved in reviewing lengthy annual reports, locating relevant financial and related-party information, applying Transfer Pricing-specific extraction rules, and transferring the results into structured Excel workbooks.

Instead of manually reviewing 100+ page reports and copying information into predefined templates, ALFA automates the workflow:

```text
Annual Report PDF
       ↓
Document Processing
       ↓
Relevant Section Detection
       ↓
AI-Powered Extraction
       ↓
Schema Validation
       ↓
Human Review
       ↓
Excel Fact Sheet
       ↓
Source-Cited Output
```

---

## Key Features

### AI-Powered Document Extraction

ALFA uses Google Gemini to extract structured financial and corporate information from annual reports.

The extraction pipeline supports two modes:

**Text Extraction**

* Uses `pdfplumber` for native PDF text extraction.
* Scores pages using domain-specific keywords.
* Sends only relevant pages to the LLM.
* Reduces unnecessary token usage and processing overhead.

**Vision Fallback**

* Detects scanned or image-heavy PDFs.
* Automatically switches to Gemini's multimodal capabilities when sufficient text cannot be extracted.
* Allows the same workflow to handle both digital and scanned reports.

### Structured Data Validation

Extracted information is converted into strongly typed Pydantic models.

Validation handles:

* Numeric values
* Missing values
* Invalid fields
* Empty company names
* Type coercion
* Nullable fields
* Structured financial data

This prevents raw LLM output from being written directly into Excel without validation.

### Transfer Pricing-Specific Extraction Rules

The system is designed around domain-specific TP requirements rather than generic document extraction.

Examples include:

* Revenue from Operations used for turnover.
* Other Income excluded from turnover calculations.
* Total Expenses used for total cost.
* Foreign Associated Enterprise service revenue extracted from relevant RPT disclosures.
* Director shareholders identified individually.
* Other promoters consolidated where applicable.
* Related-party transactions grouped by nature.
* Corporate entities separated from individual parties.
* Litigation extracted specifically from the Contingent Liabilities section.
* Extracted values linked to their source locations.

These rules are maintained separately in `SKILL.md`, allowing the extraction behaviour to be updated without rewriting the application logic.

### Source-Cited Excel Output

Every populated value can include an Excel cell comment containing the source of the extracted information.

For example:

```text
Source: Annual Report, Page 87
Section: Related Party Transactions
```

This provides traceability between the generated fact sheet and the original annual report.

### Non-Destructive Excel Processing

ALFA populates existing corporate Excel templates without unnecessarily altering their formatting.

The Excel engine preserves:

* Fonts
* Cell colors
* Borders
* Number formats
* Merged cells
* Existing formulas
* Worksheet structure

The system also supports dynamic row expansion when extracted data exceeds the original template capacity.

Formula references are adjusted when rows are inserted to avoid breaking dependent calculations.

### Human-in-the-Loop Review

AI-generated results are not immediately treated as final.

The Streamlit interface allows users to review and edit extracted:

* Financial information
* Shareholding information
* Related-party transactions
* Litigation information

before generating the final workbook.

This provides a practical human-in-the-loop workflow for high-accuracy business document processing.

### Batch Processing

Multiple annual reports can be processed in a single session.

`ThreadPoolExecutor` enables configurable concurrent processing while maintaining individual success/failure states for each company.

The application provides:

* Total files processed
* Successful extractions
* Failed extractions
* Per-company status
* Bulk ZIP download

### API Resilience

The Gemini extraction layer includes retry handling using `tenacity`.

The system handles transient:

* `429` rate-limit errors
* `503` service availability errors

using exponential backoff.

Temporary uploaded files are cleaned up using `finally` blocks.

### Docker Support

The application includes Docker configuration for reproducible deployment.

```text
Dockerfile
docker-compose.yml
```

---

## Architecture

```text
┌─────────────────────────────────────────────────────┐
│                  Streamlit Web App                  │
│                                                     │
│ Uploads · Processing · Preview · Editing · Export   │
└─────────────────────────┬───────────────────────────┘
                          │
                          ▼
┌─────────────────────────────────────────────────────┐
│              Orchestration Layer                    │
│                                                     │
│ Session State · Processing Phases · Batch Jobs      │
│ ThreadPoolExecutor                                  │
└───────────────┬─────────────────────┬───────────────┘
                │                     │
                ▼                     ▼
┌────────────────────────┐  ┌─────────────────────────┐
│ PDF Processing         │  │ AI Extraction           │
│                        │  │                         │
│ pdfplumber             │  │ Google Gemini            │
│ Page filtering         │  │ Structured extraction   │
│ PDF validation         │  │ Retry handling           │
│ Vision fallback        │  │ Pydantic validation     │
└────────────┬───────────┘  └────────────┬────────────┘
             │                           │
             └──────────────┬────────────┘
                            ▼
┌─────────────────────────────────────────────────────┐
│                 Excel Engine                        │
│                                                     │
│ FactSheet Mapping · Data Population                 │
│ Dynamic Row Expansion · Source Comments             │
└─────────────────────────┬───────────────────────────┘
                          │
                          ▼
┌─────────────────────────────────────────────────────┐
│                 Safe Sheet Layer                    │
│                                                     │
│ Formula Shifting · Merge Preservation               │
│ Style Cloning · Validation                          │
└─────────────────────────┬───────────────────────────┘
                          │
                          ▼
                   Final Excel Fact Sheet
```

---

## Repository Structure

```text
BD_automation/
│
├── app.py
├── SKILL.md
├── requirements.txt
├── Dockerfile
├── docker-compose.yml
├── .env.example
├── .gitignore
│
├── config/
│   ├── __init__.py
│   ├── settings.py
│   └── template_mapping.py
│
├── core/
│   ├── __init__.py
│   ├── schemas.py
│   ├── pdf_processor.py
│   ├── extractor.py
│   ├── excel_engine.py
│   └── safe_sheet.py
│
├── ui/
│   ├── __init__.py
│   ├── styles.py
│   └── components.py
│
├── tests/
│   ├── __init__.py
│   ├── test_schemas.py
│   ├── test_formula_shift.py
│   └── test_excel_engine.py
│
├── archive/
│   └── ...
│
└── README.md
```

---

## Technology Stack

| Technology        | Purpose                                          |
| ----------------- | ------------------------------------------------ |
| **Python 3.10+**  | Core application                                 |
| **Google Gemini** | AI-powered document understanding and extraction |
| **pdfplumber**    | Native PDF text extraction                       |
| **OpenPyXL**      | Excel template manipulation                      |
| **Pydantic v2**   | Structured schemas and validation                |
| **Tenacity**      | API retry and exponential backoff                |
| **Streamlit**     | Web interface                                    |
| **Pandas**        | Data preview and manipulation                    |
| **Pytest**        | Automated testing                                |
| **Docker**        | Containerized deployment                         |

---

## Installation

### Prerequisites

* Python 3.10+
* Google Gemini API key
* Git

### 1. Clone the repository

```bash
git clone https://github.com/Harshiv15/BD_automation.git
cd BD_automation
```

### 2. Create a virtual environment

Windows PowerShell:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

Linux/macOS:

```bash
python -m venv .venv
source .venv/bin/activate
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

### 4. Configure environment variables

Copy the example environment file:

```bash
cp .env.example .env
```

Configure:

```env
GEMINI_API_KEY=your_gemini_api_key
GEMINI_MODEL=gemini-flash-latest
```

Never commit API keys or other secrets to Git.

---

## Running the Application

Start the Streamlit application:

```bash
streamlit run app.py
```

The application will be available at:

```text
http://localhost:8501
```

---

## Docker

Build and start the application:

```bash
docker-compose up --build
```

Then open:

```text
http://localhost:8501
```

---

## Workflow

### 1. Upload the Fact Sheet Template

Upload the standardized `.xlsx` template that should be populated.

### 2. Upload Annual Reports

Select one or more annual report `.pdf` files.

### 3. Start Extraction

ALFA:

1. Validates the PDFs.
2. Extracts native text where possible.
3. Identifies relevant pages.
4. Uses Gemini for structured extraction.
5. Validates the result using Pydantic.
6. Falls back to multimodal processing when necessary.

### 4. Review the Results

Review the extracted information through the Streamlit interface.

Values can be edited before the final workbook is generated.

### 5. Generate Fact Sheets

The validated data is written into the original Excel template while preserving its formatting and structure.

Individual files or a combined ZIP archive can then be downloaded.

---

## Domain Extraction Rules

The extraction logic follows Transfer Pricing-specific business rules defined in `SKILL.md`.

| Category             | Rule                                                             |
| -------------------- | ---------------------------------------------------------------- |
| **Turnover**         | Revenue from Operations only; Other Income excluded              |
| **Total Cost**       | Total Expenses                                                   |
| **AE Revenue Split** | Foreign AE service revenue from relevant RPT disclosures         |
| **Shareholding**     | Director shareholders individually; other promoters consolidated |
| **Related Parties**  | Corporate entities extracted and grouped by transaction nature   |
| **Litigation**       | Extracted strictly from Contingent Liabilities disclosures       |
| **Source Comments**  | Populated cells include source information where available       |

The goal is not simply to extract text, but to apply the business rules required for the downstream TP fact sheet.

---

## Testing

Run the test suite with:

```bash
python -m pytest tests/ -v
```

Current tests cover areas including:

* Pydantic schema validation
* Formula shifting
* Excel row insertion
* Excel engine behaviour

---

## Design Principles

ALFA is built around five core principles:

**1. Relevant context over brute-force processing**

Only relevant portions of large annual reports are prioritized for extraction whenever native text is available.

**2. Structured output over raw LLM responses**

LLM responses are validated against explicit schemas before entering the Excel layer.

**3. Traceability**

Extracted information should remain traceable to the source document.

**4. Human-in-the-loop validation**

Users can inspect and correct AI-generated results before final output.

**5. Preserve existing business templates**

The system populates existing Excel workbooks rather than forcing users into a new output format.

---

## Future Improvements

Potential areas for further development include:

* Improved semantic section detection
* More advanced document-level OCR
* Extraction confidence scoring
* Automated anomaly detection
* Expanded evaluation datasets
* Additional financial statement formats
* Enterprise authentication
* Audit logging
* Cloud-based document storage
* Fine-grained extraction provenance

---

## Authors

**Harshiv & Apurv**

Developed as an **EY Internship Project** focused on applying AI and automation to real-world Transfer Pricing business processes.

---

## Disclaimer

This project was developed as an internship automation project. It is intended to assist with document processing and data preparation and does not replace professional review or judgment.
