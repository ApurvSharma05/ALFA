# ALFA: Automated Lead & Financial Analysis

ALFA is a Streamlit-based workflow for processing annual report PDFs and populating transfer-pricing business development fact-sheet templates in Excel.

## Overview

The application reads one or more annual report PDFs, filters the relevant pages, extracts structured financial and corporate data using Google Gemini, validates the output with Pydantic, and writes the result into an existing Excel template. A human review step is included before final generation.

Typical workflow:

1. Upload a BD template Excel file
2. Upload one or more annual report PDFs
3. Run extraction
4. Review and edit extracted values in the app
5. Generate the final workbook and download the result

## Features

- PDF validation and page filtering
- Optional vision fallback for scanned or image-heavy PDFs
- Gemini-powered extraction for financial and related-party data
- Schema validation using Pydantic
- Human-in-the-loop editing in Streamlit
- Excel fact-sheet population while preserving workbook structure
- Batch processing for multiple company PDFs
- Download of individual Excel files or a ZIP archive

## Tech stack

- Python 3.10+
- Streamlit
- Google GenAI / Gemini
- pdfplumber
- openpyxl
- Pydantic
- tenacity
- pytest
- Docker

## Repository structure

```text
ALFA/
├── app.py
├── SKILL.md
├── requirements.txt
├── Dockerfile
├── docker-compose.yml
├── .env
├── .env.example
├── .gitignore
├── config/
│   ├── __init__.py
│   ├── settings.py
│   └── template_mapping.py
├── core/
│   ├── __init__.py
│   ├── extractor.py
│   ├── excel_engine.py
│   ├── pdf_processor.py
│   ├── safe_sheet.py
│   └── schemas.py
├── data/
│   ├── inputs/
│   ├── outputs/
│   └── templates/
├── scripts/
│   ├── create_bd_template.py
│   └── generate_report.py
├── tests/
│   ├── __init__.py
│   ├── test_excel_engine.py
│   ├── test_formula_shift.py
│   └── test_schemas.py
├── ui/
│   ├── __init__.py
│   ├── components.py
│   └── styles.py
└── README.md
```

## Setup

### Prerequisites

- Python 3.10+
- A Google Gemini API key
- Git

### 1. Clone the repository

```bash
git clone <repo-url>
cd ALFA
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

Copy the sample environment file:

```bash
copy .env.example .env
```

Or on Linux/macOS:

```bash
cp .env.example .env
```

Then update the values in `.env`:

```env
GEMINI_API_KEY=your_gemini_api_key_here
GEMINI_MODEL=gemini-flash-latest
MAX_WORKERS=3
VISION_FALLBACK_THRESHOLD=2000
MAX_RELEVANT_PAGES=40
SKILL_PROMPT_PATH=SKILL.md
```

The app also accepts `GOOGLE_API_KEY` as a fallback for the Gemini key because the settings loader checks both environment variable names.

## Run the app

```bash
streamlit run app.py
```

Then open the local URL shown by Streamlit, typically:

```text
http://localhost:8501
```

## Docker

Build and run with Docker Compose:

```bash
docker compose up --build
```

Then open:

```text
http://localhost:8501
```

## How it works

1. The app validates each uploaded PDF.
2. It extracts native text and ranks relevant pages.
3. If text is too weak or the PDF is scanned, the app uses the vision-based fallback path.
4. Gemini extracts structured data using the business rules from `SKILL.md`.
5. The result is validated against the Pydantic schemas.
6. Users review the extracted values in the interface.
7. The final Excel workbook is generated and downloaded.

## Testing

Run the project tests with:

```bash
python -m pytest tests -v
```

The current suite covers most of the core logic related to schemas, formula shifting, and Excel manipulation.

## Notes

- The extraction rules are intentionally domain-specific and are maintained in [SKILL.md](SKILL.md).
- The app is designed for practical human review rather than fully autonomous output generation.
- This project is intended to support transfer-pricing and business-development fact-sheet workflows, not replace professional judgment.

## Authors

- Harshiv
- Apurv

Developed as a business automation project for annual-report and transfer-pricing fact-sheet processing.

## Disclaimer

This project is a document-processing aid for preparation and review. It does not replace professional analysis, legal review, or accounting judgment.
