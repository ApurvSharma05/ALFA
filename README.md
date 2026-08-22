# ALFA: Automated Lead & Financial Analysis 📊
### Transfer Pricing (TP) Business Development Fact Sheet Automation

[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![Streamlit](https://img.shields.io/badge/UI-Streamlit-FF4B4B.svg)](https://streamlit.io/)
[![Google GenAI](https://img.shields.io/badge/AI-Google%20Gemini-4285F4.svg)](https://ai.google.dev/)
[![OpenPyXL](https://img.shields.io/badge/Excel-openpyxl-217346.svg)](https://openpyxl.readthedocs.io/)
[![Pydantic](https://img.shields.io/badge/Schema-Pydantic%20v2-E92063.svg)](https://docs.pydantic.dev/)
[![Docker](https://img.shields.io/badge/Deploy-Docker-2496ED.svg)](https://www.docker.com/)

---

## 📌 Overview

**ALFA (Automated Lead & Financial Analysis)** is an intelligent document processing and financial analysis tool designed to automate the extraction and synthesis of complex corporate data from **Annual Reports (ARs) / Financial Statement PDFs** directly into standardized **Transfer Pricing (TP) Business Development (BD) Fact Sheet** Excel workbooks.

Developed during an EY Internship project by **Harshiv & Apurv**, ALFA addresses the tedious, error-prone manual process of reading 100+ page annual reports, extracting nuanced tax and financial data points, applying strict Transfer Pricing rules, and formatting them into strict corporate client templates.

---

## 🚀 Key Features

- **Dual-Mode Extraction Engine (Text + Vision Fallback)**:
  - **Primary (Text)**: Keyword-scored page filtering + native text extraction via `pdfplumber` — sends only the most relevant ~40 pages to the LLM, reducing token consumption by **70–80%**.
  - **Fallback (Vision)**: Automatic heuristic detection (< 2,000 extracted characters) for scanned or image-heavy PDFs, routing seamlessly to Gemini Multimodal File API.
- **Strict Structured JSON Schema**:
  - Built with **Pydantic v2** field validators for numeric coercion, empty-name rejection, and graceful null handling.
- **Fail-Safe API Resilience**:
  - Exponential-backoff retries via **tenacity** (5 attempts, 2–30s wait) for `429` / `503` errors.
  - Structured logging throughout. Uploaded files cleaned up in `finally` blocks.
- **Concurrent Batch Processing**:
  - `ThreadPoolExecutor` with configurable concurrency (1–5 workers) for parallel PDF extraction.
- **Non-Destructive Excel Manipulation**:
  - **Zero Style Degradation**: Preserves all native fonts, cell colors, borders, and number formats.
  - **Dynamic Row Expansion**: Formula-safe row shifting via regex, merged-cell rewriting, and style cloning.
  - **Source-Citing Cell Comments**: Every populated cell gets an Excel comment citing the exact AR page/section.
- **Enterprise Streamlit Web UI**:
  - Metric cards (Total, Succeeded, Failed), per-company status badges.
  - **Human-in-the-loop previews**: Edit extracted financials, shareholding, RPT, and litigation in `st.data_editor` before generating Excel.
  - **Bulk ZIP download**: One-click archive of all populated workbooks.
  - Session state persistence across reruns.
- **Docker-Ready Deployment**:
  - Multi-stage `Dockerfile` and `docker-compose.yml` included.

---

## 🏗️ Architecture

```
┌─────────────────────────────────────────────────────┐
│             Streamlit Web App (app.py)               │
│   Uploads · Metrics · Preview Editor · ZIP Download  │
└────────────────────────┬────────────────────────────┘
                         │
┌────────────────────────▼────────────────────────────┐
│         Orchestration & Session State Engine         │
│    ThreadPoolExecutor · st.session_state · Phases    │
└──────────┬──────────────────────────┬───────────────┘
           │                          │
┌──────────▼────────────┐   ┌────────▼────────────────┐
│ core/pdf_processor.py │   │   core/extractor.py      │
│ - PDF Validation      │   │ - GeminiExtractor class  │
│ - Section Filtering   │   │ - Tenacity retries       │
│ - OCR/Vision routing  │   │ - File cleanup in finally│
└───────────────────────┘   └────────┬────────────────┘
                                     │
┌────────────────────────────────────▼────────────────┐
│              core/excel_engine.py                    │
│  - FactSheetMapping config · FillResult metadata     │
│  - Dynamic table expansion · Source comments         │
└────────────────┬────────────────────────────────────┘
                 │
┌────────────────▼────────────────────────────────────┐
│              core/safe_sheet.py                       │
│  - Formula regex shifting · Merge preservation       │
│  - Style cloning · Input validation + logging        │
└─────────────────────────────────────────────────────┘
```

---

## 📁 Repository Structure

```
BD_automation/
├── app.py                         # Streamlit entry point (thin orchestration)
├── SKILL.md                       # Extraction system prompt & TP rules
├── requirements.txt               # Python dependencies
├── Dockerfile                     # Multi-stage Docker build
├── docker-compose.yml             # Docker Compose configuration
├── .env.example                   # Environment variable template
├── .gitignore                     # Git ignore rules
├── config/
│   ├── __init__.py
│   ├── settings.py                # Centralized settings (env vars, logging)
│   └── template_mapping.py        # FactSheetMapping dataclass (cell coords)
├── core/
│   ├── __init__.py
│   ├── schemas.py                 # Pydantic models with field validators
│   ├── pdf_processor.py           # PDF validation, text extraction, filtering
│   ├── extractor.py               # Gemini client with retries
│   ├── excel_engine.py            # Excel population engine
│   └── safe_sheet.py              # Formula-safe row insertion
├── ui/
│   ├── __init__.py
│   ├── styles.py                  # Custom CSS theming
│   └── components.py              # Metric cards, previews, ZIP download
├── tests/
│   ├── __init__.py
│   ├── test_schemas.py            # Schema validation tests
│   ├── test_formula_shift.py      # Formula shift & row insertion tests
│   └── test_excel_engine.py       # Excel engine integration tests
├── archive/                       # Historical iterations & references
│   ├── fill_factsheet_v1.py
│   ├── safe_insert_rows_v1.py
│   └── ...
└── README.md
```

---

## ⚙️ Installation & Setup

### 1. Prerequisites
- Python 3.10 or higher
- Google Gemini API Key ([Get an API Key here](https://aistudio.google.com/))

### 2. Clone & Create Virtual Environment
```bash
git clone https://github.com/Harshiv15/BD_automation.git
cd BD_automation
python -m venv .venv

# Activate (Windows PowerShell):
.\.venv\Scripts\Activate.ps1
# Activate (Linux / macOS):
source .venv/bin/activate
```

### 3. Install Dependencies
```bash
pip install -r requirements.txt
```

### 4. Configure Environment Variables
```bash
cp .env.example .env
```
Edit `.env` and paste your Gemini API key:
```env
GEMINI_API_KEY=your_gemini_api_key_here
GEMINI_MODEL=gemini-flash-latest
```

---

## 🖥️ Usage

### Run Locally
```bash
streamlit run app.py
```

### Run with Docker
```bash
docker-compose up --build
```
Access at [http://localhost:8501](http://localhost:8501).

### Run Tests
```bash
python -m pytest tests/ -v
```

### Step-by-Step Workflow
1. **Enter API Key** in the sidebar (or configure via `.env`)
2. **Upload Template**: Upload the BD Fact Sheet `.xlsx` template
3. **Upload Annual Reports**: Select one or more AR `.pdf` files
4. **Start Extraction**: Click 🚀 — PDFs are processed concurrently via Gemini
5. **Review & Edit**: Inspect and modify extracted values in interactive tables
6. **Generate Excel**: Click 📊 — download individual files or a bulk ZIP

---

## 📋 Domain Rules (SKILL.md)

| Category | Extraction Rule |
| :--- | :--- |
| **Turnover** | Revenue from Operations **only** (excludes Other Income) |
| **Total Cost** | Total Expenses |
| **AE Revenue Split** | Foreign AE service revenues from the RPT section |
| **Shareholding** | Director shareholders individually; other promoters consolidated |
| **Related Party (RPT)** | Corporate entities only; collated by nature; verbatim labels |
| **Litigation** | Strictly from Contingent Liabilities note; zero hallucination |
| **Review Comments** | Every cell gets a source-citing Excel comment |

---

## 📦 Dependencies

| Package | Purpose |
| :--- | :--- |
| `streamlit` | Interactive web application |
| `google-genai` | Google Gemini API SDK |
| `pydantic` | Strict data schema validation |
| `pdfplumber` | PDF text extraction |
| `openpyxl` | Excel workbook manipulation |
| `tenacity` | Retry logic with exponential backoff |
| `python-dotenv` | Environment variable management |
| `pandas` | Data preview tables |
| `pytest` | Automated testing |

---

## 👥 Authors

- **Harshiv** & **Apurv** — Project creators (EY Internship Project)
#   A L F A  
 #   A L F A  
 