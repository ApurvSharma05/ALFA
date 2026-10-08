# Learn ALFA: A Complete Codebase Guide

This document teaches the ALFA project from the outside in. It explains what the application does, how data moves through it, which technologies are used, and what every important function is responsible for.

The guide describes the current implementation in:

- `app.py`
- `config/`
- `core/`
- `ui/`
- `tests/`
- `SKILL.md`

The old `archive/` and removed `docs/` material are not part of the active application.

---

## 1. What ALFA Does

ALFA means **Automated Lead & Financial Analysis**. It is a Streamlit application for preparing Transfer Pricing Business Development Fact Sheets.

A user provides:

1. A BD Fact Sheet Excel template.
2. One or more Annual Report PDFs.
3. A Google Gemini API key.

ALFA then:

1. Validates each PDF.
2. Extracts useful text from the PDF.
3. Selects relevant pages when the PDF is large.
4. Detects whether the PDF is probably scanned.
5. Sends text PDFs to Gemini as text.
6. Sends scanned PDFs to Gemini using its file and vision capability.
7. Requests a strict structured JSON response.
8. Shows the extracted values for human review.
9. Writes reviewed values into an Excel workbook.
10. Adds source comments to populated cells.
11. Provides individual Excel downloads and a ZIP download.

The central principle is that the output is a **draft for human review**, not an automatically final legal or financial conclusion.

```text
Annual Report PDF
      |
      v
PDF validation and text extraction
      |
      v
Relevant-page filtering
      |
      +--> Text extraction path ----> Gemini text request
      |
      +--> Scanned PDF path --------> Gemini vision request
                                      |
                                      v
                           Schema-shaped JSON
                                      |
                                      v
                         Streamlit review editor
                                      |
                                      v
                         Excel population engine
                                      |
                                      v
                    Source-cited Excel draft download
```

---

## 2. The Main Mental Model

There are four layers:

### 2.1 Interface layer: `app.py` and `ui/`

Streamlit displays controls, accepts uploads, shows progress, lets the user edit extracted values, and provides downloads.

### 2.2 Orchestration layer: `app.py`

The root app coordinates the workflow. It does not contain the PDF parser, Gemini implementation, or detailed Excel cell mapping itself.

### 2.3 Domain and infrastructure layer: `core/`

The `core/` package performs the technically sensitive work:

- PDF validation and page selection.
- Gemini requests and retries.
- Extraction data modelling.
- Excel cell writing.
- Safe Excel row insertion and formula shifting.

### 2.4 Configuration layer: `config/` and `SKILL.md`

`config/` holds settings and Excel coordinates. `SKILL.md` is the domain instruction prompt sent to Gemini. It contains Transfer Pricing-specific extraction rules.

---

## 3. Repository Map

```text
ALFA/
|
+-- app.py                         Streamlit entry point and workflow
+-- SKILL.md                       Domain rules and Gemini system prompt
+-- requirements.txt                Python dependencies
+-- README.md                       Project overview and setup notes
+-- Dockerfile                      Container image definition
+-- docker-compose.yml              Container runtime definition
+|
+ +-- config/
+ |   +-- settings.py               Environment-backed settings
+ |   +-- template_mapping.py       Excel cell and row mapping
+ |
+ +-- core/
+ |   +-- pdf_processor.py           PDF validation and page filtering
+ |   +-- extractor.py               Gemini text and vision client
+ |   +-- schemas.py                 Pydantic response schema
+ |   +-- excel_engine.py            Excel population
+ |   +-- safe_sheet.py              Safe row insertion and formula shifting
+ |
+ +-- ui/
+ |   +-- components.py              Reusable Streamlit widgets
+ |   +-- styles.py                  Small CSS layer
+ |
+ +-- tests/
+     +-- test_schemas.py
+     +-- test_excel_engine.py
+     +-- test_formula_shift.py
+```

---

## 4. Technologies Used

### Python

Python is the application language. The code uses type annotations, dataclasses, context managers, exceptions, and modules.

### Streamlit

Streamlit turns the Python script into a browser-based application. Important APIs used here include:

- `st.set_page_config`: page title, icon, layout, and initial sidebar state.
- `st.session_state`: data that survives Streamlit reruns during one user session.
- `st.file_uploader`: receives the Excel template and PDFs.
- `st.sidebar`: settings panel.
- `st.columns`: two-column upload layout and metric cards.
- `st.expander`: collapsible instructions and review sections.
- `st.button`: starts extraction and Excel generation.
- `st.progress`: displays long-running progress.
- `st.data_editor`: lets users correct extracted values.
- `st.download_button`: downloads Excel files and ZIP archives.
- `st.rerun`: moves the interface to the next phase.
- `st.markdown`: renders Markdown and the small custom CSS layer.

Streamlit reruns the script from top to bottom after many interactions. `st.session_state` is therefore essential: it holds extraction results between the extraction, review, and generation phases.

### Google Gemini GenAI SDK

`google-genai` supplies:

- `genai.Client`: API client.
- `client.models.generate_content`: sends prompts and receives model responses.
- `client.files.upload`: uploads scanned PDFs for multimodal processing.
- `client.files.delete`: removes uploaded files after processing.
- `types.GenerateContentConfig`: configures system instructions, temperature, JSON output, and the response schema.

The selected model defaults to `gemini-flash-latest`, but the user can change it in the sidebar or with the `GEMINI_MODEL` environment variable.

### pdfplumber

`pdfplumber` opens PDFs and extracts text page by page. ALFA uses it for:

- Page counting.
- Basic PDF validation.
- Full text extraction.
- Relevance-filtered extraction.

It is effective for native text PDFs. Image-only PDFs usually produce little or no text, which activates the Gemini vision fallback.

### Pydantic v2

Pydantic describes the expected Gemini response. It validates and normalizes values such as numeric strings, missing values, shareholder names, and related-party labels.

The schema is also passed to Gemini through `response_schema=CompanyData`, which helps constrain the model output shape.

### OpenPyXL

`openpyxl` reads and writes `.xlsx` files. ALFA uses it to:

- Open the uploaded template.
- Write values into mapped cells.
- Preserve and use formulas.
- Add Excel comments.
- Hide unused table rows.
- Save the generated workbook.

### Tenacity

`tenacity` retries Gemini calls after failures. The current retry policy allows up to five attempts and uses exponential waits between two and thirty seconds.

### Pandas

`pandas` converts extracted lists into data frames for Streamlit's editable tables.

### python-dotenv

`python-dotenv` loads values from a local `.env` file when available. This supports local configuration without hard-coding API credentials in Python.

### Python standard library

The project also uses:

- `os`: paths and environment variables.
- `json`: serialize and parse extraction data.
- `tempfile`: temporary upload and output directories.
- `logging`: structured application logs.
- `dataclasses`: small typed result and settings objects.
- `pathlib.Path`: filesystem checks.
- `concurrent.futures`: concurrent PDF processing.
- `io` and `zipfile`: in-memory ZIP downloads.
- `re`: Excel formula reference matching.
- `copy.copy`: copy Excel styles safely.

### Pytest

Pytest runs the automated tests. The test suite currently covers schema validation, Excel writing, formula changes, row insertion, warnings, and style preservation.

### Docker

Docker packages the application and dependencies into a reproducible environment. Docker Compose runs the service on port `8501` and mounts `SKILL.md` read-only into the container.

---

## 5. How to Run ALFA

### 5.1 Local Python setup

From PowerShell:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

Start the application:

```powershell
streamlit run app.py
```

Open the URL shown by Streamlit, normally:

```text
http://localhost:8501
```

### 5.2 Environment variables

Create a local `.env` file if desired:

```text
GEMINI_API_KEY=your-key-here
GEMINI_MODEL=gemini-flash-latest
MAX_WORKERS=3
VISION_FALLBACK_THRESHOLD=2000
MAX_RPT_ROWS=15
MAX_LIT_ROWS=10
MAX_RELEVANT_PAGES=40
LOG_LEVEL=INFO
SKILL_PROMPT_PATH=SKILL.md
```

The API key can also be entered in the Streamlit sidebar. Never commit a real API key or `.env` file.

### 5.3 Run the tests

```powershell
.\.venv\Scripts\python.exe -m pytest -q
```

### 5.4 Run with Docker

```powershell
docker compose up --build
```

The app is exposed at `http://localhost:8501`.

---

## 6. `app.py`: The Application Orchestrator

`app.py` is a Streamlit script. It executes from top to bottom on each rerun.

### Imports

The app imports:

- Standard-library tools for paths, JSON, temporary files, logging, and concurrency.
- Streamlit for the interface.
- OpenPyXL for checking the uploaded workbook.
- Settings and mappings from `config/`.
- PDF, Gemini, Excel, and safe-sheet services from `core/`.
- Rendering helpers from `ui/`.

### `st.set_page_config(...)`

This must occur before other Streamlit commands. ALFA sets:

- Page title.
- Material icon.
- Wide layout.
- Expanded initial sidebar.

### `inject_styles()`

This imports and injects the small CSS layer from `ui/styles.py`.

### `_DEFAULTS`

The dictionary defines initial session-state values:

- `extraction_results`: successful company name to extracted data.
- `extraction_errors`: company name to error message.
- `extraction_modes`: company name to `text`, `vision`, or another mode.
- `excel_bytes`: generated workbook bytes by company.
- `fill_warnings`: Excel warnings by company.
- `phase`: `upload`, `extract`, `review`, or `generate`.
- `processing_complete`: whether extraction has completed.

The following loop initializes a key only when it is absent, so a Streamlit rerun does not erase existing session data.

### Sidebar controls

The sidebar collects:

- Gemini API key.
- Gemini model name.
- Maximum concurrent workers.
- Vision fallback threshold.
- Maximum relevant pages.

`load_settings()` supplies defaults, while widget values override them for the current run.

### Upload controls

`st.file_uploader` accepts:

- One `.xlsx` BD template.
- Multiple `.pdf` Annual Reports.

The extraction button is disabled until all required inputs and the API key are present.

### `_process_single_pdf(pdf_path, company_name, extractor, vision_threshold, max_pages)`

This is the main per-file pipeline function.

1. Calls `validate_pdf()`.
2. Returns an error result immediately if the PDF is invalid.
3. Calls `extract_relevant_pages()`.
4. Calls `is_scanned_pdf()` on the extracted text.
5. Uses `extract_from_pdf()` for scanned PDFs.
6. Uses `extract_from_text()` for normal text PDFs.
7. Converts the `ExtractionResult` into a simple dictionary for session state.

Returned success shape:

```python
{
    "company": "Example Ltd",
    "data": {"entity": "Example Ltd", "fields": {...}},
    "mode": "text",
}
```

Returned failure shape:

```python
{
    "company": "Example Ltd",
    "error": "Invalid PDF: ...",
    "mode": "none",
}
```

### Extraction button block

When the user clicks **Start extraction**, the app:

1. Clears old extraction and output state.
2. Sets the phase to `extract`.
3. Creates one `GeminiExtractor`.
4. Saves uploaded PDFs into a temporary directory.
5. Creates one future per PDF with `ThreadPoolExecutor`.
6. Uses `as_completed()` so the UI records whichever company finishes next.
7. Stores successes, errors, and extraction modes in session state.
8. Moves to the `review` phase.
9. Calls `st.rerun()` so the new phase renders cleanly.

The temporary directory is deleted automatically after the `with` block finishes.

### Metrics and company cards

After extraction, `render_metrics_bar()` displays total, success, failure, and pending counts. `render_company_card()` displays the status and extraction mode for each company.

### Review phase

For each successful result, `render_preview_editor()` creates editable tables. The updated dictionary replaces the previous dictionary in session state.

The review phase supports edits to:

- Financial summary fields.
- Shareholder rows.
- Related-party transaction rows.
- Litigation rows.

### Generate Excel button

When the user clicks **Generate Excel files**:

1. A `FactSheetMapping` instance is created.
2. The uploaded template is saved temporarily.
3. OpenPyXL reads the first worksheet name.
4. Each company's reviewed data is saved as temporary JSON.
5. The shareholding table is expanded when required.
6. `fill_factsheet()` writes values into the workbook.
7. Warnings are stored and displayed.
8. The completed workbook is read as bytes into session state.
9. The phase changes to `generate`.
10. The app reruns.

### Download phase

`render_download_section()` creates individual `.xlsx` downloads and, when more than one file succeeded, one ZIP download.

---

## 7. `core/pdf_processor.py`: PDF Processing

### `SECTION_KEYWORDS`

This dictionary maps keywords to relevance weights. Examples:

- `statement of profit and loss`: 10.
- `revenue from operations`: 10.
- `related party`: 10.
- `shareholding pattern`: 9.
- `contingent liabilities`: 10.
- `registered office`: 6.

Higher weights make pages more likely to be selected for Gemini.

### `PDFInfo`

A dataclass describing a PDF:

- `path`: source path.
- `page_count`: number of pages.
- `is_encrypted`: whether the file appears password-protected.
- `is_empty`: whether it has no pages.
- `error`: diagnostic text, if any.

### `PDFInfo.is_valid`

A property returning `True` only when the file is not encrypted, not empty, and has no error.

### `validate_pdf(path)`

Checks whether a file exists and whether `pdfplumber` can open it. It catches parser errors and recognizes messages containing `encrypt` or `password` as encryption-related.

This function prevents an invalid PDF from consuming Gemini API quota.

### `extract_text(path)`

Extracts text from every page and joins non-empty page results with newlines. It is a general full-document extraction helper.

The current application mainly uses `extract_relevant_pages()` instead, because sending fewer relevant pages can reduce token usage.

### `_score_page(text)`

Private helper that:

1. Lowercases page text.
2. Checks every keyword in `SECTION_KEYWORDS`.
3. Adds the matching keyword weight.
4. Returns the page score.

A page can receive multiple weights if it contains several relevant phrases.

### `extract_relevant_pages(path, max_pages=40, min_chars_fallback=5000)`

Reads every page and stores `(page_index, score, text)` tuples.

Behavior:

- If the PDF has at most `max_pages`, returns all text.
- Otherwise, sorts pages by score descending.
- Selects the top `max_pages` pages.
- Restores selected pages to original order.
- If the filtered text is too short, falls back to the full PDF text.

This is a heuristic, not a perfect semantic retrieval system. A page with no matching keyword can be omitted from a large report even when it contains useful context.

### `is_scanned_pdf(text, threshold=2000)`

Returns `True` when stripped text is shorter than the threshold. The assumption is that image-only or scanned reports produce very little extractable text.

---

## 8. `core/extractor.py`: Gemini Integration

### `ExtractionResult`

A dataclass containing:

- `company_name`.
- `data`: parsed JSON dictionary, or `None` on failure.
- `error`: error text, or `None` on success.
- `mode`: `text` or `vision`.

### `ExtractionResult.success`

A property that returns `True` only when `data` exists and `error` is `None`.

### `GeminiExtractor.__init__(api_key, model, prompt_path)`

Creates the Gemini client and loads the system prompt.

It builds a `GenerateContentConfig` with:

- `system_instruction`: contents of `SKILL.md`.
- `temperature=0.0`: more deterministic extraction.
- `response_mime_type="application/json"`: asks for JSON.
- `response_schema=CompanyData`: provides the expected structured schema.

### `_call_text(prompt)`

Private retried Gemini text call. It sends a text prompt to `generate_content()` and returns `response.text`.

The `@retry` decorator uses:

- `stop_after_attempt(5)`.
- `wait_exponential(multiplier=1, min=2, max=30)`.
- Logging before sleeps.
- Re-raising after the final failed attempt.

### `extract_from_text(text, company_name)`

Builds a prompt containing the company name and extracted annual-report text. It calls `_call_text()`, parses the returned JSON, and returns a successful `ExtractionResult`.

It separately catches:

- `json.JSONDecodeError` for invalid JSON.
- Any other exception for API or runtime failures.

### `_call_vision(uploaded_file, prompt)`

Private retried Gemini call for multimodal content. Its `contents` list contains the uploaded PDF object and the prompt.

### `extract_from_pdf(pdf_path, company_name)`

Vision-mode extraction function.

1. Uploads the PDF using `client.files.upload()`.
2. Builds a prompt referring to the attached PDF.
3. Calls `_call_vision()`.
4. Parses JSON.
5. Always attempts to delete the remote uploaded file in `finally`.

The cleanup is important because temporary API files should not accumulate.

---

## 9. `core/schemas.py`: Data Contracts

The schema is the contract between Gemini and the Excel engine.

### `_coerce_numeric(v)`

Normalizes values before Pydantic validation:

- `None` stays `None`.
- Integers and floats pass through.
- Strings have commas and spaces removed.
- Empty strings, dashes, `N/A`, `NA`, `null`, and `None` become `None`.
- Numeric strings become floats.
- Other strings become `None`.

Example:

```python
_coerce_numeric("1,234.56")  # 1234.56
_coerce_numeric("N/A")        # None
```

### `FieldBase`

Generic extracted field with:

- `value`: extracted value.
- `source`: citation such as a page and section.
- `flag`: optional reviewer note.

### `AERevenueSplitValue`

Contains four numeric values:

- `domestic_ae_fy25`.
- `export_ae_fy25`.
- `domestic_third_party_fy25`.
- `export_third_party_fy25`.

Its wildcard validator applies numeric coercion to every field.

### `AERevenueSplit`

Wraps the split values with a source and optional flag.

### `ShareholderRow`

Contains shareholder name and share counts for FY25 and FY24.

Validators:

- Reject empty or whitespace-only names.
- Trim surrounding whitespace from names.
- Coerce share counts through `_coerce_numeric()`.

### `Shareholding`

Contains shareholder rows, year totals, and a source.

### `RPTItem`

A related-party transaction item with:

- Exact transaction label.
- FY25 value.

Its validators reject blank labels and normalize the numeric value.

### `RelatedPartyTransactions`

Contains RPT items and a source.

### `LitigationItem`

Contains:

- Nature of dues.
- Amount demanded in lakhs.
- Amount paid in lakhs.
- Period.
- Forum.

Amounts are normalized numerically.

### `Litigation`

Contains litigation items and a source.

### `ExtractedFields`

The complete list of fields expected for one company, including headquarters, descriptions, financials, auditors, AE revenue, cash, receivables, PE investment, shareholding, RPT, countries, litigation, website, and LinkedIn.

### `CompanyData`

Top-level model:

```python
CompanyData(
    entity="Company name",
    fields=ExtractedFields(...),
)
```

---

## 10. `config/settings.py`: Runtime Configuration

### `Settings`

A frozen dataclass containing application defaults:

- Gemini API key.
- Gemini model.
- Concurrent worker count.
- Vision threshold.
- RPT and litigation row limits.
- Maximum relevant pages.
- System prompt path.

`frozen=True` prevents accidental mutation after creation.

### `load_settings(**overrides)`

Builds `Settings` from environment variables, then applies explicit keyword overrides.

This gives the application three configuration levels:

1. Hard-coded safe defaults.
2. Environment variables.
3. UI or test overrides.

### `get_system_prompt(path="SKILL.md")`

Reads the domain prompt from disk and caches the result with `@lru_cache(maxsize=1)`. The prompt is loaded once per process for a given normal usage path.

---

## 11. `config/template_mapping.py`: Excel Coordinates

### `FactSheetMapping`

A dataclass that centralizes Excel layout knowledge.

Examples:

```python
mapping.hq_india_cell       # "C3"
mapping.turnover_fy25_cell # "D12"
mapping.shareholding_start_row  # 42
```

Important groups:

- Header: title cell.
- Company overview: headquarters and descriptions.
- Standalone summary: turnover, cost, PBT.
- Consolidated summary: D18:F21 cells.
- Auditors.
- AE revenue split.
- Cash and receivables.
- PE investment.
- Shareholding rows and total row.
- RPT capacities and offsets.
- Countries, litigation, website, and LinkedIn.

The mapping exists so a template layout change can usually be handled in one file instead of scattered throughout `excel_engine.py`.

---

## 12. `core/excel_engine.py`: Workbook Population

### `FillResult`

Dataclass reporting the Excel operation:

- Output path.
- Company name.
- Warnings.
- Number of shareholding rows written.
- Number of RPT rows written and dropped.
- Number of litigation rows written and dropped.
- Optional error.

### `FillResult.success`

Returns `True` when `error` is `None`.

### `_note(ws, cell_ref, value, source)`

Writes a value to an Excel cell. When a source is supplied, it attaches an OpenPyXL `Comment` such as:

```text
Source: Annual Report, Page 87
```

This is the main traceability mechanism.

### `_find_total_row(ws, start_data_row, label="Total", max_scan=30)`

Scans column C from the start row until it finds the requested label. It locates the shareholding total row even if the template has been expanded.

### `fill_factsheet(data, template_path, out_path, sheet_name, mapping=None)`

This is the main Excel function.

High-level steps:

1. Create default mapping if none was provided.
2. Load the workbook and worksheet.
3. Rename the worksheet to the first 31 characters of the entity name.
4. Write the title.
5. Write headquarters and descriptions with source comments.
6. Write standalone financial values.
7. Leave the PBT formula in place and add an explanatory comment.
8. Fill consolidated cells with `N/A` and add a source comment.
9. Write statutory auditors.
10. Write AE, domestic, export, and third-party revenue values.
11. Write cash and AE receivables.
12. Write PE investment.
13. Find the shareholding total row.
14. Write shareholder rows and percentage formulas.
15. Write shareholding totals.
16. Write RPT items up to configured capacity.
17. Record warnings when RPT items exceed capacity.
18. Write a cross-check comment for the AE revenue row.
19. Write countries of presence.
20. Write litigation rows up to configured capacity.
21. Record warnings when litigation items exceed capacity.
22. Write website and LinkedIn.
23. Clear and hide unused table rows.
24. Save the workbook and return `FillResult`.

Important behavior:

- Formatting is intended to remain from the template.
- Values are source-cited with comments.
- RPT and litigation rows have capacity limits.
- Excess rows are reported as warnings instead of silently pretending they were written.
- Shareholding expansion is expected to happen before this function when required.

---

## 13. `core/safe_sheet.py`: Safe Excel Row Insertion

Normal `openpyxl` row insertion can leave formulas, merged cells, formatting, or dimensions inconsistent. This module uses a snapshot-and-rebuild algorithm.

### `CELL_REF_RE`

Regular expression that finds A1-style Excel references such as:

- `D12`.
- `$D$44`.
- `AA100`.

### `shift_formula(formula, threshold, delta)`

Shifts row numbers in cell references when the row is greater than or equal to `threshold`.

Example:

```python
shift_formula("=D12-D13", threshold=13, delta=2)
# "=D12-D15"
```

The dollar sign marks an absolute reference in Excel, but this function still shifts the physical row number because the worksheet rows moved.

### `_shift_ref(ref, threshold, delta)`

Applies `shift_formula()` to a merged-cell range string such as `B40:B45`.

### `insert_rows_safe(src_path, out_path, sheet_name, threshold, delta)`

Safely expands a worksheet:

1. Validates threshold and delta.
2. Opens the workbook.
3. Verifies the worksheet exists.
4. Snapshots merged ranges.
5. Unmerges cells.
6. Snapshots row heights.
7. Snapshots cell values and styles.
8. Clears the sheet.
9. Rewrites each cell at its shifted row.
10. Shifts formulas through `shift_formula()`.
11. Copies style from the row above into the new rows.
12. Restores row heights.
13. Recreates shifted merged ranges.
14. Saves the output workbook.

When `delta` is zero, it simply copies the workbook to the output path.

---

## 14. `ui/components.py`: Reusable UI Functions

### `render_metrics_bar(total, success, failed, pending)`

Displays four metric cards in columns:

- Total.
- Succeeded.
- Failed.
- Pending.

It uses Material Symbols labels and CSS classes from `ui/styles.py`.

### `_STATUS_LABELS`

Maps internal status identifiers to visible labels.

### `render_company_card(name, status, mode="", error="")`

Displays company name, extraction mode, status badge, and optional error text.

### `render_preview_editor(data, company_name)`

Creates editable Streamlit data editors for:

- Financial summary.
- Shareholding.
- Related-party transactions.
- Litigation.

It writes edited values back into the original nested `data` dictionary and returns it.

The function uses Pandas data frames because `st.data_editor` works naturally with tabular data.

### `create_zip_download(file_dict)`

Creates an in-memory ZIP file:

1. Creates `io.BytesIO()`.
2. Opens a `zipfile.ZipFile` using `ZIP_DEFLATED` compression.
3. Writes each filename and byte value.
4. Returns ZIP bytes.

No ZIP file has to be written to disk.

### `render_download_section(results, errors)`

Shows individual Excel download buttons. If more than one result exists, it also creates one bulk ZIP button.

The `errors` argument is available to the rendering function, although the current implementation mainly uses it as contextual input while warning when there are no successful files.

---

## 15. `ui/styles.py`: Visual Styling

### `CUSTOM_CSS`

Contains CSS for:

- Metric cards.
- Company cards.
- Status badges.
- Hover behavior.
- Processing badge animation.

Theme colors are primarily defined in `.streamlit/config.toml`.

### `inject_styles()`

Calls `st.markdown(CUSTOM_CSS, unsafe_allow_html=True)` to place the CSS into the page.

The application uses custom HTML only for these compact visual components. Most interaction uses native Streamlit widgets.

---

## 16. `SKILL.md`: The Gemini's Domain Instructions

`SKILL.md` is not a Python module. It is a system prompt and domain policy document.

It tells Gemini:

- Every output is a draft for human review.
- Every filled Excel cell should have a source comment.
- Formatting must not be changed unnecessarily.
- Turnover means Revenue from Operations only.
- Other Income must not be included in turnover.
- Foreign AE service revenue must be identified specifically.
- Director shareholders should be listed individually.
- Individual related parties should be excluded from RPT extraction.
- RPT labels should be collated by transaction nature.
- Litigation must come only from the Annual Report.
- Countries should be represented as country names.
- Amounts should be converted to lakhs.

This is an important separation of concerns:

- Python controls execution, storage, validation, and Excel writing.
- `SKILL.md` controls the language model's extraction instructions.

Changing `SKILL.md` can change the model's behavior without changing Python code.

---

## 17. Complete Data Flow Example

Suppose the user uploads `Example_Annual_Report.pdf`.

### Step 1: Streamlit receives the file

`st.file_uploader` returns an uploaded-file object. `app.py` writes its bytes into a temporary path.

### Step 2: Validate the PDF

```python
info = validate_pdf(pdf_path)
```

If the file is missing, empty, encrypted, or corrupt, the company gets an error and no Gemini quota is used.

### Step 3: Extract relevant pages

```python
text = extract_relevant_pages(pdf_path, max_pages=40)
```

Pages containing financial, RPT, shareholding, litigation, and corporate keywords receive higher scores.

### Step 4: Select extraction mode

```python
if is_scanned_pdf(text, threshold=2000):
    result = extractor.extract_from_pdf(pdf_path, company_name)
else:
    result = extractor.extract_from_text(text, company_name)
```

### Step 5: Gemini returns structured data

The response is expected to contain an entity and all fields from `CompanyData`.

### Step 6: Session state stores the result

The result is stored under the company name so the review page can access it after a rerun.

### Step 7: Human review

The user edits tables. Those edits modify the nested dictionary in session state.

### Step 8: Excel generation

`fill_factsheet()` maps the nested data to workbook cells using `FactSheetMapping`.

### Step 9: Source comments

Every relevant populated value receives a comment containing the extracted source.

### Step 10: Download

The workbook is loaded into bytes and passed to `st.download_button`.

---

## 18. Learning the Code in the Best Order

Study in this order:

1. Read this document's architecture and data-flow sections.
2. Read `app.py` from top to bottom.
3. Read `config/template_mapping.py` to understand the Excel layout.
4. Read `core/schemas.py` to understand the JSON contract.
5. Read `core/pdf_processor.py` to understand input preparation.
6. Read `core/extractor.py` to understand Gemini calls and retries.
7. Read `core/excel_engine.py` to understand output generation.
8. Read `core/safe_sheet.py` to understand row expansion and formula safety.
9. Read `ui/components.py` to understand the review and download experience.
10. Read `SKILL.md` to understand the extraction rules.
11. Read the tests and deliberately break small functions to see how failures appear.

A useful first exercise is to trace one field, such as turnover:

```text
SKILL.md rule
  -> CompanyData field
  -> review data editor
  -> FactSheetMapping.turnover_fy25_cell
  -> Excel cell D12
  -> source comment
```

---

## 19. Testing Strategy

### Schema tests

`tests/test_schemas.py` checks:

- Numeric coercion.
- Empty-name rejection.
- Empty-label rejection.
- Whitespace trimming.
- Nullable values.
- A complete minimal `CompanyData` payload.

### Excel engine tests

`tests/test_excel_engine.py` checks:

- Basic workbook filling.
- Sheet renaming.
- Cell values.
- Source comments.
- RPT capacity warnings.
- Hidden unused rows.
- Custom mappings.
- FillResult metadata.

### Safe-sheet tests

`tests/test_formula_shift.py` checks:

- Simple cell references.
- Ranges.
- Absolute rows and columns.
- Mixed formulas.
- Multi-letter columns.
- Invalid arguments.
- Missing worksheets.
- Formula and style preservation after insertion.

The tests use temporary directories, so they do not modify the user's real templates.

---

## 20. Common Debugging Questions

### Why did a PDF use Vision mode?

The extracted text was shorter than `VISION_FALLBACK_THRESHOLD`, which defaults to 2,000 characters. This usually means the report is scanned or image-heavy.

### Why was a PDF rejected?

Inspect the `PDFInfo.error`, `is_encrypted`, and `is_empty` values from `validate_pdf()`.

### Why is an RPT row missing?

`FactSheetMapping.rpt_default_capacity` limits how many rows are written. `FillResult.rpt_rows_dropped` and `FillResult.warnings` identify overflow.

### Why is the calculated PBT different from the Annual Report?

The template formula is turnover minus total cost, while the project rules exclude Other Income from turnover. The comment in `fill_factsheet()` explicitly explains this expected difference.

### Why did the Excel table expand?

The number of extracted shareholders exceeded `shareholding_default_capacity`. `insert_rows_safe()` was called before `fill_factsheet()`.

### Why did a review edit disappear?

Streamlit reruns the script. The edited value must be written back to `st.session_state`; `render_preview_editor()` does this for its supported tables.

### Why did the model return an error?

Possible causes include an API key problem, rate limiting, a temporary Gemini failure, invalid JSON, or an incomplete response. `GeminiExtractor` retries API calls and returns an `ExtractionResult` with an error when retries are exhausted.

### Why should I verify the generated workbook?

The model can misread tables, page context, units, or company relationships. ALFA adds traceability and review controls, but the human reviewer remains responsible for the final fact sheet.

---

## 21. Important Design Decisions

### Text first, vision second

Native PDF text is cheaper and easier to inspect. Vision is reserved for reports that do not yield enough text.

### Relevant-page filtering

Large Annual Reports can contain hundreds of pages. Keyword scoring reduces the amount of text sent to the model while retaining the most relevant sections.

### Schema-constrained extraction

A strict schema makes downstream Excel writing predictable. It also gives the model a clear target shape.

### Source comments

A value without provenance is difficult to audit. Source comments connect the Excel output back to the Annual Report.

### Mapping instead of scattered coordinates

All template coordinates live in `FactSheetMapping`. This reduces the cost of changing the workbook layout.

### Safe row insertion

Shareholding rows can vary by company. The custom snapshot-and-rebuild algorithm protects formulas, styles, merged ranges, and row heights better than a blind row insertion.

### Human review

The application is designed for assisted automation. It accelerates document review while preserving a deliberate verification step.

---

## 22. Small Practice Tasks

1. Change the default maximum relevant pages in `Settings`.
2. Add one keyword to `SECTION_KEYWORDS` and write a test for `_score_page()`.
3. Add a new optional field to the schema and update the Excel mapping.
4. Add a new review table to `render_preview_editor()`.
5. Add a test proving the new field receives a source comment.
6. Change the model default through `.env` without editing Python.
7. Run `shift_formula()` with a formula containing multiple ranges.
8. Temporarily reduce RPT capacity and observe the warning.
9. Run the app with a text PDF and confirm the extraction mode.
10. Run the app with an image-only PDF and confirm the vision fallback.

After every code change:

```powershell
.\.venv\Scripts\python.exe -m pytest -q
```

---

## 23. Final Summary

The shortest accurate description of ALFA is:

> ALFA is a Streamlit human-in-the-loop pipeline that prepares Transfer Pricing fact sheets by combining PDF extraction, keyword-based page selection, Gemini structured extraction, Pydantic contracts, and source-cited OpenPyXL workbook generation.

The most important files to understand are:

- `app.py`: when and in what order things happen.
- `SKILL.md`: what the AI is instructed to extract.
- `core/schemas.py`: what shape the AI response must have.
- `core/excel_engine.py`: how data becomes a workbook.
- `config/template_mapping.py`: where each value is written.
- `core/safe_sheet.py`: how dynamic Excel rows remain safe.
- `ui/components.py`: how users review and download results.
- `tests/`: what the project considers important behavior.
