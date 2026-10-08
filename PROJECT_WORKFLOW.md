# ALFA Project Work Workflow

This document describes the standard workflow for developing, testing, and reviewing ALFA changes.

## 1. Start a work session

From the repository root:

```powershell
.\.venv\Scripts\Activate.ps1
```

If the virtual environment does not exist yet:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Confirm that configuration is available before running extraction features. Copy `.env.example` to `.env` when needed and set `GEMINI_API_KEY` or `GOOGLE_API_KEY`.

## 2. Understand the change surface

Use the smallest owning layer for each change:

| Area | Primary location | Typical changes |
| --- | --- | --- |
| Streamlit workflow and session state | `app.py` | Upload flow, review flow, downloads |
| PDF parsing and page selection | `core/pdf_processor.py` | Validation, text extraction, relevance filtering |
| Gemini extraction and retries | `core/extractor.py` | Prompts, model calls, retry behavior |
| Validated extraction data | `core/schemas.py` | Fields, normalization, validation rules |
| Excel generation | `core/excel_engine.py` | Cell writes, comments, workbook output |
| Safe row and formula operations | `core/safe_sheet.py` | Row insertion and formula shifting |
| Excel coordinates and settings | `config/` | Template mappings and environment settings |
| Reusable interface pieces | `ui/` | Components and visual styling |
| Domain extraction rules | `SKILL.md` | Transfer-pricing instructions sent to Gemini |

Before editing, trace the existing call path from the UI or test to the function that directly decides the behavior. Keep the change within that boundary unless the contract genuinely requires another layer.

## 3. Implement the change

1. Create a focused branch or work item according to the team workflow.
2. Read the nearby implementation and its tests before changing behavior.
3. Preserve existing public interfaces and workbook formatting.
4. Keep extracted values source-cited and suitable for human review.
5. Do not place API keys or other secrets in source files, tests, committed `.env` files, or generated artifacts.
6. Add or update a focused test for behavior that can be tested without the Streamlit UI or Gemini service.

For Excel changes, verify that only intended cell values and comments change. For extraction changes, keep the structured response compatible with the Pydantic schema and the review editor.

## 4. Validate locally

Run the unit test suite from the repository root:

```powershell
python -m pytest tests -v
```

For a fast focused check, run the relevant file:

```powershell
python -m pytest tests/test_schemas.py -v
python -m pytest tests/test_excel_engine.py tests/test_formula_shift.py -v
```

Start the application for UI changes:

```powershell
streamlit run app.py
```

Open the local URL printed by Streamlit, normally `http://localhost:8501`, and verify the complete path:

1. Upload a valid BD template.
2. Upload a representative annual report PDF.
3. Run extraction and inspect the review values.
4. Edit at least one value in the review step.
5. Generate and download the workbook.
6. Open the workbook and check values, comments, formulas, and formatting.

When the default port is busy, use another port:

```powershell
streamlit run app.py --server.port 8502
```

## 5. Review the diff

Before handing off the work:

```powershell
git status --short
git diff --check
git diff
```

Check that:

- Only files related to the task changed.
- No credentials, temporary PDFs, generated workbooks, or debug output were added.
- New behavior has a regression test where practical.
- Error messages remain actionable for users.
- Domain rules in `SKILL.md` are still respected.
- Generated Excel files preserve the template structure and formatting.

## 6. Commit and hand off

Use a concise commit message that describes the behavior change. In the handoff, include:

- What changed and why.
- Tests and commands run.
- Whether the Streamlit UI was manually checked.
- Any required environment variables or known limitations.

Do not treat generated extraction output as final professional advice. ALFA output remains a draft that requires human review.

## Troubleshooting

### Streamlit exits immediately

Activate the project virtual environment, confirm dependencies are installed, and run:

```powershell
python -m streamlit run app.py
```

This uses the active Python interpreter explicitly and usually exposes missing-package or configuration errors directly in the terminal.

### Tests cannot import project modules

Run tests from the repository root with the virtual environment activated:

```powershell
python -m pytest tests -v
```

### Extraction is unavailable

Check that the API key is present in `.env` or the process environment, that the configured model is available, and that the uploaded PDF contains the expected annual-report content.