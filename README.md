# Excell - Enterprise Excel Translation Automation

This project translates textual content in `.xlsx` files while preserving workbook structure, formulas, merged cells, and formatting.

## Features
- Streamlit UI for multiple file upload (`.xlsx`) and ZIP upload.
- Deterministic engine routing:
  - `azure`: Azure Translator first, automatic fallback to local Ollama Gemma.
  - `local`: local Ollama Gemma only.
- Openpyxl-only workbook processing (no pandas, no CSV conversions).
- Translation of:
  - worksheet names,
  - string cell values (excluding formulas),
  - comments/notes,
  - drawing/chart XML text nodes under `xl/drawings/*.xml` and `xl/charts/*.xml`.
- Per-item logs with file, sheet, object id, original/translated text, engine and errors.

## Setup
```bash
python -m venv .venv
source .venv/bin/activate
pip install -e .[test]
```

## Environment variables
For Azure:
- `AZURE_TRANSLATOR_ENDPOINT`
- `AZURE_TRANSLATOR_KEY`
- `AZURE_TRANSLATOR_REGION`

For Ollama:
- `OLLAMA_ENDPOINT` (default: `http://localhost:11434/api/generate`)
- `OLLAMA_MODEL` (default: `gemma:2b`)

## Run UI
```bash
streamlit run app.py
```

## Tests
```bash
pytest -q
```

## Test assets and validation
Generate sample workbook:
```bash
python scripts/generate_test_assets.py
```

Validate original vs translated workbook:
```bash
python scripts/validate_translation.py tests/assets/sample_input.xlsx /path/to/translated.xlsx
```

Basic translator setup check (Azure + Ollama using "hello"):
```bash
python test_setup.py
```
