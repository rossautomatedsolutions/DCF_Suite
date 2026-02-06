# dcf-template-factory

Deterministic DCF spreadsheet generator using Python 3.11+ and `openpyxl` only.

## Features
- Generates a fixed 10-year single-scenario DCF model.
- All assumptions live on the **Inputs** sheet and are exposed as named ranges.
- Formulas avoid numeric literals by referencing named ranges.
- No macros and no pandas.

## Repo structure
```
dcf_factory/
  __init__.py
  __main__.py
  cli.py
  build_dcf.py
  formatting.py
  named_ranges.py
outputs/  (gitignored)
README.md
requirements.txt
```

## Usage

```bash
python -m dcf_factory build --out outputs/DCF_SingleScenario.xlsx
```

The workbook includes these tabs: **Inputs**, **Operating_Model**, **Valuation**, **Outputs**, **Notes**.

## Development

Install dependencies:

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```
