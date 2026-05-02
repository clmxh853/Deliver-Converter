# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Python-based Excel automation project that generates trade confirmations from an Excel template. It reads trade data from a `TBLT` sheet, uses a `Template` sheet to render confirmations, and exports each as a new Excel sheet and PDF file.

## Commands

```bash
# Install dependencies
pip install -r requirements.txt

# Run the main script
python Confo_Generator.py

# Syntax check all Python modules
python test_imports.py
```

## Architecture

The project has a layered architecture:

1. **Confo_Generator.py** — Entry point. Sets up logging, initializes Excel automation, creates TradeProcessor, and runs `process_all_trades()`.

2. **excel_automation.py** — `ExcelAutomation` context manager wrapping Excel COM lifecycle via xlwings.
   - `update_formula_references()` — Updates TBLT row references in formulas
   - `restore_default_references()` — Resets formulas to default row
   - `copy_sheet_formulas_and_formats()` — Copies range with formulas (not values)
   - `export_to_pdf()` — Exports sheet to PDF

3. **trade_processor.py** — `TradeProcessor` class handling trade processing.
   - `_find_last_data_row()` — Finds last non-empty row in column N
   - `_process_single_trade()` — For each trade: updates formulas, creates sheet, exports PDF

## Key Implementation Notes

- Uses **xlwings** for Excel COM automation — requires Microsoft Excel installed
- Use `.formula` (not `.value`) to get/set formulas in xlwings
- Use `.address` (not `.coordinate`) for cell addresses
- Regex replacement uses `\g<1>\g<2>` syntax to avoid "invalid group reference" errors with digits
- `TBLT` sheet: Header in row 1, data starts row 2, column N determines valid rows (non-empty = trade exists)
- `Template` sheet: Print area B1:I46, contains formulas referencing TBLT (e.g., `=TBLT!A2`)
- Output PDF naming: `<column-N-value> Confo.pdf` (e.g., `3471704 Confo.pdf`)
- Output workbook: saved with `_output` suffix
- Error log: saved to `error log.txt` with total rows scanned, confirmations created, PDFs written, skipped rows, and error details

## Dependencies

```
xlwings>=0.30.0
pywin32>=300
```
