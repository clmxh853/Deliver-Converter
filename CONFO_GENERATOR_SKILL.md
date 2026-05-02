# Trade Confirmation Generator Skill

## Overview

This skill generates trade confirmations from an Excel template. It reads trade data from a `TBLT` sheet, uses a `Template` sheet to render confirmations, and exports each as both a new Excel sheet and a PDF file.

## Project Structure

```
Confo_Template_test.xlsx   # Input template with Template + TBLT sheets
Confo_Generator.py             # Main entry point
excel_automation.py           # Excel COM automation via xlwings
trade_processor.py            # Trade processing logic
requirements.txt              # Dependencies (xlwings, pywin32)
pdf_output/                   # Generated PDF files
Confo_Template_output.xlsx  # Output workbook with confirmation sheets
```

## How It Works

1. **Open template workbook** using xlwings (Excel COM automation)
2. **Scan TBLT sheet** for rows with non-empty column N (trade IDs)
3. **For each trade row**:
   - Restore Template formulas to default row (2)
   - Update formulas to reference current trade row (e.g., row 3, 4, 5...)
   - Recalculate workbook
   - Copy Template B1:I46 (formulas + formats) to new sheet
   - Export new sheet to PDF named `<trade-id> Confo.pdf`
4. **Save output workbook** with `_output` suffix containing all confirmation sheets

## Running the Script

```bash
pip install -r requirements.txt
python Confo_Generator.py
```

## Key Components

### Confo_Generator.py
- Main entry point
- Sets up logging
- Initializes Excel automation
- Processes all trades and prints summary

### excel_automation.py
- `ExcelAutomation` class - Context manager for Excel COM lifecycle
- `update_formula_references()` - Updates TBLT row references in formulas
- `restore_default_references()` - Resets formulas to default row
- `copy_sheet_formulas_and_formats()` - Copies range with formulas (not values)
- `export_to_pdf()` - Exports sheet to PDF using Excel API

### trade_processor.py
- `TradeProcessor` class - Processes trades from TBLT
- `_find_last_data_row()` - Finds last non-empty row in column N
- `_process_single_trade()` - Processes one trade: updates formulas, creates sheet, exports PDF

## Important Implementation Notes

- Uses **xlwings** for Excel COM automation (requires Microsoft Excel installed)
- In xlwings, use `.formula` to get/set formulas, not `.value` (which returns calculated results)
- Use `.address` (not `.coordinate`) for cell address
- Regex replacement uses `\g<1>\g<2>` syntax to avoid "invalid group reference" errors with digits

## Dependencies

```
xlwings>=0.30.0
pywin32>=300
```

## Input Format

- **TBLT sheet**: Header in row 1, data starts row 2, column N determines valid rows (non-empty = trade exists)
- **Template sheet**: Print area B1:I46, contains formulas referencing TBLT (e.g., `=TBLT!A2`)

## Output

- **PDF files**: `<column-N-value> Confo.pdf` (e.g., `3471704 Confo.pdf`)
- **Excel workbook**: Original filename with `_output` suffix, contains all confirmation sheets
- **Error log**: `error log.txt` in project root — total rows scanned, confirmations created, PDFs written, skipped rows, and error details

## Error Handling

- If a row has an Excel error in `Template` after recalculation, log the row and error, skip PDF export, continue.
- If column N is empty for a row, treat it as not present (skip).
- If a file write/PDF export fails, log error and continue processing remaining rows.
- Final summary saved to **`error log.txt`** with: total rows scanned, confirmations created, PDFs written, skipped rows, and error details.
