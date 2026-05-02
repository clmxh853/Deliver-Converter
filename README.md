# Trade Confirmation Generator

Generates trade confirmations from an Excel template. Reads trade data from a `TBLT` sheet, uses a `Template` sheet to render confirmations, and exports each as a new Excel sheet and PDF file.

## File Purposes & Relationships

```
INPUT: Confo_Template_test.xlsx
  ├── "TBLT" sheet     → source data (one row per trade)
  └── "Template" sheet → format template with formulas pointing to TBLT row 2
```

### The 3 Core Python Files

**Confo_Generator.py** — The entry point / orchestrator
- It doesn't do any Excel work directly — it wires everything together
- Sets up logging to a timestamped `.log` file
- Opens the workbook via `ExcelAutomation` context manager
- Creates a `TradeProcessor` and calls `process_all_trades()`
- Prints a summary at the end (total scanned, successes, PDFs, errors)

```
Confo_Generator.py
       │
       ├── excel_automation.py  (ExcelAutomation context manager)
       │              └── talks to Excel via xlwings COM
       │
       └── trade_processor.py   (TradeProcessor)
                      └── uses ExcelAutomation to do the actual work
```

**excel_automation.py** — Thin wrapper around Excel COM
- `ExcelAutomation` context manager: opens Excel on `__enter__`, saves and closes on `__exit__`
- `update_formula_references()` — finds `=TBLT!A2` style formulas and changes the row number to the target trade row
- `restore_default_references()` — resets all TBLT formulas back to row 2 (clean slate before each trade)
- `copy_sheet_formulas_and_formats()` — copies `B1:I46` from Template into a new sheet as formulas + formats (not static values)
- `export_to_pdf()` — uses Excel's native `ExportAsFixedFormat` API to export a sheet as PDF

**trade_processor.py** — Business logic / processing pipeline
- `TradeProcessor` holds references to both the `TBLT` sheet and `Template` sheet
- `process_all_trades()` — loops row 2 → last non-empty row in column N, builds the summary dict
- `_process_single_trade(row_idx)` — the per-trade workflow:
  1. Read trade ID from column N of that row
  2. Restore Template formulas to default row 2
  3. Update Template formulas to point to the current trade row
  4. Recalculate
  5. Check for `#` error cells — skip if found
  6. Copy Template `B1:I46` to a new sheet named `Confo_<trade_id>`
  7. Export that new sheet to `pdf_output/<trade_id> Confo.pdf`

### The Flow, Step by Step

```
User runs: python Confo_Generator.py

  1. Confo_Generator.setup_logging()
     → creates confo_generator_YYYYMMDD_HHMMSS.log

  2. Confo_Generator.main()
     → ExcelAutomation opens the .xlsx file (xlwings COM)
        └─ ExcelAutomation.__enter__() launches Excel, opens workbook

  3. TradeProcessor.process_all_trades() starts looping TBLT rows

     For each row:
       a. TradeProcessor reads col N → trade ID
       b. ExcelAutomation.restore_default_references() → resets Template to row 2
       c. ExcelAutomation.update_formula_references() → changes row 2 → current row
       d. Recalculate workbook
       e. TradeProcessor._has_template_errors() → skip if # errors found
       f. ExcelAutomation.copy_sheet_formulas_and_formats() → copies B1:I46 to new sheet
       g. ExcelAutomation.export_to_pdf() → saves <trade_id> Confo.pdf to pdf_output/

  4. ExcelAutomation.__exit__() → saves workbook as Confo_Template_test_output.xlsx, closes Excel

  5. Confo_Generator.print_summary() → prints totals to console
```

### Supporting Files

| File | Purpose |
|------|---------|
| `inspect_excel.py` | Debug utility — inspects the Excel file structure (sheets, ranges, formulas) |
| `test_imports.py` | Syntax checker — uses `ast.parse()` to validate all `.py` files |
| `requirements.txt` | Dependencies: `xlwings>=0.30.0`, `pywin32>=300` |
| `Confo_Template_test.xlsx` | Input template workbook |
| `pdf_output/` | Folder where generated PDFs are written |
