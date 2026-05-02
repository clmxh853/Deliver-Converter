# Trade Confirmation Generator — PRD

## 1. Objective

Automatically generate one trade confirmation per trade in the `TBLT` sheet by using the `Template` sheet, materializing each confirmation as a new Excel sheet and exporting it to PDF.

---

## 2. Inputs

### Workbook
- **File**: `Confo_Template_test.xlsx` (or whichever file is configured in `Confo_Generator.py`)
- **Required sheets**: `TBLT` (data source) and `Template` (confirmation layout)

### TBLT Sheet (Data Source)
- Header is in **row 1**
- Data begins at **row 2**
- A row is considered present if **column N is non-empty** (column N holds the trade ID)
- Rows with empty column N are skipped

### Template Sheet (Confirmation Layout)
- Print area: **B1:I46**
- Contains Excel formulas that reference `TBLT` data — e.g., `=TBLT!A2`, `=TBLT!B2`, etc.
- These formulas are dynamically updated to point to the current trade's row
- The formula structure must never be changed — only the row reference is updated per trade
- Initially references row 2 (as an example)

---

## 3. Outputs

For each valid trade row (row 2 to last non-empty row in column N):

1. **Confirmation sheet** — new sheet created in the workbook, named `Confo_<column-N-value>` (e.g., `Confo_3471704`), containing the rendered confirmation copied from `Template!B1:I46` with formulas and formatting intact
2. **PDF file** — exported to `pdf_output/<column-N-value> Confo.pdf` (e.g., `pdf_output/3471704 Confo.pdf`)
3. **Output workbook** — saved with `_output` suffix (e.g., `Confo_Template_test_output.xlsx`), containing all confirmation sheets
4. **Error log** — `error log.txt` in the project root, containing:
   - Total rows scanned
   - Confirmations created
   - PDFs written
   - Skipped rows
   - Error details

---

## 4. Core Functional Flow

### Step 1 — Open Workbook
Open `Confo_Template_test.xlsx` using xlwings (Excel COM automation).

### Step 2 — Determine Last Row
Scan column N of the `TBLT` sheet to find the last non-empty row. All rows from 2 to this row will be processed.

### Step 3 — Per-Trade Processing (loop from row 2 to lastRow)

For each row:

1. **Read trade ID** from column N. If empty, skip the row.
2. **Restore Template** — reset all `TBLT!` formula references back to row 2 (clean slate).
3. **Update Template** — change all `TBLT!` formula references to point to the current trade row (e.g., row 2 → row 5).
4. **Recalculate** — force Excel to recalculate the entire workbook so Template renders with the new row's data.
5. **Check for errors** — if any cell in Template's print area starts with `#`, log the error and skip this row.
6. **Create confirmation sheet** — copy `Template!B1:I46` (formulas + formats) into a new sheet named `Confo_<trade-id>`.
7. **Auto-fit columns** — auto-fit column D and column I so long strings display fully.
8. **Export to PDF** — set print area to `B1:I46`, fit to page width, export as PDF to `pdf_output/<trade-id> Confo.pdf`.
9. **Record result** — log success or failure, continue to next row.

### Step 4 — Save & Report
After the loop completes:
- Save the workbook with `_output` suffix
- Print/log a summary: total scanned, processed, successes, skipped, errors

---

## 5. Non-Functional Requirements

| Requirement | Description |
|-------------|-------------|
| No Template alteration | Do not change the Template sheet's formula structure — only update row references temporarily |
| Preserve original | Default behavior saves to a new file with `_output` suffix; original is never overwritten |
| Confirmation sheet content | Must contain rendered values and formatting — formulas must not be stripped |
| PDF naming | Must follow pattern `<column-N-value> Confo.pdf` exactly (e.g., `3471704 Confo.pdf`) |
| PDF scaling | Must use fit-to-width so all columns (especially D and I) are visible in the output PDF |

---

## 6. Error Handling

| Scenario | Behavior |
|----------|----------|
| Column N is empty | Skip row, log as skipped |
| Template has `#` error cells after recalculation | Skip row, log error detail |
| PDF export or file write fails | Log error, continue processing remaining rows |
| Excel COM failure | Log fatal error and exit |

---

## 7. Architecture

The project has three Python modules in a layered architecture:

```
Confo_Generator.py          # Entry point — orchestrates everything
    ├── sets up logging
    ├── opens workbook via ExcelAutomation
    └── calls TradeProcessor.process_all_trades()

excel_automation.py         # Excel COM wrapper (xlwings)
    ├── ExcelAutomation     — context manager: opens/saves/closes workbook
    ├── update_formula_references()     — updates TBLT row refs in formulas
    ├── restore_default_references()    — resets formulas to row 2
    ├── copy_sheet_formulas_and_formats() — copies B1:I46 as formulas+formats
    ├── auto_fit_columns()             — auto-fits specified columns
    └── export_to_pdf()               — exports sheet to PDF

trade_processor.py          # Business logic
    ├── TradeProcessor      — holds TBLT and Template sheet references
    ├── process_all_trades()         — main loop
    ├── _find_last_data_row()        — scans column N for last row
    ├── _process_single_trade()      — per-trade workflow
    ├── _update_template_reference() — delegates to ExcelAutomation
    ├── _has_template_errors()       — checks for # error cells
    └── _create_confirmation_sheet() — delegates to ExcelAutomation
```

---

## 8. Key Implementation Notes

- **Excel COM automation** via `xlwings` — requires Microsoft Excel to be installed on the machine
- **Formulas vs values** in xlwings: use `.formula` (not `.value`) to read/set formulas; `.value` returns calculated results
- **Cell addresses** in xlwings: use `.address` (not `.coordinate`)
- **Regex group references**: when replacing row numbers in formulas, use `\g<1>\g<2>` syntax to avoid "invalid group reference" errors with numeric group names
- **TBLT sheet**: header in row 1, data from row 2, column N (14th column) determines validity
- **Template sheet**: print area B1:I46, all formulas reference `TBLT!` with row 2 as default
- **PDF output folder**: `pdf_output/` in the project root
- **Output workbook suffix**: `_output` (e.g., `Confo_Template_test_output.xlsx`)

---

## 9. Dependencies

```
xlwings>=0.30.0
pywin32>=300
```

Microsoft Excel must be installed (Windows only — uses COM automation).

---

## 10. Running the Script

```bash
pip install -r requirements.txt
python Confo_Generator.py
```

After a successful run:
- PDFs are in `pdf_output/`
- Output workbook is saved with `_output` suffix
- Summary is printed to console and logged