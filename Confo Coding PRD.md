# Confo Coding PRD

## Objective

Automatically generate one trade confirmation per trade in `TBLT` sheet by using the existing `Template` sheet, materializing each confirmation as a new Excel sheet and exporting it to PDF.

## Inputs

- Workbook: `/mnt/data/Confo_Template_test.xlsx` (uploaded).
- The excel named "Confo_Template_test.xlsx" in the project contains 4 initial sheets
- Template sheet: `Template`. Print area (confirmation area) = **B1:I46**.
  - `Template` sheet has built-in formula that will be recursively used for each trade. The formula structure should never be changed, but the sourcing data shall be changed subject to the trade.
  - each trade info are described in Data source sheet: `TBLT` from row 2 to the last non-empty row.
  - template  sheet initially uses row 2 in `TBLT` data source sheet as an example.

- Data source sheet: `TBLT` — header in row 1, data begins row 2. A row is considered present if **column N** is non-empty.

## Output

For each data row in `TBLT` sheet (row 2 to last non-empty row by column N), step by step:

1. Create a new worksheet in workbook named "Confo_Template_test.xlsx" , containing the rendered confirmation (formulas+ formats) copied from `Template` print area B1:I46.
2. Export that new worksheet as a PDF named: `"<value-in-column-N> Confo.pdf"` (example: `3471704 Confo.pdf`), the column-N here refers to column N in the TBLT sheet
3. Save workbook result (recommended: new file with `_output` suffix) and produce a processing log summary (rows processed, successes, failures, file paths).

## Core functional flow

1. Open workbook.

2. Determine `lastRow` by scanning sheet `TBLT` column N for the last non-empty cell.

3. For each 

   ```
   row
   ```

    from 2 to 

   ```
   lastRow
   ```

   :

   - Feed the `Template` with values for that `row` so `Template` recalculates to show the confirmation for that trade. (Implementation may: temporarily copy the single `row` into the input area `Template` expects, or directly set the input cells Template references.)
   - Force a full workbook recalculation.
   - Copy `Template!B1:I46` as values + formats into a new sheet (name can be `Confo_<rowIndex>` or `Confo_<colNvalue>`).
   - Export that new sheet to PDF named `"<colN value> Confo.pdf"`.
   - Record success/failure and any error message.

4. After loop completes, save the workbook (preferably with `_output`) and return a processing summary.

## Non-functional requirements & constraints

- Do not alter `Template` formulas permanently — only set inputs so the Template recalculates.
- Prefer to preserve original workbook: default behavior should save outputs to a copy (e.g., `Confo_Template_test_output.xlsx`) unless you explicitly instruct the script to overwrite.
- Generated confirmation sheets must contain **rendered values and formatting** (not formulas).
- PDF filenames must exactly follow the pattern: `<column N value> Confo.pdf`.

## Error handling & logging

- If a row has an Excel error in `Template` after recalculation, log the row and error, skip PDF export for that row, continue.
- If column N is empty for a row, treat it as not present (skip).
- If a file write/PDF export fails, log error and continue processing remaining rows.
- Save final summary in format of txt file and  named "error log.txt" , with content: total rows scanned, confirmations created, PDFs written, skipped rows, and error details.

## Acceptance criteria

- For N trade rows with non-empty column N, the script produces N new sheets (or N properly named sheets) and N PDF files in the workbook/output folder, each PDF named `"<colN value> Confo.pdf"`.
- Each exported PDF visually matches `Template!B1:I46` after recalculation for the corresponding trade.
- Script saves an output workbook and returns a short processing log.

## Implementation notes (high level)

- Preferred execution environments that guarantee Excel recalculation: **VBA** (embedded macro) or **Python with Excel COM (e.g., xlwings / pywin32)**. If `Template` uses Excel-only features (volatile functions, macros, external links), choose a COM approach or VBA.
- Ensure proper handling of print area and page setup so PDF matches the look of `Template` B1:I46.
- Keep the PDF export destination path clear (e.g., same folder as workbook or a specified output folder).





> Done! I've created the skill documentation at [CONFO_GENERATOR_SKILL.md](vscode-webview://02d9jb1eisfompca89v4pfg00va2l3f4mtcm0ouno6ip1hkl2a64/CONFO_GENERATOR_SKILL.md) in the project folder.
>
> **Project Summary:**
>
> | File                                                         | Description                    |
> | ------------------------------------------------------------ | ------------------------------ |
> | [Confo_Generator.py](vscode-webview://02d9jb1eisfompca89v4pfg00va2l3f4mtcm0ouno6ip1hkl2a64/Confo_Generator.py) | Main entry point               |
> | [excel_automation.py](vscode-webview://02d9jb1eisfompca89v4pfg00va2l3f4mtcm0ouno6ip1hkl2a64/excel_automation.py) | Excel COM automation (xlwings) |
> | [trade_processor.py](vscode-webview://02d9jb1eisfompca89v4pfg00va2l3f4mtcm0ouno6ip1hkl2a64/trade_processor.py) | Trade processing logic         |
> | [requirements.txt](vscode-webview://02d9jb1eisfompca89v4pfg00va2l3f4mtcm0ouno6ip1hkl2a64/requirements.txt) | Dependencies                   |
> | [CONFO_GENERATOR_SKILL.md](vscode-webview://02d9jb1eisfompca89v4pfg00va2l3f4mtcm0ouno6ip1hkl2a64/CONFO_GENERATOR_SKILL.md) | Skill documentation            |
>
> **To run:** `python Confo_Generator.py`
>
> The generator processes trades from TBLT sheet, updates Template formulas dynamically, and exports each confirmation as both an Excel sheet and PDF.