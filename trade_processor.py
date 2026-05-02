"""
Trade Processor Module - Processes trade data from TBLT sheet and generates confirmations.
"""

from typing import Dict, List, Any
import logging
from pathlib import Path


class TradeProcessor:
    """
    Processes trade data from TBLT sheet and generates confirmations.
    """

    TEMPLATE_SHEET = "Template"
    DATA_SHEET = "TBLT"
    PRINT_RANGE = "B1:I46"
    KEY_COLUMN = 14  # Column N (1-indexed)
    DEFAULT_DATA_ROW = 2  # Row that Template initially references

    def __init__(self, excel_app, output_folder: str):
        self.logger = logging.getLogger(__name__)
        self.excel = excel_app
        self.wb = excel_app.wb
        self.app = excel_app.app
        self.output_folder = Path(output_folder)
        self.output_folder.mkdir(exist_ok=True)

        # Cache sheets
        self.template_sheet = self.wb.sheets[self.TEMPLATE_SHEET]
        self.tblt_sheet = self.wb.sheets[self.DATA_SHEET]

    def process_all_trades(self) -> Dict[str, Any]:
        """
        Main processing loop: iterate through TBLT rows and generate confirmations.
        """
        # Find last row with data in column N
        last_row = self._find_last_data_row()

        self.logger.info(f"Found {last_row - 1} potential trade rows")

        # Initialize summary
        summary = {
            'total_scanned': 0,
            'processed': 0,
            'successes': 0,
            'errors': 0,
            'skipped': 0,
            'pdfs_created': [],
            'errors_detail': []
        }

        # Process each row
        for row_idx in range(2, last_row + 1):
            summary['total_scanned'] += 1

            result = self._process_single_trade(row_idx)

            if result['status'] == 'success':
                summary['successes'] += 1
                summary['processed'] += 1
                summary['pdfs_created'].append(result['pdf_path'])
            elif result['status'] == 'skipped':
                summary['skipped'] += 1
            else:  # error
                summary['errors'] += 1
                summary['errors_detail'].append(result['error'])

        return summary

    def _find_last_data_row(self) -> int:
        """Find the last row with non-empty column N"""
        # Use Excel's end method for efficiency
        last_cell = self.tblt_sheet.range(f"N{self.tblt_sheet.cells.last_cell.row}").end('up')
        return last_cell.row if last_cell.row > 1 else 1

    def _process_single_trade(self, row_idx: int) -> Dict[str, Any]:
        """
        Process a single trade row:
        1. Check if column N has a value (trade ID)
        2. Update Template to use that row's data
        3. Recalculate
        4. Copy to new sheet as values
        5. Export to PDF
        """
        try:
            # Get trade ID from column N
            trade_id = self.tblt_sheet.range(f"N{row_idx}").value

            if not trade_id:
                self.logger.debug(f"Row {row_idx}: Column N empty, skipping")
                return {'status': 'skipped', 'reason': 'empty_column_n'}

            # Convert to string for filename
            trade_id_str = str(trade_id).strip()
            self.logger.info(f"Processing trade: {trade_id_str} (row {row_idx})")

            # Step 0: Restore Template to default row (2) first for clean state
            self.excel.restore_default_references(self.TEMPLATE_SHEET, self.DEFAULT_DATA_ROW)

            # Step 1: Configure Template to reference this row
            self._update_template_reference(row_idx)

            # Step 2: Recalculate to get fresh values
            self.app.calculate()

            # Step 3: Check for Excel errors in template
            if self._has_template_errors():
                raise ValueError("Template contains calculation errors")

            # Step 4: Create confirmation sheet with values + formats
            sheet_name = f"Confo_{trade_id_str}"
            self._create_confirmation_sheet(sheet_name)

            # Step 5: Export to PDF
            pdf_path = self.output_folder / f"{trade_id_str} Confo.pdf"
            success = self.excel.export_to_pdf(
                self.wb.sheets[sheet_name],
                str(pdf_path)
            )

            if not success:
                raise RuntimeError("PDF export failed")

            return {
                'status': 'success',
                'trade_id': trade_id_str,
                'pdf_path': str(pdf_path)
            }

        except Exception as e:
            self.logger.error(f"Error processing row {row_idx}: {e}")
            return {
                'status': 'error',
                'row': row_idx,
                'error': f"Row {row_idx}: {str(e)}"
            }

    def _update_template_reference(self, data_row: int):
        """
        Update Template sheet to reference the correct TBLT row.

        The Template sheet has formulas like =TBLT!A2, =TBLT!B2, etc.
        We need to change them to reference row data_row.
        """
        # Use the excel automation method to update formulas
        self.excel.update_formula_references(
            sheet_name=self.TEMPLATE_SHEET,
            old_row=self.DEFAULT_DATA_ROW,
            new_row=data_row
        )

        self.logger.debug(f"Updated Template references to row {data_row}")

    def _has_template_errors(self) -> bool:
        """Check if any cells in Template contain Excel errors"""
        for cell in self.template_sheet.range(self.PRINT_RANGE):
            if cell.value and isinstance(cell.value, str):
                if cell.value.startswith('#'):  # Error values start with #
                    self.logger.warning(f"Found error in Template: {cell.value}")
                    return True
        return False

    def _create_confirmation_sheet(self, sheet_name: str):
        """Create a new sheet with formulas and formatting"""
        # Copy as formulas + formats (formulas will reference TBLT data)
        self.excel.copy_sheet_formulas_and_formats(
            source_sheet_name=self.TEMPLATE_SHEET,
            source_range=self.PRINT_RANGE,
            dest_sheet_name=sheet_name
        )
