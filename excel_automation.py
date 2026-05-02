"""
Excel Automation Module using xlwings for COM automation.
Handles Excel workbook operations, formula recalculation, and PDF export.
"""

import xlwings as xw
from pathlib import Path
from typing import Optional
import logging
import re


class ExcelAutomation:
    """
    Manages Excel COM automation using xlwings.
    Ensures proper recalculation and workbook handling.
    """

    def __init__(self, workbook_path: str, visible: bool = False):
        self.logger = logging.getLogger(__name__)
        self.workbook_path = Path(workbook_path)
        self.visible = visible
        self.app: Optional[xw.App] = None
        self.wb: Optional[xw.Book] = None

    def __enter__(self):
        """Context manager entry - opens Excel"""
        self.logger.info(f"Opening workbook: {self.workbook_path}")

        # Launch Excel application
        self.app = xw.App(visible=self.visible, add_book=False)
        self.app.display_alerts = False  # Suppress prompts
        self.app.screen_updating = False  # Performance

        # Open workbook
        self.wb = self.app.books.open(str(self.workbook_path.absolute()))

        # Force initial calculation
        self.app.calculation = 'automatic'
        self.wb.calculation = 'automatic'
        self.app.calculate()

        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit - saves and closes"""
        if self.wb:
            # Save output workbook
            output_path = self._get_output_path()
            self.logger.info(f"Saving output workbook: {output_path}")
            self.wb.save(str(output_path))
            self.wb.close()

        if self.app:
            self.app.quit()

        return False  # Don't suppress exceptions

    def _get_output_path(self) -> Path:
        """Generate output workbook path with _output suffix"""
        stem = self.workbook_path.stem
        suffix = self.workbook_path.suffix
        parent = self.workbook_path.parent
        return parent / f"{stem}_output{suffix}"

    def get_sheet(self, name: str) -> xw.Sheet:
        """Get sheet by name"""
        return self.wb.sheets[name]

    def recalculate_workbook(self):
        """Force full workbook recalculation"""
        self.app.calculate()
        self.logger.debug("Workbook recalculated")

    def auto_fit_columns(self, sheet: xw.Sheet, columns: list):
        """
        Auto-fit the width of specified columns in a sheet.
        Columns should be provided as letter strings, e.g. ['D', 'I'].
        """
        for col in columns:
            try:
                sheet.range(f'{col}:{col}').columns.autofit()
                self.logger.debug(f"Auto-fitted column {col}")
            except Exception as e:
                self.logger.warning(f"Auto-fit failed for column {col}: {e}")

    def copy_sheet_formulas_and_formats(self, source_sheet_name: str,
                                         source_range: str,
                                         dest_sheet_name: str) -> xw.Sheet:
        """
        Copy range as FORMULAS + FORMATS to new sheet.
        The formulas will reference the TBLT data and recalculate properly.
        """
        source = self.wb.sheets[source_sheet_name].range(source_range)

        # Create new sheet
        sheet_names = [s.name for s in self.wb.sheets]
        if dest_sheet_name in sheet_names:
            dest = self.wb.sheets[dest_sheet_name]
            dest.delete()

        dest = self.wb.sheets.add(name=dest_sheet_name)

        # Copy the range with formulas and formats
        source.copy(destination=dest.range('B1'))

        # Auto-fit columns D and I so long strings display fully
        self.auto_fit_columns(dest, ['D', 'I'])

        self.logger.debug(f"Copied {source_range} to sheet {dest_sheet_name} as formulas + formats")

        return dest

    def export_to_pdf(self, sheet: xw.Sheet, output_path: str) -> bool:
        """Export worksheet to PDF"""
        try:
            # Ensure output directory exists
            output_path_obj = Path(output_path)
            output_path_obj.parent.mkdir(parents=True, exist_ok=True)

            # Convert to absolute path
            abs_output_path = str(output_path_obj.absolute())

            self.logger.debug(f"Exporting PDF to: {abs_output_path}")

            # Ensure print area matches B1:I46
            sheet.page_setup.print_area = 'B1:I46'
            sheet.page_setup.orientation = 'portrait'
            # Scale to fit columns — ensures no designated columns are cut off in PDF
            sheet.page_setup.fit_to_width = 1
            sheet.page_setup.fit_to_height = 0  # allow multiple rows if needed
            sheet.page_setup.zoom = 100

            # Export to PDF using Excel API
            sheet.api.ExportAsFixedFormat(
                Type=0,  # xlTypePDF
                Filename=abs_output_path,
                Quality=0,  # xlQualityStandard
                IncludeDocProperties=False,
                IgnorePrintAreas=False,
                OpenAfterPublish=False
            )

            self.logger.info(f"PDF exported: {abs_output_path}")
            return True

        except Exception as e:
            self.logger.error(f"PDF export failed: {e}")
            import traceback
            self.logger.error(traceback.format_exc())
            return False

    def update_formula_references(self, sheet_name: str, old_row: int, new_row: int):
        """
        Update formula references in a sheet to point to a different row.
        Finds formulas like =TBLT!A2 and updates to =TBLT!A5 (for new_row=5)
        """
        sheet = self.wb.sheets[sheet_name]
        updated_count = 0

        # Iterate through all cells in the used range
        used_range = sheet.used_range
        for row in used_range:
            for cell in row:
                # Use .formula to get the formula string, not .value (calculated result)
                formula = getattr(cell, 'formula', None)
                if formula and isinstance(formula, str) and formula.startswith('='):
                    # Check if this formula references TBLT
                    if 'TBLT!' in formula.upper():
                        # Update the row number in the formula
                        # Pattern: TBLT![A-Z]+number (e.g., TBLT!A2, TBLT!$A$2, etc.)
                        # Handle both absolute ($A$2) and relative (A2) references
                        # Use raw string with explicit group references
                        pattern = r'(TBLT!\$?)([A-Z]+)\$?' + str(old_row) + r'(\b|\$)'
                        # Use named groups or explicit string formatting to avoid digit issues
                        replacement = rf'\g<1>\g<2>{new_row}\g<3>'

                        updated_formula = re.sub(
                            pattern,
                            replacement,
                            formula,
                            flags=re.IGNORECASE
                        )

                        if updated_formula != formula:
                            cell.formula = updated_formula
                            updated_count += 1

        if updated_count > 0:
            self.logger.info(f"Updated {updated_count} formulas in {sheet_name}: row {old_row} -> {new_row}")

    def restore_default_references(self, sheet_name: str, default_row: int = 2):
        """
        Restore all TBLT references in a sheet to point to the default row.
        This is needed before processing each trade to ensure clean state.
        """
        sheet = self.wb.sheets[sheet_name]
        restored_count = 0

        # Find all TBLT references and restore to default row
        used_range = sheet.used_range
        for row in used_range:
            for cell in row:
                # Use .formula to get the formula string
                formula = getattr(cell, 'formula', None)
                if formula and isinstance(formula, str) and formula.startswith('='):
                    if 'TBLT!' in formula.upper():
                        original_formula = formula
                        # Match any row number after TBLT! and replace with default
                        pattern = r'(TBLT!\$?)([A-Z]+)\$?\d+(\b|\$)'
                        # Use named group references to avoid digit issues
                        replacement = rf'\g<1>\g<2>{default_row}\g<3>'

                        restored_formula = re.sub(
                            pattern,
                            replacement,
                            formula,
                            flags=re.IGNORECASE
                        )

                        if restored_formula != formula:
                            cell.formula = restored_formula
                            restored_count += 1
                            self.logger.debug(f"  {cell.address}: {original_formula} -> {restored_formula}")

        self.logger.info(f"Restored {restored_count} formulas in {sheet_name} to row {default_row}")
        return restored_count

    def debug_formulas(self, sheet_name: str):
        """Debug: Print ALL formulas in a sheet"""
        sheet = self.wb.sheets[sheet_name]
        formulas_found = []

        used_range = sheet.used_range
        for row in used_range:
            for cell in row:
                # Use .formula to get the formula, not .value (which gives calculated result)
                formula = getattr(cell, 'formula', None)
                if formula and isinstance(formula, str) and formula.startswith('='):
                    formulas_found.append(f"{cell.address}: {formula}")

        self.logger.info(f"Found {len(formulas_found)} total formulas in {sheet_name}:")
        for f in formulas_found[:20]:  # Show first 20
            self.logger.info(f"  {f}")
        if len(formulas_found) > 20:
            self.logger.info(f"  ... and {len(formulas_found) - 20} more")

        return formulas_found
