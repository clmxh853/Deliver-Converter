"""
Trade Confirmation Generator - Main Script

Generates trade confirmations from Confo_Template_20230317.xlsx:
- Reads trade data from TBLT sheet
- Uses Template sheet to generate confirmations
- Exports each confirmation to PDF
- Saves output workbook with all confirmation sheets
"""

import os
import sys
import logging
from datetime import datetime
from pathlib import Path

# Import our modules
from excel_automation import ExcelAutomation
from trade_processor import TradeProcessor


def setup_logging(log_folder: str = ".") -> Path:
    """Configure logging for the application"""

    log_folder = Path(log_folder)
    log_folder.mkdir(exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = log_folder / f"confo_generator_{timestamp}.log"

    # Configure root logger
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s | %(levelname)-8s | %(name)s | %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )

    return log_file


def print_summary(summary: dict):
    """Print and log processing summary"""

    print("\n" + "=" * 60)
    print("PROCESSING SUMMARY")
    print("=" * 60)
    print(f"Total rows scanned:     {summary['total_scanned']}")
    print(f"Trades processed:      {summary['processed']}")
    print(f"Confirmations created: {summary['successes']}")
    print(f"PDFs exported:         {len(summary['pdfs_created'])}")
    print(f"Skipped (empty):       {summary['skipped']}")
    print(f"Errors:                {summary['errors']}")

    if summary['errors_detail']:
        print("\nError Details:")
        for error in summary['errors_detail']:
            print(f"  - {error}")

    if summary['pdfs_created']:
        print("\nGenerated PDFs:")
        for pdf in summary['pdfs_created']:
            print(f"  - {pdf}")

    print("=" * 60)


def main():
    """Main entry point"""
    # Setup logging
    log_file = setup_logging()
    logger = logging.getLogger(__name__)

    # Configuration
    INPUT_FILE = "Confo_Template_test.xlsx"
    OUTPUT_FOLDER = Path(__file__).parent / "pdf_output"

    logger.info("=" * 60)
    logger.info("Trade Confirmation Generator Starting")
    logger.info(f"Input file: {INPUT_FILE}")
    logger.info(f"Output folder: {OUTPUT_FOLDER}")
    logger.info("=" * 60)

    try:
        # Get absolute path for input file
        input_path = Path(__file__).parent / INPUT_FILE

        if not input_path.exists():
            logger.error(f"Input file not found: {input_path}")
            return 1

        # Initialize Excel automation
        with ExcelAutomation(str(input_path), visible=False) as excel_app:
            # Create trade processor
            processor = TradeProcessor(excel_app, OUTPUT_FOLDER)

            # Process all trades
            summary = processor.process_all_trades()

        # Print summary
        print_summary(summary)

        logger.info("Processing complete")

        return 0 if summary['errors'] == 0 else 1

    except Exception as e:
        logger.error(f"Fatal error: {e}", exc_info=True)
        return 1


if __name__ == "__main__":
    sys.exit(main())
