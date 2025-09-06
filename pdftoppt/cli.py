"""
Command-line interface for pdftoppt package.
"""

import argparse
import logging
import sys
from pathlib import Path
from typing import Optional, Tuple

from .converter import AdvancedPDFToPowerPointConverter


def setup_logging(verbose: bool = False) -> None:
    """Set up logging configuration."""
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )


def parse_page_range(page_range_str: str) -> Tuple[int, int]:
    """
    Parse page range string into tuple of integers.
    
    Args:
        page_range_str: String in format "start-end" or "start,end"
        
    Returns:
        Tuple of (start_page, end_page)
        
    Raises:
        ValueError: If format is invalid
    """
    if '-' in page_range_str:
        parts = page_range_str.split('-')
    elif ',' in page_range_str:
        parts = page_range_str.split(',')
    else:
        raise ValueError("Page range must be in format 'start-end' or 'start,end'")
    
    if len(parts) != 2:
        raise ValueError("Page range must contain exactly two values")
    
    try:
        start_page = int(parts[0].strip())
        end_page = int(parts[1].strip())
        return start_page, end_page
    except ValueError:
        raise ValueError("Page range values must be integers")


def main():
    """Main entry point for the CLI."""
    parser = argparse.ArgumentParser(
        description="Convert PDF documents to PowerPoint presentations with high fidelity",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  pdftoppt input.pdf output.pptx
  pdftoppt input.pdf output.pptx --pages 1-5
  pdftoppt input.pdf output.pptx --pages 2,8 --verbose
        """
    )
    
    parser.add_argument(
        "input_pdf",
        type=str,
        help="Path to the input PDF file"
    )
    
    parser.add_argument(
        "output_pptx",
        type=str,
        help="Path for the output PowerPoint file"
    )
    
    parser.add_argument(
        "--pages", "-p",
        type=str,
        help="Page range to convert (e.g., '1-5' or '2,8'). If not specified, all pages will be converted."
    )
    
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Enable verbose logging"
    )
    
    parser.add_argument(
        "--version",
        action="version",
        version="%(prog)s 1.0.0"
    )
    
    args = parser.parse_args()
    
    # Set up logging
    setup_logging(args.verbose)
    logger = logging.getLogger(__name__)
    
    # Validate input file
    input_path = Path(args.input_pdf)
    if not input_path.exists():
        logger.error(f"Input file does not exist: {input_path}")
        sys.exit(1)
    
    if not input_path.suffix.lower() == '.pdf':
        logger.error(f"Input file must be a PDF: {input_path}")
        sys.exit(1)
    
    # Validate output path
    output_path = Path(args.output_pptx)
    if not output_path.suffix.lower() in ['.pptx', '.ppt']:
        logger.error(f"Output file must be a PowerPoint file (.pptx or .ppt): {output_path}")
        sys.exit(1)
    
    # Create output directory if it doesn't exist
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    # Parse page range if provided
    page_range: Optional[Tuple[int, int]] = None
    if args.pages:
        try:
            page_range = parse_page_range(args.pages)
            logger.info(f"Converting pages {page_range[0]} to {page_range[1]}")
        except ValueError as e:
            logger.error(f"Invalid page range: {e}")
            sys.exit(1)
    
    # Perform conversion
    try:
        with AdvancedPDFToPowerPointConverter() as converter:
            success = converter.convert(
                pdf_path=str(input_path),
                output_path=str(output_path),
                page_range=page_range
            )
            
            if success:
                logger.info(f"Successfully converted {converter.slides_created} pages")
                logger.info(f"Output saved to: {output_path}")
                sys.exit(0)
            else:
                logger.error("Conversion failed")
                sys.exit(1)
                
    except FileNotFoundError as e:
        logger.error(f"File not found: {e}")
        sys.exit(1)
    except ValueError as e:
        logger.error(f"Invalid parameters: {e}")
        sys.exit(1)
    except Exception as e:
        logger.error(f"Unexpected error during conversion: {e}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
