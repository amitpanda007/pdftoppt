#!/usr/bin/env python3
"""
Example usage of the pdftoppt package.

This script demonstrates various ways to use the PDF to PowerPoint converter.
"""

import logging
from pathlib import Path
from pdftoppt import AdvancedPDFToPowerPointConverter


def basic_example():
    """Basic conversion example."""
    print("=== Basic Conversion Example ===")
    
    # Note: You'll need to provide actual PDF files for testing
    input_pdf = "example.pdf"  # Replace with your PDF file
    output_pptx = "output.pptx"
    
    try:
        with AdvancedPDFToPowerPointConverter() as converter:
            success = converter.convert(input_pdf, output_pptx)
            
            if success:
                print(f"✓ Conversion successful!")
                print(f"✓ Created {converter.slides_created} slides")
                print(f"✓ Output saved to: {output_pptx}")
            else:
                print("✗ Conversion failed")
                
    except FileNotFoundError:
        print(f"✗ PDF file not found: {input_pdf}")
        print("  Please create a test PDF file or update the path")
    except Exception as e:
        print(f"✗ Error: {e}")


def page_range_example():
    """Convert specific page range."""
    print("\n=== Page Range Conversion Example ===")
    
    input_pdf = "example.pdf"
    output_pptx = "output_pages_1_to_3.pptx"
    
    try:
        with AdvancedPDFToPowerPointConverter() as converter:
            success = converter.convert(
                pdf_path=input_pdf,
                output_path=output_pptx,
                page_range=(1, 3)  # Convert only pages 1-3
            )
            
            if success:
                print(f"✓ Converted pages 1-3 successfully!")
                print(f"✓ Created {converter.slides_created} slides")
            else:
                print("✗ Conversion failed")
                
    except FileNotFoundError:
        print(f"✗ PDF file not found: {input_pdf}")
    except ValueError as e:
        print(f"✗ Invalid page range: {e}")
    except Exception as e:
        print(f"✗ Error: {e}")


def batch_conversion_example():
    """Batch convert multiple PDFs."""
    print("\n=== Batch Conversion Example ===")
    
    # Example PDF files (you would replace these with actual files)
    pdf_files = [
        "document1.pdf",
        "document2.pdf", 
        "document3.pdf"
    ]
    
    with AdvancedPDFToPowerPointConverter() as converter:
        for pdf_file in pdf_files:
            try:
                output_file = pdf_file.replace('.pdf', '.pptx')
                success = converter.convert(pdf_file, output_file)
                
                status = "✓" if success else "✗"
                print(f"{status} {pdf_file} -> {output_file}")
                
            except FileNotFoundError:
                print(f"✗ {pdf_file} (file not found)")
            except Exception as e:
                print(f"✗ {pdf_file} (error: {e})")


def verbose_logging_example():
    """Example with verbose logging enabled."""
    print("\n=== Verbose Logging Example ===")
    
    # Enable debug logging
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    
    input_pdf = "example.pdf"
    output_pptx = "output_verbose.pptx"
    
    try:
        with AdvancedPDFToPowerPointConverter() as converter:
            success = converter.convert(input_pdf, output_pptx)
            print(f"Conversion completed: {success}")
            
    except FileNotFoundError:
        print(f"✗ PDF file not found: {input_pdf}")
    except Exception as e:
        print(f"✗ Error: {e}")


def main():
    """Run all examples."""
    print("PDFToPPT Package Examples")
    print("=" * 50)
    
    # Create a sample message about test files
    print("\nNote: These examples expect PDF files to exist.")
    print("Please ensure you have test PDF files or update the file paths.\n")
    
    # Run examples
    basic_example()
    page_range_example()
    batch_conversion_example()
    verbose_logging_example()
    
    print("\n" + "=" * 50)
    print("Examples completed!")
    print("\nFor more information, see:")
    print("- README.md for detailed documentation")
    print("- CLI usage: pdftoppt --help")


if __name__ == "__main__":
    main()
