# PDFToPPT - Advanced PDF to PowerPoint Converter

[![PyPI version](https://badge.fury.io/py/pdftoppt.svg)](https://badge.fury.io/py/pdftoppt)
[![Python Support](https://img.shields.io/pypi/pyversions/pdftoppt.svg)](https://pypi.org/project/pdftoppt/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

A high-fidelity Python package for converting PDF documents to PowerPoint presentations. This tool preserves layouts, images, text formatting, and vector graphics with exceptional accuracy.

## Features

- **High Fidelity Conversion**: Preserves original PDF layouts, fonts, colors, and formatting
- **Vector Graphics Support**: Converts PDF vector elements (lines, rectangles) to PowerPoint shapes
- **Image Preservation**: Extracts and embeds images with transparency support
- **Text Formatting**: Maintains font styles, sizes, colors, bold, and italic formatting
- **Custom Page Ranges**: Convert specific pages or page ranges
- **Batch Processing**: Process multiple pages efficiently
- **Command Line Interface**: Easy-to-use CLI for automation
- **Python API**: Full programmatic access for integration

## Installation

### From PyPI (Recommended)

```bash
pip install pdftoppt
```

### From Source

```bash
git clone https://github.com/amitpanda007/pdftoppt.git
cd pdftoppt
pip install -e .
```

## Dependencies

- Python 3.7+
- PyMuPDF (fitz) >= 1.20.0
- python-pptx >= 0.6.18
- Pillow >= 8.0.0

## Quick Start

### Command Line Usage

```bash
# Convert entire PDF
pdftoppt input.pdf output.pptx

# Convert specific page range
pdftoppt input.pdf output.pptx --pages 1-5

# Convert with verbose logging
pdftoppt input.pdf output.pptx --verbose

# Get help
pdftoppt --help
```

### Python API Usage

```python
from pdftoppt import AdvancedPDFToPowerPointConverter

# Basic conversion
with AdvancedPDFToPowerPointConverter() as converter:
    success = converter.convert("input.pdf", "output.pptx")
    print(f"Conversion successful: {success}")
    print(f"Slides created: {converter.slides_created}")

# Convert specific pages
with AdvancedPDFToPowerPointConverter() as converter:
    success = converter.convert(
        pdf_path="input.pdf",
        output_path="output.pptx",
        page_range=(1, 5)  # Convert pages 1-5
    )

# With error handling
try:
    converter = AdvancedPDFToPowerPointConverter()
    converter.convert("input.pdf", "output.pptx")
except FileNotFoundError:
    print("PDF file not found")
except ValueError as e:
    print(f"Invalid parameters: {e}")
finally:
    converter._cleanup_temp_files()
```

## Advanced Usage

### Logging Configuration

```python
import logging
from pdftoppt import AdvancedPDFToPowerPointConverter

# Enable debug logging
logging.basicConfig(level=logging.DEBUG)

converter = AdvancedPDFToPowerPointConverter()
converter.convert("input.pdf", "output.pptx")
```

### Batch Processing

```python
import os
from pathlib import Path
from pdftoppt import AdvancedPDFToPowerPointConverter

def batch_convert_pdfs(input_dir, output_dir):
    \"\"\"Convert all PDFs in a directory.\"\"\"
    input_path = Path(input_dir)
    output_path = Path(output_dir)
    output_path.mkdir(exist_ok=True)

    with AdvancedPDFToPowerPointConverter() as converter:
        for pdf_file in input_path.glob("*.pdf"):
            output_file = output_path / f"{pdf_file.stem}.pptx"
            try:
                success = converter.convert(str(pdf_file), str(output_file))
                print(f"{'✓' if success else '✗'} {pdf_file.name}")
            except Exception as e:
                print(f"✗ {pdf_file.name}: {e}")

# Usage
batch_convert_pdfs("./pdfs", "./presentations")
```

## How It Works

The converter uses a multi-step process to ensure high-fidelity conversion:

1. **PDF Analysis**: Extracts text, images, and vector graphics from each PDF page
2. **Element Processing**: Processes fonts, colors, positioning, and formatting
3. **PowerPoint Generation**: Creates custom-sized presentation matching PDF dimensions
4. **Content Reconstruction**: Rebuilds all elements as native PowerPoint objects

## Supported Elements

- ✅ Text with formatting (fonts, sizes, colors, bold, italic)
- ✅ Images (JPEG, PNG) with transparency support
- ✅ Vector graphics (rectangles, lines)
- ✅ Colors and fill patterns
- ✅ Positioning and layouts
- ⚠️ Complex vector paths (simplified to basic shapes)
- ❌ Interactive elements (forms, hyperlinks)
- ❌ Animations and transitions

## Performance

Typical conversion speeds:

- Simple text documents: ~2-5 pages/second
- Image-heavy documents: ~0.5-2 pages/second
- Complex mixed content: ~1-3 pages/second

Memory usage scales with document complexity and image content.

## Troubleshooting

### Common Issues

**Import Error for PyMuPDF:**

```bash
pip install --upgrade PyMuPDF
```

**Memory issues with large PDFs:**

```python
# Process in smaller page ranges
for start in range(1, total_pages, 10):
    end = min(start + 9, total_pages)
    converter.convert("large.pdf", f"output_part_{start}.pptx",
                     page_range=(start, end))
```

**Font rendering issues:**

- Ensure system has required fonts installed
- Check PDF for embedded fonts

### Debug Mode

Enable verbose logging to diagnose issues:

```bash
pdftoppt input.pdf output.pptx --verbose
```

## API Reference

### AdvancedPDFToPowerPointConverter

#### Methods

**`__init__()`**

- Initializes converter with temporary directory for processing

**`convert(pdf_path, output_path, page_range=None)`**

- Main conversion method
- **Parameters:**
  - `pdf_path` (str): Path to input PDF file
  - `output_path` (str): Path for output PowerPoint file
  - `page_range` (tuple, optional): (start_page, end_page) for partial conversion
- **Returns:** bool - True if successful
- **Raises:** FileNotFoundError, ValueError

**Context Manager Support:**

```python
with AdvancedPDFToPowerPointConverter() as converter:
    converter.convert("input.pdf", "output.pptx")
# Automatic cleanup
```

## Contributing

We welcome contributions! Please see our [Contributing Guide](CONTRIBUTING.md) for details.

### Development Setup

```bash
git clone https://github.com/amitpanda007/pdftoppt.git
cd pdftoppt
pip install -e ".[dev]"
```

### Running Tests

```bash
pytest tests/
```

### Code Quality

```bash
black pdftoppt/
flake8 pdftoppt/
mypy pdftoppt/
```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Changelog

### v1.0.0

- Initial release
- High-fidelity PDF to PowerPoint conversion
- Support for text, images, and vector graphics
- Command-line interface
- Python API with context manager support

## Support

- 📖 [Documentation](https://github.com/amitpanda007/pdftoppt#readme)
- 🐛 [Bug Reports](https://github.com/amitpanda007/pdftoppt/issues)
- 💬 [Discussions](https://github.com/amitpanda007/pdftoppt/discussions)

## Acknowledgments

- Built with [PyMuPDF](https://pymupdf.readthedocs.io/) for PDF processing
- Uses [python-pptx](https://python-pptx.readthedocs.io/) for PowerPoint generation
- Image processing powered by [Pillow](https://pillow.readthedocs.io/)

---

Made with ❤️ for the Python community
