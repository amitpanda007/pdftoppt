# Contributing to PDFToPPT

Thank you for your interest in contributing to PDFToPPT! This document provides guidelines and information for contributors.

## Table of Contents

- [Code of Conduct](#code-of-conduct)
- [Getting Started](#getting-started)
- [Development Setup](#development-setup)
- [How to Contribute](#how-to-contribute)
- [Pull Request Process](#pull-request-process)
- [Code Style Guidelines](#code-style-guidelines)
- [Testing](#testing)
- [Documentation](#documentation)
- [Reporting Issues](#reporting-issues)
- [Feature Requests](#feature-requests)
- [Release Process](#release-process)

## Code of Conduct

By participating in this project, you agree to abide by our Code of Conduct:

- **Be respectful**: Treat everyone with respect and kindness
- **Be inclusive**: Welcome newcomers and help them learn
- **Be constructive**: Provide helpful feedback and suggestions
- **Be patient**: Remember that everyone has different skill levels
- **Be collaborative**: Work together towards common goals

## Getting Started

### Prerequisites

- Python 3.7 or higher
- Git
- Basic knowledge of PDF processing and PowerPoint manipulation

### Fork and Clone

1. Fork the repository on GitHub
2. Clone your fork locally:

   ```bash
   git clone https://github.com/yourusername/pdftoppt.git
   cd pdftoppt
   ```

3. Add the upstream repository:
   ```bash
   git remote add upstream https://github.com/amitpanda007/pdftoppt.git
   ```

## Development Setup

### 1. Create a Virtual Environment

```bash
# Create virtual environment
python -m venv venv

# Activate it (Windows)
venv\Scripts\activate

# Activate it (Unix/macOS)
source venv/bin/activate
```

### 2. Install Dependencies

```bash
# Install the package in development mode
pip install -e ".[dev]"

# Or install dependencies manually
pip install -r requirements.txt
pip install pytest pytest-cov black flake8 mypy
```

### 3. Verify Installation

```bash
# Test the CLI
pdftoppt --help

# Run tests
python -m pytest tests/ -v

# Test code formatting
black --check pdftoppt/
```

## How to Contribute

### Types of Contributions

We welcome various types of contributions:

- **Bug fixes**: Fix issues in the codebase
- **Feature enhancements**: Improve existing functionality
- **New features**: Add new capabilities
- **Documentation**: Improve docs, examples, or guides
- **Tests**: Add or improve test coverage
- **Performance**: Optimize code performance
- **Compatibility**: Support for new Python versions or platforms

### Areas for Contribution

1. **PDF Processing Improvements**:

   - Better text extraction and positioning
   - Enhanced vector graphics support
   - Improved image handling
   - Support for more PDF features

2. **PowerPoint Generation**:

   - Better formatting preservation
   - Advanced slide layouts
   - Animation and transition support
   - Template support

3. **User Experience**:

   - CLI improvements
   - Better error messages
   - Progress indicators
   - Batch processing features

4. **Testing and Quality**:
   - Unit tests for new features
   - Integration tests
   - Performance benchmarks
   - Cross-platform testing

## Pull Request Process

### Before You Start

1. **Check existing issues**: Look for related issues or discussions
2. **Create an issue**: For significant changes, create an issue first
3. **Discuss your approach**: Get feedback before implementing

### Development Workflow

1. **Create a branch**:

   ```bash
   git checkout -b feature/your-feature-name
   # or
   git checkout -b fix/issue-number
   ```

2. **Make your changes**:

   - Write clean, readable code
   - Follow the coding standards
   - Add tests for new functionality
   - Update documentation as needed

3. **Test your changes**:

   ```bash
   # Run all tests
   python -m pytest tests/ -v

   # Run with coverage
   python -m pytest tests/ --cov=pdftoppt --cov-report=html

   # Check code quality
   black pdftoppt/ tests/
   flake8 pdftoppt/
   mypy pdftoppt/
   ```

4. **Commit your changes**:

   ```bash
   git add .
   git commit -m "feat: add support for animated transitions"
   # or
   git commit -m "fix: resolve memory leak in image processing"
   ```

5. **Push and create PR**:
   ```bash
   git push origin feature/your-feature-name
   ```
   Then create a Pull Request on GitHub.

### Commit Message Guidelines

Use conventional commit format:

- `feat:` for new features
- `fix:` for bug fixes
- `docs:` for documentation changes
- `test:` for adding tests
- `refactor:` for code refactoring
- `perf:` for performance improvements
- `chore:` for maintenance tasks

Examples:

```
feat: add support for PDF bookmarks in conversion
fix: resolve issue with transparent images (#123)
docs: update installation instructions
test: add unit tests for page range validation
```

### Pull Request Checklist

Before submitting your PR, ensure:

- [ ] Code follows the project's style guidelines
- [ ] Tests pass locally (`pytest tests/`)
- [ ] New functionality includes tests
- [ ] Documentation is updated (if applicable)
- [ ] Commit messages follow the conventional format
- [ ] PR description clearly explains the changes
- [ ] No merge conflicts with the main branch

## Code Style Guidelines

### Python Code Style

We follow [PEP 8](https://pep8.org/) with some specific guidelines:

#### Formatting

```python
# Use Black for automatic formatting
black pdftoppt/ tests/

# Line length: 88 characters (Black default)
# Use double quotes for strings
# Use type hints for function parameters and return values
```

#### Code Structure

```python
# Good: Clear function with type hints and docstring
def convert_pdf_page(
    page: fitz.Page,
    slide: pptx.slide.Slide
) -> bool:
    """
    Convert a PDF page to a PowerPoint slide.

    Args:
        page: PDF page object from PyMuPDF
        slide: PowerPoint slide object

    Returns:
        True if conversion successful, False otherwise
    """
    try:
        # Implementation here
        return True
    except Exception as e:
        logger.error(f"Conversion failed: {e}")
        return False
```

#### Error Handling

```python
# Good: Specific exceptions with meaningful messages
try:
    pdf_doc = fitz.open(pdf_path)
except FileNotFoundError:
    raise FileNotFoundError(f"PDF file not found: {pdf_path}")
except fitz.FitzError as e:
    raise ValueError(f"Invalid PDF file: {e}")

# Good: Logging instead of print statements
import logging
logger = logging.getLogger(__name__)
logger.info("Starting conversion process")
logger.error(f"Conversion failed: {error_message}")
```

### Documentation Style

#### Docstrings

Use Google-style docstrings:

```python
def extract_images(self, page: fitz.Page) -> List[Dict[str, Any]]:
    """
    Extract images from a PDF page.

    Args:
        page: PDF page object to extract images from

    Returns:
        List of dictionaries containing image data and metadata

    Raises:
        ValueError: If page is invalid
        IOError: If image extraction fails

    Example:
        >>> converter = AdvancedPDFToPowerPointConverter()
        >>> page = pdf_doc.load_page(0)
        >>> images = converter.extract_images(page)
        >>> print(f"Found {len(images)} images")
    """
```

## Testing

### Running Tests

```bash
# Run all tests
python -m pytest tests/

# Run with verbose output
python -m pytest tests/ -v

# Run with coverage
python -m pytest tests/ --cov=pdftoppt --cov-report=html

# Run specific test file
python -m pytest tests/test_converter.py

# Run specific test
python -m pytest tests/test_converter.py::TestConverter::test_convert_basic
```

### Writing Tests

#### Unit Tests

```python
import unittest
from unittest.mock import patch, MagicMock
from pdftoppt import AdvancedPDFToPowerPointConverter

class TestConverter(unittest.TestCase):
    def setUp(self):
        self.converter = AdvancedPDFToPowerPointConverter()

    def tearDown(self):
        self.converter._cleanup_temp_files()

    def test_color_conversion(self):
        """Test color tuple to RGB conversion."""
        color = (0.5, 0.7, 0.9)
        rgb = self.converter._convert_color_tuple(color)
        self.assertEqual(rgb.r, 127)
        self.assertEqual(rgb.g, 178)
        self.assertEqual(rgb.b, 229)

    @patch('fitz.open')
    def test_convert_success(self, mock_fitz):
        """Test successful PDF conversion."""
        # Mock setup
        mock_doc = MagicMock()
        mock_fitz.return_value = mock_doc

        # Test
        result = self.converter.convert("test.pdf", "output.pptx")
        self.assertTrue(result)
```

#### Integration Tests

```python
def test_full_conversion_workflow(self):
    """Test complete conversion with real files."""
    # Use small test PDF files
    test_pdf = "tests/fixtures/simple.pdf"
    output_pptx = "tests/output/test_output.pptx"

    with AdvancedPDFToPowerPointConverter() as converter:
        success = converter.convert(test_pdf, output_pptx)
        self.assertTrue(success)
        self.assertTrue(os.path.exists(output_pptx))
```

### Test Guidelines

- Write tests for all new functionality
- Use descriptive test names
- Test both success and failure cases
- Mock external dependencies (file I/O, network calls)
- Use fixtures for test data
- Keep tests fast and independent

## Documentation

### Types of Documentation

1. **Code Documentation**: Docstrings and inline comments
2. **API Documentation**: Function and class documentation
3. **User Guide**: README.md and usage examples
4. **Developer Guide**: This CONTRIBUTING.md file
5. **Release Notes**: CHANGELOG.md

### Documentation Guidelines

- Keep documentation up-to-date with code changes
- Use clear, concise language
- Include practical examples
- Document edge cases and limitations
- Use proper Markdown formatting

### Building Documentation

```bash
# Generate API documentation (if using Sphinx)
sphinx-build -b html docs/ docs/_build/

# Check documentation locally
# Open README.md in a Markdown preview
```

## Reporting Issues

### Bug Reports

When reporting bugs, please include:

1. **Description**: Clear description of the issue
2. **Steps to reproduce**: Exact steps to trigger the bug
3. **Expected behavior**: What should happen
4. **Actual behavior**: What actually happens
5. **Environment**: Python version, OS, package versions
6. **Sample files**: If possible, provide sample PDF files
7. **Error messages**: Full error messages and stack traces

#### Bug Report Template

```markdown
**Bug Description**
A clear description of what the bug is.

**To Reproduce**
Steps to reproduce the behavior:

1. Run command '...'
2. Use PDF file with '...'
3. See error

**Expected Behavior**
What you expected to happen.

**Screenshots/Output**
If applicable, add screenshots or error output.

**Environment:**

- OS: [e.g., Windows 10, Ubuntu 20.04]
- Python Version: [e.g., 3.9.7]
- Package Version: [e.g., 1.0.0]
- Dependencies: [versions of PyMuPDF, python-pptx, etc.]

**Additional Context**
Any other context about the problem.
```

### Security Issues

For security-related issues:

- **Do not** create public issues
- Email the maintainers directly
- Provide detailed information about the vulnerability
- Allow time for investigation and patching

## Feature Requests

### Suggesting Features

1. **Check existing issues** for similar requests
2. **Create a detailed issue** describing the feature
3. **Explain the use case** and benefits
4. **Discuss implementation** if you have ideas

#### Feature Request Template

```markdown
**Feature Description**
A clear description of the proposed feature.

**Use Case**
Explain why this feature would be useful and what problem it solves.

**Proposed Solution**
If you have ideas about implementation, describe them here.

**Alternatives**
Describe any alternative solutions you've considered.

**Additional Context**
Any other context, screenshots, or examples.
```

## Release Process

### Versioning

We follow [Semantic Versioning](https://semver.org/):

- **MAJOR** version: Breaking changes
- **MINOR** version: New features (backwards compatible)
- **PATCH** version: Bug fixes

### Release Checklist

1. Update version numbers in `setup.py` and `pyproject.toml`
2. Update `CHANGELOG.md` with new features and fixes
3. Run full test suite
4. Build and test the package
5. Create release notes
6. Tag the release
7. Publish to PyPI

### Creating a Release

```bash
# Update version and changelog
# Commit changes
git add .
git commit -m "chore: prepare release v1.1.0"

# Create tag
git tag v1.1.0
git push origin v1.1.0

# Build and upload
python -m build
python -m twine upload dist/*
```

## Getting Help

### Community

- **GitHub Issues**: For bugs and feature requests
- **GitHub Discussions**: For questions and general discussion
- **Email**: Contact maintainers for sensitive issues

### Resources

- [Python Packaging Guide](https://packaging.python.org/)
- [PyMuPDF Documentation](https://pymupdf.readthedocs.io/)
- [python-pptx Documentation](https://python-pptx.readthedocs.io/)
- [Git Workflow Guide](https://docs.github.com/en/get-started/quickstart/github-flow)

## Recognition

Contributors will be recognized in:

- `AUTHORS.md` file
- Release notes
- GitHub contributors page

Thank you for contributing to PDFToPPT! 🚀

---

_This document is a living guide and will be updated as the project evolves. If you have suggestions for improvements, please open an issue or submit a pull request._
