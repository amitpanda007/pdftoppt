"""Tests for pdftoppt package."""

import unittest
from unittest.mock import patch, MagicMock
import tempfile
import os
from pdftoppt import AdvancedPDFToPowerPointConverter


class TestAdvancedPDFToPowerPointConverter(unittest.TestCase):
    """Test cases for the AdvancedPDFToPowerPointConverter class."""

    def setUp(self):
        """Set up test fixtures."""
        self.converter = AdvancedPDFToPowerPointConverter()

    def tearDown(self):
        """Clean up after tests."""
        self.converter._cleanup_temp_files()

    def test_initialization(self):
        """Test converter initialization."""
        self.assertIsNotNone(self.converter.temp_dir)
        self.assertTrue(os.path.exists(self.converter.temp_dir))
        self.assertEqual(self.converter.slides_created, 0)

    def test_context_manager(self):
        """Test context manager functionality."""
        with AdvancedPDFToPowerPointConverter() as converter:
            self.assertIsNotNone(converter.temp_dir)
        # After context exit, temp directory should be cleaned up

    def test_convert_color_tuple(self):
        """Test color tuple conversion."""
        # Test valid color tuple
        color = (0.5, 0.7, 0.9)
        rgb = self.converter._convert_color_tuple(color)
        self.assertEqual(rgb[0], 127)  # 0.5 * 255
        self.assertEqual(rgb[1], 178)  # 0.7 * 255
        self.assertEqual(rgb[2], 229)  # 0.9 * 255

        # Test None color
        self.assertIsNone(self.converter._convert_color_tuple(None))

        # Test invalid color tuple
        self.assertIsNone(self.converter._convert_color_tuple((0.5,)))

    def test_convert_srgb_color(self):
        """Test sRGB color conversion."""
        # Test red color (0xFF0000)
        rgb = self.converter._convert_srgb_color(0xFF0000)
        self.assertEqual(rgb[0], 255)
        self.assertEqual(rgb[1], 0)
        self.assertEqual(rgb[2], 0)

        # Test None color
        rgb = self.converter._convert_srgb_color(None)
        self.assertEqual(rgb[0], 0)
        self.assertEqual(rgb[1], 0)
        self.assertEqual(rgb[2], 0)

    def test_convert_file_not_found(self):
        """Test conversion with non-existent file."""
        with self.assertRaises(FileNotFoundError):
            self.converter.convert("nonexistent.pdf", "output.pptx")

    @patch('fitz.open')
    def test_convert_invalid_page_range(self, mock_fitz):
        """Test conversion with invalid page range."""
        # Mock a PDF document with 10 pages
        mock_doc = MagicMock()
        mock_doc.__len__.return_value = 10
        mock_fitz.return_value = mock_doc
        
        # Create a temporary file that exists but won't be opened by fitz
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp_file:
            tmp_file.write(b"dummy pdf content")
            tmp_file.flush()
            tmp_file_path = tmp_file.name

        try:
            # Test invalid page range types (strings instead of integers)
            with self.assertRaises(ValueError) as cm:
                self.converter.convert(tmp_file_path, "output.pptx", ("1", "2"))
            self.assertIn("must be integers", str(cm.exception))

            # Test start > end
            with self.assertRaises(ValueError) as cm:
                self.converter.convert(tmp_file_path, "output.pptx", (5, 2))
            self.assertIn("cannot be greater than", str(cm.exception))

            # Test out of bounds page range
            with self.assertRaises(ValueError) as cm:
                self.converter.convert(tmp_file_path, "output.pptx", (1, 15))
            self.assertIn("out of bounds", str(cm.exception))

        finally:
            # Clean up the temporary file
            try:
                os.unlink(tmp_file_path)
            except (OSError, PermissionError):
                pass  # Ignore cleanup errors on Windows

    @patch("fitz.open")
    @patch("pdftoppt.converter.Presentation")
    def test_successful_conversion(self, mock_presentation, mock_fitz):
        """Test successful conversion (mocked)."""
        # Mock PDF document
        mock_doc = MagicMock()
        mock_doc.__len__.return_value = 2
        mock_page = MagicMock()
        mock_page.rect.width = 612
        mock_page.rect.height = 792
        mock_page.get_drawings.return_value = []
        mock_page.get_images.return_value = []
        mock_page.get_text.return_value = {"blocks": []}
        mock_doc.load_page.return_value = mock_page
        mock_fitz.return_value = mock_doc

        # Mock presentation
        mock_prs = MagicMock()
        mock_presentation.return_value = mock_prs

        # Create temporary files without opening them for PyMuPDF
        with tempfile.NamedTemporaryFile(
            suffix=".pdf", delete=False
        ) as pdf_file, tempfile.NamedTemporaryFile(
            suffix=".pptx", delete=False
        ) as pptx_file:

            pdf_file.write(b"dummy pdf content")
            pdf_file.flush()
            pdf_file_path = pdf_file.name
            pptx_file_path = pptx_file.name

        try:
            result = self.converter.convert(pdf_file_path, pptx_file_path)
            self.assertTrue(result)
            mock_prs.save.assert_called_once_with(pptx_file_path)

        finally:
            # Clean up temporary files
            for file_path in [pdf_file_path, pptx_file_path]:
                try:
                    if os.path.exists(file_path):
                        os.unlink(file_path)
                except (OSError, PermissionError):
                    pass  # Ignore cleanup errors on Windows


class TestCLI(unittest.TestCase):
    """Test cases for CLI functionality."""

    def test_parse_page_range(self):
        """Test page range parsing."""
        from pdftoppt.cli import parse_page_range

        # Test dash format
        self.assertEqual(parse_page_range("1-5"), (1, 5))

        # Test comma format
        self.assertEqual(parse_page_range("2,8"), (2, 8))

        # Test invalid formats
        with self.assertRaises(ValueError):
            parse_page_range("invalid")

        with self.assertRaises(ValueError):
            parse_page_range("1-2-3")


if __name__ == "__main__":
    unittest.main()
