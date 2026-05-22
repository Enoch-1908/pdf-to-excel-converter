#!/usr/bin/env python3
"""
Unit tests for PDF to Excel converter.
"""

import unittest
from pathlib import Path
from unittest.mock import patch, MagicMock
import pandas as pd

from pdf_converter import PDFtoExcelConverter


class TestPDFtoExcelConverter(unittest.TestCase):
    """Test cases for PDFtoExcelConverter class."""

    def setUp(self):
        """Set up test fixtures."""
        self.converter = PDFtoExcelConverter(
            input_path='test_input.pdf',
            output_path='test_output.xlsx'
        )

    def test_initialization(self):
        """Test converter initialization."""
        self.assertEqual(str(self.converter.input_path), 'test_input.pdf')
        self.assertEqual(str(self.converter.output_path), 'test_output.xlsx')
        self.assertEqual(self.converter.tables, [])
        self.assertEqual(self.converter.metadata, {})

    @patch('pdf_converter.pdfplumber.open')
    def test_extract_tables_success(self, mock_pdf_open):
        """Test successful table extraction."""
        # Mock PDF structure
        mock_pdf = MagicMock()
        mock_page = MagicMock()
        
        sample_table = [
            ['Header1', 'Header2', 'Header3'],
            ['Value1', 'Value2', 'Value3'],
            ['Value4', 'Value5', 'Value6'],
        ]
        
        mock_page.extract_tables.return_value = [sample_table]
        mock_pdf.pages = [mock_page]
        mock_pdf_open.return_value.__enter__.return_value = mock_pdf
        
        result = self.converter.extract_tables()
        
        self.assertTrue(result)
        self.assertEqual(len(self.converter.tables), 1)
        self.assertEqual(self.converter.metadata['total_pages'], 1)

    @patch('pdf_converter.pdfplumber.open')
    def test_extract_tables_no_tables(self, mock_pdf_open):
        """Test extraction with no tables found."""
        mock_pdf = MagicMock()
        mock_page = MagicMock()
        
        mock_page.extract_tables.return_value = None
        mock_pdf.pages = [mock_page]
        mock_pdf_open.return_value.__enter__.return_value = mock_pdf
        
        result = self.converter.extract_tables()
        
        self.assertFalse(result)
        self.assertEqual(len(self.converter.tables), 0)

    @patch('pdf_converter.Workbook')
    def test_create_excel(self, mock_workbook):
        """Test Excel file creation."""
        # Add sample data
        df = pd.DataFrame({
            'Column1': ['A', 'B', 'C'],
            'Column2': [1, 2, 3],
        })
        self.converter.tables = [df]
        self.converter.metadata = {'total_pages': 1}
        
        mock_wb = MagicMock()
        mock_workbook.return_value = mock_wb
        mock_ws = MagicMock()
        mock_wb.create_sheet.return_value = mock_ws
        
        result = self.converter.create_excel()
        
        self.assertTrue(result)
        mock_wb.save.assert_called_once()

    @patch('pdf_converter.Workbook')
    def test_create_excel_error(self, mock_workbook):
        """Test Excel creation error handling."""
        mock_workbook.side_effect = Exception("Workbook error")
        
        result = self.converter.create_excel()
        
        self.assertFalse(result)

    def test_metadata_structure(self):
        """Test metadata structure."""
        self.converter.metadata = {
            'total_pages': 2,
            'source': 'test.pdf',
        }
        
        self.assertIn('total_pages', self.converter.metadata)
        self.assertEqual(self.converter.metadata['total_pages'], 2)


class TestDataFrameOperations(unittest.TestCase):
    """Test DataFrame operations."""

    def test_dataframe_creation(self):
        """Test DataFrame creation from table data."""
        table = [
            ['Name', 'Age', 'City'],
            ['John', '30', 'NYC'],
            ['Jane', '25', 'LA'],
        ]
        
        df = pd.DataFrame(table[1:], columns=table[0])
        
        self.assertEqual(df.shape, (2, 3))
        self.assertEqual(list(df.columns), ['Name', 'Age', 'City'])

    def test_dataframe_empty(self):
        """Test empty DataFrame handling."""
        df = pd.DataFrame()
        
        self.assertEqual(df.shape[0], 0)
        self.assertEqual(len(df.columns), 0)


if __name__ == '__main__':
    unittest.main()
