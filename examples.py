#!/usr/bin/env python3
"""
Usage examples for PDF to Excel converter.
"""

from pdf_converter import PDFtoExcelConverter
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)


def example_basic_conversion():
    """
    Example 1: Basic PDF to Excel conversion.
    """
    print("\n" + "="*50)
    print("Example 1: Basic PDF to Excel Conversion")
    print("="*50)
    
    converter = PDFtoExcelConverter(
        input_path='sample.pdf',
        output_path='output.xlsx'
    )
    
    success = converter.convert()
    
    if success:
        print("✓ Conversion successful!")
        print(f"  Input: sample.pdf")
        print(f"  Output: output.xlsx")
        print(f"  Tables extracted: {len(converter.tables)}")
    else:
        print("✗ Conversion failed!")


def example_multiple_files():
    """
    Example 2: Convert multiple PDF files.
    """
    print("\n" + "="*50)
    print("Example 2: Convert Multiple PDF Files")
    print("="*50)
    
    pdf_files = [
        ('bank_statement_jan.pdf', 'jan_output.xlsx'),
        ('bank_statement_feb.pdf', 'feb_output.xlsx'),
        ('bank_statement_mar.pdf', 'mar_output.xlsx'),
    ]
    
    for input_file, output_file in pdf_files:
        print(f"\nProcessing: {input_file}")
        converter = PDFtoExcelConverter(
            input_path=input_file,
            output_path=output_file
        )
        
        if converter.convert():
            print(f"✓ {output_file} created successfully")
        else:
            print(f"✗ Failed to convert {input_file}")


def example_batch_processing():
    """
    Example 3: Batch process PDFs from a directory.
    """
    print("\n" + "="*50)
    print("Example 3: Batch Process PDFs from Directory")
    print("="*50)
    
    from pathlib import Path
    
    pdf_directory = Path('pdfs')
    output_directory = Path('excel_output')
    output_directory.mkdir(exist_ok=True)
    
    # Find all PDF files
    pdf_files = list(pdf_directory.glob('*.pdf'))
    
    if not pdf_files:
        print(f"No PDF files found in {pdf_directory}")
        return
    
    print(f"Found {len(pdf_files)} PDF file(s)")
    
    for pdf_file in pdf_files:
        output_file = output_directory / pdf_file.stem + '.xlsx'
        
        print(f"\nProcessing: {pdf_file.name}")
        converter = PDFtoExcelConverter(
            input_path=str(pdf_file),
            output_path=str(output_file)
        )
        
        if converter.convert():
            print(f"✓ {output_file.name} created")
        else:
            print(f"✗ Failed to convert {pdf_file.name}")


def example_custom_handling():
    """
    Example 4: Custom handling of tables.
    """
    print("\n" + "="*50)
    print("Example 4: Custom Table Handling")
    print("="*50)
    
    converter = PDFtoExcelConverter(
        input_path='sample.pdf',
        output_path='output.xlsx'
    )
    
    # Extract tables
    if converter.extract_tables():
        print(f"Extracted {len(converter.tables)} table(s)")
        
        # Access individual tables
        for idx, table in enumerate(converter.tables, 1):
            print(f"\nTable {idx}:")
            print(f"  Shape: {table.shape}")
            print(f"  Columns: {list(table.columns)}")
            print(f"  Rows: {len(table)}")
            
            # Display first few rows
            print("\n  First 3 rows:")
            print(table.head(3))
        
        # Create Excel file
        if converter.create_excel():
            print("\n✓ Excel file created successfully")
    else:
        print("✗ Failed to extract tables")


if __name__ == '__main__':
    print("PDF to Excel Converter - Usage Examples")
    print("\nChoose an example to run:")
    print("1. Basic conversion")
    print("2. Multiple files")
    print("3. Batch processing")
    print("4. Custom handling")
    
    # Uncomment the example you want to run:
    # example_basic_conversion()
    # example_multiple_files()
    # example_batch_processing()
    # example_custom_handling()
