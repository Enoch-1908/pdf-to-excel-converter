# Setup Guide - PDF to Excel Converter

## Prerequisites

- Python 3.8 or higher
- pip (Python package installer)
- Virtual environment (recommended)

## Installation Steps

### 1. Clone the Repository

```bash
git clone https://github.com/Enoch-1908/pdf-to-excel-converter.git
cd pdf-to-excel-converter
```

### 2. Create Virtual Environment (Recommended)

**On macOS/Linux:**
```bash
python3 -m venv venv
source venv/bin/activate
```

**On Windows:**
```bash
python -m venv venv
venv\Scripts\activate
```

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

## Usage

### Command Line

```bash
python pdf_converter.py --input input.pdf --output output.xlsx
```

**Options:**
- `--input, -i`: Path to input PDF file (required)
- `--output, -o`: Path to output Excel file (required)

### Generate Sample Excel

```bash
python create_sample.py
```

This creates a `sample_bank_register.xlsx` file with sample data.

### Run Examples

```bash
python examples.py
```

Edit the file to uncomment the example you want to run.

## Running Tests

```bash
python -m pytest tests/ -v
```

## Configuration

### Environment Variables (Optional)

Create a `.env` file in the project root:

```
LOG_LEVEL=INFO
OUTPUT_DIR=./output
```

## Troubleshooting

### Issue: "pdfplumber: No module named"
**Solution:** Ensure all dependencies are installed:
```bash
pip install -r requirements.txt --upgrade
```

### Issue: "PDF file not found"
**Solution:** Verify the PDF file path is correct and the file exists.

### Issue: "No tables extracted"
**Solution:** 
- Ensure the PDF contains tables
- Try with a different PDF file
- Check if the PDF has scanned images instead of text

### Issue: "Permission denied when saving Excel"
**Solution:** 
- Close any open Excel files with the same name
- Ensure you have write permissions in the output directory

## Project Structure

```
pdf-to-excel-converter/
├── pdf_converter.py          # Main converter class
├── create_sample.py          # Sample data generator
├── examples.py               # Usage examples
├── requirements.txt          # Python dependencies
├── SETUP.md                 # This file
├── README.md                # Project documentation
├── tests/                   # Test directory
│   └── test_converter.py    # Unit tests
└── .env.example             # Example environment variables
```

## Dependencies Overview

- **pandas**: Data manipulation and analysis
- **openpyxl**: Excel file creation and formatting
- **pdfplumber**: PDF table extraction
- **PyPDF2**: PDF parsing and manipulation
- **tabula-py**: Alternative PDF table extraction
- **python-dotenv**: Environment variable management

## Advanced Configuration

### Logging

Adjust logging level in `pdf_converter.py`:

```python
logging.basicConfig(level=logging.DEBUG)  # For more verbose output
```

### Custom Formatting

Modify the `_format_sheet()` method in `pdf_converter.py` to customize Excel formatting.

## Next Steps

1. Generate sample data: `python create_sample.py`
2. Review the sample Excel output
3. Try converting your own PDF files
4. Read examples.py for advanced usage

## Support

For issues or questions, please open an issue on GitHub.
