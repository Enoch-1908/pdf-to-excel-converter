# PDF to Excel Converter

A powerful Python tool to convert PDF documents (bank statements, registers, financial reports) into formatted Excel spreadsheets.

## Features

✨ **Core Features:**
- Extract tabular data from PDF files automatically
- Convert to professionally formatted Excel files
- Auto-adjust column widths for readability
- Include metadata sheets with document information
- Handle bank statements, registers, and financial reports
- Command-line interface for easy automation
- Comprehensive error handling and logging
- Unit tests included

## Installation

### Prerequisites
- Python 3.8+
- pip package manager

### Setup

1. Clone the repository:
```bash
git clone https://github.com/Enoch-1908/pdf-to-excel-converter.git
cd pdf-to-excel-converter
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Basic Conversion

```bash
python pdf_converter.py --input sample.pdf --output output.xlsx
```

### Generate Sample Excel

```bash
python create_sample.py
```

### Run Tests

```bash
python -m pytest tests/
```

## Project Structure

```
pdf-to-excel-converter/
├── pdf_converter.py       # Main converter module
├── create_sample.py       # Sample data generator
├── examples.py            # Usage examples
├── requirements.txt       # Python dependencies
├── SETUP.md              # Detailed setup guide
├── README.md             # This file
└── tests/
    └── test_converter.py  # Unit tests
```

## Supported File Types

- PDF with tables and structured data
- Bank statements
- Financial registers
- Tabular reports

## Configuration

See `SETUP.md` for detailed configuration options.

## Contributing

Contributions are welcome! Please feel free to submit pull requests.

## License

Apache License 2.0

## Support

For issues and feature requests, please open an issue on GitHub.
