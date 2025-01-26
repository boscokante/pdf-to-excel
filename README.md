# PDF Credit Report Parser

A Python script that extracts tradeline information from credit report PDFs and outputs to both text and Excel formats.

## Features

- Extracts bank name, account number, and reported balance from credit report PDFs
- Outputs data to both text and Excel files
- Formats currency in Excel with proper formatting and column widths
- Automatically creates input/output directories
- Protects sensitive data through .gitignore

## Prerequisites

- Python 3.9 or higher
- pip (Python package installer)

## Quick Start

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Place your credit report PDF in the `input` directory

3. Run:
```bash
python credit_parser.py
```

4. Find results in `output` directory:
   - `tradeline_data.txt` - Plain text format
   - `tradeline_data.xlsx` - Excel format with currency formatting

## Detailed Setup

### Using venv (recommended)

```bash
# Create virtual environment
python -m venv venv

# Activate virtual environment
# On macOS/Linux:
source venv/bin/activate
# On Windows:
.\venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

## Directory Structure

```
.
├── input/           # Place PDF files here
├── output/          # Results appear here
├── credit_parser.py # Main script
├── requirements.txt # Dependencies
└── README.md
```

## Troubleshooting

Common issues and solutions:

- **ModuleNotFoundError**: Run `pip install -r requirements.txt`
- **No PDF files found**: Ensure PDF is in the `input` directory
- **Excel formatting issues**: Ensure xlsxwriter is installed
- **PDF parsing errors**: Check PDF format matches expected structure

## Notes

- Script processes the first PDF it finds in input directory
- Directories are git-ignored to protect sensitive data
- Excel output includes:
  - Currency formatting ($XXX.XX)
  - Adjusted column widths
  - Clean headers
