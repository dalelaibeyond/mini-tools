# Mini Tools Collection

A collection of Python utilities for handling `.eml` email files.

## Tools Overview

| Tool | Description | Output Formats |
|------|-------------|----------------|
| **eml-to-pdf** | Converts email files to PDF | PDF |
| **eml_separator** | Separates emails and attachments | EML / TXT / PDF |

## Quick Start

### 1. Install Dependencies

```bash
pip install reportlab beautifulsoup4
```

### 2. Run a Tool

```bash
# Convert emails to PDF
python eml-to-pdf/eml_to_pdf.py

# Separate emails and attachments
python eml_separator/eml_separator.py
```

## eml-to-pdf

Converts `.eml` email files to PDF format with clean text extraction.

**Features:**
- Batch processing of multiple emails
- Smart HTML-to-text conversion
- Merge option (single PDF or individual files)
- Asian character support (CJK fonts)

**Usage:**
1. Run the script
2. Enter folder path containing `.eml` files
3. Choose merge option (Y/N)
4. Find PDFs in the same folder

## eml_separator

Separates emails into content and extracts attachments into organized folders.

**Features:**
- Three output formats: EML (original), TXT (clean text), PDF (formatted)
- Automatic attachment extraction
- Organized folder structure
- Duplicate handling

**Usage:**
1. Run the script
2. Enter folder path containing `.eml` files
3. Select output format (1=EML, 2=TXT, 3=PDF)
4. Find organized output in:
   - `email-list/` - Email content
   - `email-attachments/` - Extracted files

## Requirements

- Python 3.6+
- reportlab (PDF generation)
- beautifulsoup4 (HTML parsing)

## Platform Support

- **Windows** (optimized) - Uses Windows Fonts
- **macOS/Linux** - Falls back to Helvetica font

## File Structure

```
mini-tools/
├── eml-to-pdf/
│   ├── eml_to_pdf.py
│   └── README.md
├── eml_separator/
│   └── eml_separator.py
└── README.md (this file)
```

## License

MIT License - Free to use and modify.
