# EML to PDF Converter

A simple Python utility that converts `.eml` email files to PDF format. Extracts email content (headers and body) and creates clean, readable PDF documents.

## Features

- 📧 Converts `.eml` files to PDF format
- 📝 Extracts email headers (From, To, Cc, Subject)
- 🔤 Smart text extraction from HTML emails
- 🌏 Supports Asian characters (CJK fonts)
- 📂 Batch processing with merge option
- 🪟 Windows-optimized with cross-platform fallback

## Installation

### Prerequisites
- Python 3.6 or higher

### Install Dependencies
```bash
pip install reportlab beautifulsoup4
```

## Usage

### Quick Start
1. Run the script:
```bash
python eml_to_pdf.py
```

2. Enter the folder path containing your `.eml` files when prompted

3. Choose whether to merge all emails into a single PDF or create individual PDFs

### Example
```
=== EML to PDF Text Extractor ===
Enter the folder path containing .eml files: C:\Emails
Merge all emails into ONE single PDF? (Y/N): y

Scanning for .eml files...
Found 5 files. Extracting clean text...
  Processing: email1.eml
  Processing: email2.eml
  ...

Saving merged PDF to: C:\Emails\Merged_Emails.pdf

Done!
```

## How It Works

1. **Scanning**: Recursively searches for all `.eml` files in the specified directory
2. **Parsing**: Extracts email headers and body content (supports both HTML and plain text)
3. **Cleaning**: Intelligently converts HTML to readable text without garbled formatting
4. **PDF Generation**: Creates PDFs with proper font support for international characters

### HTML to Text Conversion
The tool smartly handles HTML emails by:
- Removing scripts, styles, and invisible elements
- Converting block-level elements to proper paragraphs
- Preserving text flow without chopped words
- Handling table cells with proper spacing

## Font Support

The tool attempts to register fonts in this order:
1. **Microsoft YaHei** (Chinese/Japanese/Korean support)
2. **Arial** (Windows standard)
3. **Helvetica** (Cross-platform fallback)

## Output

- **Individual mode**: Creates `[filename].pdf` for each `.eml` file
- **Merge mode**: Creates `Merged_Emails.pdf` with all emails separated by dividers

## Error Handling

The tool gracefully handles:
- Missing dependencies (with installation instructions)
- Invalid folder paths
- Corrupted or unreadable email files
- Missing fonts (falls back to available fonts)

## File Structure

```
eml-to-pdf/
├── eml_to_pdf.py    # Main script
└── README.md        # This file
```

## Technical Details

### Dependencies
- **reportlab**: PDF generation library
- **beautifulsoup4**: HTML parsing and text extraction

### Supported Email Formats
- HTML emails (preferred)
- Plain text emails
- Multi-part MIME messages

### Platform Support
- Windows (primary) - uses Windows Fonts directory
- Linux/macOS - falls back to Helvetica

## Limitations

- Does not extract email attachments
- Does not preserve HTML styling (extracts text only)
- No command-line arguments (interactive only)
- Requires manual font path configuration for non-Windows systems

## License

MIT License - feel free to use and modify as needed.

## Contributing

This is a simple utility script. To contribute:
1. Fork the repository
2. Make your changes following the code style in `AGENTS.md`
3. Test with various `.eml` file formats
4. Submit a pull request

## Troubleshooting

### "Missing libraries!" Error
Install the required dependencies:
```bash
pip install reportlab beautifulsoup4
```

### Garbled Text in PDF
The tool automatically tries different fonts. If you see garbled characters:
- On Windows: Ensure `C:\Windows\Fonts\msyh.ttc` exists
- On other systems: Modify font paths in the script or use the Helvetica fallback

### No Output
- Verify the folder path is correct
- Ensure the folder contains `.eml` files
- Check file permissions

## Future Enhancements

- [ ] Command-line argument support
- [ ] Attachment extraction
- [ ] Custom font configuration
- [ ] Progress bar for large batches
- [ ] Logging instead of print statements
- [ ] Unit tests
