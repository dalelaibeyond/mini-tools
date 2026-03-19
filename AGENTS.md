# AGENTS.md - Coding Guidelines for mini-tools

## Project Overview

Collection of Python utilities for handling `.eml` email files. Each tool is a self-contained single-file script.

- **eml-to-pdf/**: Converts `.eml` files to PDF format
- **eml_separator/**: Separates emails into content and attachments (EML/TXT/PDF output)

## Commands

### Running Scripts

```bash
# EML to PDF converter
python eml-to-pdf/eml_to_pdf.py

# EML separator
python eml_separator/eml_separator.py
```

### Install Dependencies

```bash
pip install reportlab beautifulsoup4
```

### Linting & Formatting

Since this project lacks formal test/build configs, use these commands:

```bash
# Format with black
black eml-to-pdf/eml_to_pdf.py eml_separator/eml_separator.py

# Lint with ruff
ruff check eml-to-pdf/eml_to_pdf.py eml_separator/eml_separator.py

# Type check with mypy
mypy eml-to-pdf/eml_to_pdf.py eml_separator/eml_separator.py
```

### Testing

No formal test suite exists. To test functionality:
1. Place `.eml` files in a test directory
2. Run the script
3. Follow interactive prompts

## Code Style Guidelines

### Imports

- Group imports: stdlib → third-party → local
- Use absolute imports
- Handle optional imports gracefully with try/except

Example:

```python
import os
import email
from email import policy
from pathlib import Path

# Optional imports with graceful fallback
try:
    from bs4 import BeautifulSoup
    from reportlab.lib.pagesizes import letter
except ImportError:
    print("Missing libraries!")
    exit()
```

### Formatting

- 4 spaces for indentation (no tabs)
- Max line length: ~100 characters
- Use single quotes for strings: `'string'`
- Add blank lines between logical sections
- Space after list brackets: `block_tags = ["p", "div"]` (note the space)

### Naming Conventions

- `snake_case` for functions and variables: `clean_html_text`, `output_path`
- `PascalCase` for class names (when added)
- `UPPER_CASE` for module-level constants: `FONT_NORMAL`, `PDF_SUPPORT`
- Descriptive names: `parsed_emails` not `pe`

### Types

- No strict type hints currently used
- Add type hints for new functions:

```python
def extract_email_data(eml_path: Path) -> dict | None:
    """Parses the .eml file and extracts headers and clean text."""
```

### Functions

- Use docstrings for all public functions
- Keep functions focused (Single Responsibility Principle)
- Handle errors gracefully with try/except
- Return early for error conditions

Example:

```python
def clean_html_text(html_content: str) -> str:
    """Smartly extracts text from HTML exactly like a web browser does."""
    # Implementation
    return cleaned_text
```

### Error Handling

- Use specific exceptions when possible
- Provide meaningful error messages with `[Error]` prefix
- Handle optional dependencies gracefully
- Use try/except blocks for file I/O
- Continue processing other files on individual errors

Pattern:

```python
try:
    with open(eml_path, 'rb') as f:
        data = process(f)
except Exception as e:
    print(f"  [Error] Could not read {eml_path}: {e}")
    return None
```

### Comments

- Use inline comments sparingly, only for complex logic
- Prefer descriptive variable names over comments
- Use docstrings for module and function documentation
- Numbered comments for multi-step processes: `# 1. Remove scripts`

### Main Entry Point

Always wrap CLI execution in `if __name__ == "__main__":`:

```python
if __name__ == "__main__":
    print("=== Tool Name ===")
    # CLI logic here
```

## Platform Notes

- Windows-focused (uses `C:\Windows\Fonts` paths)
- Falls back to Helvetica on non-Windows systems
- Handle both quoted and unquoted paths from user input
- Use `pathlib.Path` for cross-platform path handling

## Dependencies

- **reportlab**: PDF generation
- **beautifulsoup4**: HTML parsing
- Standard library: `email`, `pathlib`, `xml.sax.saxutils`, `re`, `shutil`

## Future Improvements

- Add proper test suite with pytest
- Add pyproject.toml with build configuration
- Add type hints throughout
- Add CLI argument parsing (argparse/click)
- Add logging instead of print statements
- Support for non-Windows font configuration

## Cursor/Copilot Rules

No specific AI coding assistant rules are configured in this repository.
