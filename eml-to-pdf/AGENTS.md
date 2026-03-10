# AGENTS.md - Coding Guidelines for eml-to-pdf

## Project Overview
Simple Python utility that converts `.eml` email files to PDF format. Single-file project using ReportLab and BeautifulSoup4.

## Commands

### Running the Script
```bash
python eml_to_pdf.py
```

### Dependencies
Install required packages:
```bash
pip install reportlab beautifulsoup4
```

### Linting & Formatting (Recommended)
Since this project lacks formal test/build configs, use these commands:
```bash
# Format with black
black eml_to_pdf.py

# Lint with ruff
ruff check eml_to_pdf.py

# Type check with mypy
mypy eml_to_pdf.py
```

### Testing
No test suite exists. To test functionality:
1. Place `.eml` files in a test directory
2. Run `python eml_to_pdf.py`
3. Follow interactive prompts

## Code Style Guidelines

### Imports
- Group imports: stdlib → third-party → local
- Use absolute imports
- Handle optional imports gracefully with try/except

Example from codebase:
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

### Naming Conventions
- `snake_case` for functions and variables: `clean_html_text`, `output_path`
- `PascalCase` for class names (when added)
- `UPPER_CASE` for module-level constants: `FONT_NORMAL`
- Descriptive names: `parsed_emails` not `pe`

### Types
- No strict type hints currently used
- Add type hints for new functions:
  ```python
  def extract_email_data(eml_path: Path) -> dict | None:
  ```

### Functions
- Use docstrings for all public functions
- Keep functions focused (SRP)
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
- Provide meaningful error messages
- Handle optional dependencies gracefully
- Use try/except blocks for file I/O

Pattern:
```python
try:
    with open(eml_path, 'rb') as f:
        data = process(f)
except Exception as e:
    print(f"[Error] Could not read {eml_path}: {e}")
    return None
```

### Comments
- Use inline comments sparingly, only for complex logic
- Prefer descriptive variable names over comments
- Use docstrings for module and function documentation

### Main Entry Point
Always wrap CLI execution in `if __name__ == "__main__":`:
```python
if __name__ == "__main__":
    main()
```

## Dependencies
- **reportlab**: PDF generation
- **beautifulsoup4**: HTML parsing
- Standard library: `email`, `pathlib`, `xml.sax.saxutils`, `re`

## Platform Notes
- Windows-focused (uses `C:\Windows\Fonts` paths)
- Falls back to Helvetica on non-Windows systems

## Future Improvements
- Add proper test suite with pytest
- Add pyproject.toml with build configuration
- Add type hints throughout
- Add CLI argument parsing (argparse/click)
- Add logging instead of print statements
