import os
import email
from email import policy
from pathlib import Path
import xml.sax.saxutils as saxutils
import re

try:
    from bs4 import BeautifulSoup
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.pdfbase import pdfmetrics
except ImportError:
    print("Missing libraries! Please open Command Prompt and run:")
    print("pip install reportlab beautifulsoup4")
    input("\nPress Enter to exit...")
    exit()

# Setup Fonts to prevent Garbled Text
try:
    pdfmetrics.registerFont(TTFont('YaHei', 'C:\\Windows\\Fonts\\msyh.ttc'))
    pdfmetrics.registerFont(TTFont('YaHei-Bold', 'C:\\Windows\\Fonts\\msyhbd.ttc'))
    FONT_NORMAL = 'YaHei'
    FONT_BOLD = 'YaHei-Bold'
except Exception:
    try:
        pdfmetrics.registerFont(TTFont('Arial', 'C:\\Windows\\Fonts\\arial.ttf'))
        pdfmetrics.registerFont(TTFont('Arial-Bold', 'C:\\Windows\\Fonts\\arialbd.ttf'))
        FONT_NORMAL = 'Arial'
        FONT_BOLD = 'Arial-Bold'
    except Exception:
        FONT_NORMAL = 'Helvetica'
        FONT_BOLD = 'Helvetica-Bold'

def clean_html_text(html_content):
    """Smartly extracts text from HTML exactly like a web browser does."""
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # 1. Remove invisible formatting code and scripts
    for element in soup(["script", "style", "meta", "title", "noscript"]):
        element.extract()
        
    # 2. Replace hard line breaks with standard newlines
    for br in soup.find_all("br"):
        br.replace_with("\n")
        
    # 3. Surround block-level elements with newlines to create proper paragraphs
    block_tags =["p", "div", "h1", "h2", "h3", "h4", "h5", "h6", "tr", "li", "ul", "ol", "blockquote"]
    for block in soup.find_all(block_tags):
        block.insert_before("\n")
        block.insert_after("\n")
        
    # Add a space after table cells to prevent words in columns from merging
    for cell in soup.find_all(["td", "th"]):
        cell.insert_after(" ")
        
    # 4. Extract text. Because we injected '\n' manually above, we do NOT use a separator here.
    # This prevents inline styles from chopping words! (e.g., <span>T</span>el becomes "Tel")
    raw_text = soup.get_text()
    
    # 5. Clean up the text lines
    lines = raw_text.split('\n')
    cleaned_lines =[]
    
    for line in lines:
        # Collapse all inner spaces/tabs into a single space
        clean_line = " ".join(line.split())
        if clean_line:
            cleaned_lines.append(clean_line)
            
    # 6. Join paragraphs together and return
    return "\n".join(cleaned_lines)

def extract_email_data(eml_path):
    """Parses the .eml file and extracts headers and clean text."""
    try:
        with open(eml_path, 'rb') as f:
            msg = email.message_from_binary_file(f, policy=policy.default)
    except Exception as e:
        print(f"  [Error] Could not read {eml_path}: {e}")
        return None

    # Extract required fields
    sender = str(msg.get('from', ''))
    to_addr = str(msg.get('to', ''))
    cc_addr = str(msg.get('cc', ''))
    subject = str(msg.get('subject', ''))

    plain_text = ""
    html_text = ""

    # Walk through the email parts
    for part in msg.walk():
        if part.is_multipart():
            continue

        content_type = part.get_content_type()
        content_disp = str(part.get('Content-Disposition', ''))

        # Ignore attachments and images
        if 'attachment' in content_disp or part.get_filename():
            continue

        if content_type == 'text/plain':
            try:
                plain_text += part.get_content() + "\n"
            except Exception:
                plain_text += part.get_payload(decode=True).decode('utf-8', errors='ignore') + "\n"
                
        elif content_type == 'text/html':
            try:
                html_content = part.get_content()
            except Exception:
                html_content = part.get_payload(decode=True).decode('utf-8', errors='ignore')
            html_text += html_content + "\n"

    # STRATEGY: Prefer HTML text and clean it beautifully to avoid chopped words
    if html_text.strip():
        body_text = clean_html_text(html_text)
    elif plain_text.strip():
        # Fallback if the email is strictly plain-text
        body_text = plain_text
        # Remove raw http links to clean up reading
        body_text = re.sub(r'\(?https?://[^\s)]+\)?', '', body_text)
    else:
        body_text = "[No readable text found in this email]"

    return {
        'from': sender,
        'to': to_addr,
        'cc': cc_addr,
        'subject': subject,
        'body': body_text.strip()
    }

def create_pdf(emails_data, output_path):
    """Generates a PDF file from the extracted email data."""
    doc = SimpleDocTemplate(output_path, pagesize=letter)
    styles = getSampleStyleSheet()
    
    # Create a custom style that supports Asian characters and line wrapping
    custom_style = ParagraphStyle(
        'CustomStyle',
        parent=styles['Normal'],
        fontName=FONT_NORMAL,
        wordWrap='CJK',
        fontSize=10,
        leading=14
    )
    
    story =[]

    for i, data in enumerate(emails_data):
        safefrom = saxutils.escape(data['from'])
        safeto = saxutils.escape(data['to'])
        safecc = saxutils.escape(data['cc'])
        safesubject = saxutils.escape(data['subject'])

        story.append(Paragraph(f"<font name='{FONT_BOLD}'>From:</font> {safefrom}", custom_style))
        story.append(Paragraph(f"<font name='{FONT_BOLD}'>To:</font> {safeto}", custom_style))
        
        if safecc.strip() and safecc.strip() != 'None':
            story.append(Paragraph(f"<font name='{FONT_BOLD}'>Cc:</font> {safecc}", custom_style))
            
        story.append(Paragraph(f"<font name='{FONT_BOLD}'>Subject:</font> {safesubject}", custom_style))
        story.append(Spacer(1, 15))

        # Add the clean body text
        for line in data['body'].split('\n'):
            clean_line = line.strip()
            if clean_line:
                story.append(Paragraph(saxutils.escape(clean_line), custom_style))

        # Add separator line if it's not the very last email
        if i < len(emails_data) - 1:
            story.append(Spacer(1, 15)) # Blank line before separator
            story.append(Paragraph("--------------------------------------------------------", custom_style))
            story.append(Spacer(1, 15)) # Blank line after separator

    try:
        if story:
            doc.build(story)
    except Exception as e:
        print(f"[Error] Failed to build PDF {output_path}: {e}")

if __name__ == "__main__":
    print("=== EML to PDF Text Extractor ===")
    folder_input = input("Enter the folder path containing .eml files: ").strip()
    
    folder_input = folder_input.strip('"').strip("'")
    
    if not os.path.isdir(folder_input):
        print("Invalid folder path!")
        input("Press Enter to exit...")
        exit()

    merge_input = input("Merge all emails into ONE single PDF? (Y/N): ").strip().lower()
    merge_files = merge_input == 'y'

    print("\nScanning for .eml files...")
    folder_path = Path(folder_input)
    
    eml_files = list(folder_path.rglob("*.eml"))
    
    if not eml_files:
        print("No .eml files found in the specified directory.")
        input("Press Enter to exit...")
        exit()

    print(f"Found {len(eml_files)} files. Extracting clean text...")
    
    parsed_emails =[]
    for f in eml_files:
        print(f"  Processing: {f.name}")
        data = extract_email_data(f)
        if data:
            parsed_emails.append((f, data))

    if merge_files:
        output_pdf = folder_path / "Merged_Emails.pdf"
        print(f"\nSaving merged PDF to: {output_pdf}")
        create_pdf([data for _, data in parsed_emails], str(output_pdf))
    else:
        print("\nSaving individual PDFs...")
        for f, data in parsed_emails:
            output_pdf = f.with_suffix('.pdf')
            create_pdf([data], str(output_pdf))

    print("\nDone!")
    input("Press Enter to exit...")