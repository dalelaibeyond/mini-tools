import os
import email
from email import policy
from pathlib import Path
import re
import shutil
import xml.sax.saxutils as saxutils

try:
    from bs4 import BeautifulSoup
except ImportError:
    print("Missing library! Please open Command Prompt and run:")
    print("pip install beautifulsoup4")
    input("\nPress Enter to exit...")
    exit()

# We only strictly need reportlab if the user chooses PDF. 
PDF_SUPPORT = True
try:
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.pdfbase import pdfmetrics
    
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
except ImportError:
    PDF_SUPPORT = False


def clean_filename(name):
    """Removes invalid characters from Windows file/folder names."""
    if not name:
        return "Untitled_Email"
    clean_name = re.sub(r'[<>:"/\\|?*\n\r\t]', '_', str(name))
    return clean_name[:100].strip()

def clean_html_text(html_content):
    """Smartly extracts text from HTML exactly like a web browser does."""
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Remove invisible formatting code
    for element in soup(["script", "style", "meta", "title", "noscript"]):
        element.extract()
        
    # Replace hard line breaks
    for br in soup.find_all("br"):
        br.replace_with("\n")
        
    # Surround blocks with newlines to create proper paragraphs
    block_tags =["p", "div", "h1", "h2", "h3", "h4", "h5", "h6", "tr", "li", "ul", "ol", "blockquote"]
    for block in soup.find_all(block_tags):
        block.insert_before("\n")
        block.insert_after("\n")
        
    for cell in soup.find_all(["td", "th"]):
        cell.insert_after(" ")
        
    raw_text = soup.get_text()
    
    # Clean up the text lines without deleting any actual content/URLs
    lines = raw_text.split('\n')
    cleaned_lines =[]
    
    for line in lines:
        clean_line = " ".join(line.split())
        if clean_line:
            cleaned_lines.append(clean_line)
            
    return "\n".join(cleaned_lines)

def process_email(eml_path, index, email_list_dir, attachments_dir, output_format):
    """Parses the .eml file, separates attachments, and outputs the chosen format."""
    try:
        with open(eml_path, 'rb') as f:
            msg = email.message_from_binary_file(f, policy=policy.default)
    except Exception as e:
        print(f"  [Error] Could not read {eml_path.name}: {e}")
        return False

    # Extract Headers
    subject = str(msg.get('subject', 'No_Subject'))
    sender = str(msg.get('from', 'Unknown'))
    to_addr = str(msg.get('to', ''))
    cc_addr = str(msg.get('cc', ''))
    date = str(msg.get('date', 'Unknown Date'))

    safe_subject = clean_filename(subject)
    identifier = f"{index:03d}_{safe_subject}"

    attachments =[]
    plain_text = ""
    html_text = ""

    # Walk through the email parts to find Attachments and Text
    for part in msg.walk():
        if part.is_multipart():
            continue

        content_type = part.get_content_type()
        content_disp = str(part.get('Content-Disposition', ''))
        filename = part.get_filename()

        # Is it an attachment?
        if filename or 'attachment' in content_disp:
            payload = part.get_payload(decode=True)
            if payload:
                attachments.append({
                    'filename': clean_filename(filename),
                    'data': payload
                })
            continue

        # If not an attachment, read the body text
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

    # Save Attachments to "email-attachments" folder (if any exist)
    if attachments:
        email_attach_folder = attachments_dir / identifier
        email_attach_folder.mkdir(exist_ok=True)
        for att in attachments:
            att_path = email_attach_folder / att['filename']
            try:
                with open(att_path, 'wb') as att_file:
                    att_file.write(att['data'])
            except Exception as e:
                print(f"  [Error] Could not save attachment {att['filename']}: {e}")

    # Process and Output Email Content based on chosen format
    if output_format == '1':
        # EML FORMAT: Copy the exact original file to preserve formatting completely
        dest_path = email_list_dir / f"{identifier}.eml"
        shutil.copy2(eml_path, dest_path)

    else:
        # TXT or PDF FORMAT: We need to clean and structure the text
        if html_text.strip():
            body_text = clean_html_text(html_text)
        elif plain_text.strip():
            body_text = plain_text
        else:
            body_text = "[No readable text found in this email]"

        headers_dict = {
            'Subject': subject,
            'From': sender,
            'To': to_addr,
            'Cc': cc_addr,
            'Date': date
        }

        if output_format == '2':
            # TXT FORMAT
            text_output_path = email_list_dir / f"{identifier}.txt"
            try:
                with open(text_output_path, 'w', encoding='utf-8') as text_file:
                    for key, val in headers_dict.items():
                        if val and val.strip() != 'None':
                            text_file.write(f"{key}: {val}\n")
                    text_file.write("\n--------------------------------------------------------\n\n")
                    text_file.write(body_text)
            except Exception as e:
                print(f"  [Error] Could not save text file for {identifier}: {e}")

        elif output_format == '3' and PDF_SUPPORT:
            # PDF FORMAT
            pdf_output_path = email_list_dir / f"{identifier}.pdf"
            try:
                doc = SimpleDocTemplate(str(pdf_output_path), pagesize=letter)
                styles = getSampleStyleSheet()
                custom_style = ParagraphStyle(
                    'CustomStyle', parent=styles['Normal'], fontName=FONT_NORMAL,
                    wordWrap='CJK', fontSize=10, leading=14
                )
                
                story =[]
                for key, val in headers_dict.items():
                    if val and val.strip() != 'None':
                        safe_val = saxutils.escape(val)
                        story.append(Paragraph(f"<font name='{FONT_BOLD}'>{key}:</font> {safe_val}", custom_style))
                
                story.append(Spacer(1, 15))
                story.append(Paragraph("--------------------------------------------------------", custom_style))
                story.append(Spacer(1, 15))
                
                for line in body_text.split('\n'):
                    clean_line = line.strip()
                    if clean_line:
                        story.append(Paragraph(saxutils.escape(clean_line), custom_style))
                        
                doc.build(story)
            except Exception as e:
                print(f"  [Error] Failed to build PDF for {identifier}: {e}")

    return True

if __name__ == "__main__":
    print("=== EML Text and Attachment Separator ===")
    
    # 1. Get Folder Path (Safely replacing quotes to prevent SyntaxErrors)
    folder_input = input("Enter the folder path containing .eml files: ")
    folder_input = folder_input.replace('"', '').replace("'", "").strip()
    
    if not os.path.isdir(folder_input):
        print("Invalid folder path!")
        input("Press Enter to exit...")
        exit()

    # 2. Get User Format Choice
    print("\nSelect the output format for the emails (Attachments will always be extracted):")
    print("  1) .eml  (Original format, perfectly preserved)")
    print("  2) .txt  (Clean text, easy to read)")
    print("  3) .pdf  (Formatted PDF document)")
    
    format_choice = input("Enter 1, 2, or 3: ").strip()
    if format_choice not in ['1', '2', '3']:
        print("Invalid choice. Defaulting to .txt (2).")
        format_choice = '2'
        
    if format_choice == '3' and not PDF_SUPPORT:
        print("\n[Warning] ReportLab is not installed. Falling back to .txt output.")
        print("To enable PDF, run: pip install reportlab")
        format_choice = '2'

    base_path = Path(folder_input)
    email_list_dir = base_path / "email-list"
    attachments_dir = base_path / "email-attachments"
    
    email_list_dir.mkdir(exist_ok=True)
    attachments_dir.mkdir(exist_ok=True)

    print("\nScanning for .eml files...")
    eml_files = list(base_path.rglob("*.eml"))
    
    if not eml_files:
        print("No .eml files found in the specified directory.")
        input("Press Enter to exit...")
        exit()

    print(f"Found {len(eml_files)} files. Processing...\n")
    
    success_count = 0
    for index, file_path in enumerate(eml_files, start=1):
        # Skip files that were already moved into our own output folder
        if "email-list" in file_path.parts or "email-attachments" in file_path.parts:
            continue
            
        print(f"  Processing[{index}/{len(eml_files)}]: {file_path.name}")
        if process_email(file_path, index, email_list_dir, attachments_dir, format_choice):
            success_count += 1

    print("\n=== Done! ===")
    print(f"Successfully processed {success_count} emails.")
    print(f"Emails saved to: {email_list_dir}")
    print(f"Attachments saved to: {attachments_dir}")
    input("Press Enter to exit...")