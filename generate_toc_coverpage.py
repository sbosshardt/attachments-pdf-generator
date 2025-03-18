#!/usr/bin/env python3
"""
Generate Table of Contents and Cover Pages PDF

This script reads attachment information from an Excel file and generates
a PDF containing a table of contents and cover pages for each attachment.
"""

import os
import openpyxl
from weasyprint import HTML
from datetime import datetime

# File paths
EXCEL_FILE = os.path.join('input-files', 'input-pdfs.xlsx')
OUTPUT_PDF = os.path.join('output-files', 'weasyoutput.pdf')

def read_attachment_data():
    """
    Read attachment data from Excel file.
    
    Returns:
        list: List of dictionaries containing attachment data
    """
    print(f"Opening Excel file: {EXCEL_FILE}")
    
    # Check if input file exists
    if not os.path.exists(EXCEL_FILE):
        raise FileNotFoundError(f"Excel file not found: {EXCEL_FILE}")
    
    workbook = openpyxl.load_workbook(EXCEL_FILE)
    
    # Check if sheet exists
    if "Attachments Prep" not in workbook.sheetnames:
        raise ValueError(f"Sheet 'Attachments Prep' not found in {EXCEL_FILE}")
    
    sheet = workbook["Attachments Prep"]
    
    # Find headers
    headers = []
    for cell in sheet[1]:
        headers.append(cell.value)
    
    # Field mapping (convert sheet headers to our standardized field names)
    field_mapping = {
        'Attachment Number': ['Attachment Number', 'Attachment #', 'Number'],
        'Title': ['Title', 'Document Title'],
        'Page count': ['Page count', 'Pages', 'Page Count'],
        'Additional Remarks about File': ['Additional Remarks about File', 'Remarks', 'Notes'],
        'Body': ['Body', 'Description', 'Body (Description)'],
        'Filename Reference': ['Filename Reference', 'Filename', 'File'],
        'Date': ['Date', 'Date (time Pacific)', 'Document Date'],
        'Language': ['Language', 'Lang', 'Language Code'],
        'Exclude': ['Exclude', 'Skip']
    }
    
    # Map headers to indices
    header_indices = {}
    for field, possible_headers in field_mapping.items():
        for i, header in enumerate(headers):
            if header and any(possible_match.lower() == header.lower() for possible_match in possible_headers):
                header_indices[field] = i
                break
    
    # Check if we found all required fields
    required_fields = ['Attachment Number', 'Title']
    for field in required_fields:
        if field not in header_indices:
            raise ValueError(f"Required field '{field}' not found in Excel headers")
    
    data = []
    # Start from row 2 (skip header)
    for row in sheet.iter_rows(min_row=2, values_only=True):
        # Skip empty rows
        if not any(row):
            continue
        
        # Skip rows marked as excluded
        exclude_idx = header_indices.get('Exclude')
        if exclude_idx is not None and row[exclude_idx] in (True, 'TRUE', 'True', 'true', 'YES', 'Yes', 'yes', '1', 1):
            continue
        
        attachment = {}
        for field, idx in header_indices.items():
            if field != 'Exclude':  # We've already used this field for filtering
                attachment[field] = row[idx] if idx < len(row) else None
        
        # Set default value for Language if not found
        if 'Language' not in attachment or not attachment['Language']:
            attachment['Language'] = 'EN'
        
        data.append(attachment)
    
    print(f"Found {len(data)} attachments")
    return data

def generate_html(attachments):
    """
    Generate HTML with table of contents and cover pages.
    
    Args:
        attachments (list): List of attachment data dictionaries
    
    Returns:
        str: HTML document as string
    """
    # CSS styling
    css = """
    @page {
        margin: 1cm;
        @top-center {
            content: "Attachments";
        }
        @bottom-center {
            content: counter(page);
        }
    }
    body {
        font-family: Arial, sans-serif;
        line-height: 1.5;
    }
    h1 {
        color: #333;
        font-size: 24pt;
        margin-top: 20pt;
        margin-bottom: 15pt;
        text-align: center;
    }
    h2 {
        font-size: 18pt;
        margin-top: 15pt;
        margin-bottom: 10pt;
    }
    h3 {
        font-size: 14pt;
        margin-top: 12pt;
        margin-bottom: 8pt;
    }
    .toc-entry {
        margin: 5pt 0;
    }
    .cover-page {
        page-break-before: always;
    }
    .first-cover-page {
        page-break-before: always;
    }
    .field-label {
        font-weight: bold;
    }
    .page-break {
        page-break-after: always;
    }
    .metadata {
        margin: 10pt 0;
    }
    """
    
    # Start building HTML
    html = f"""<!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>Attachments</title>
        <style>{css}</style>
    </head>
    <body>
        <h1>Table of Contents</h1>
    """
    
    # Generate Table of Contents
    for attachment in attachments:
        attachment_num = attachment.get('Attachment Number', '')
        title = attachment.get('Title', 'Untitled')
        html += f"""
        <div class="toc-entry">
            <a href="#cover-{attachment_num}">Attachment {attachment_num}</a>: {title}
        </div>
        """
    
    # Add page break after TOC
    html += '<div class="page-break"></div>'
    
    # Generate Cover Pages
    for i, attachment in enumerate(attachments):
        attachment_num = attachment.get('Attachment Number', '')
        title = attachment.get('Title', 'Untitled')
        page_count = attachment.get('Page count', '')
        remarks = attachment.get('Additional Remarks about File', '')
        body = attachment.get('Body', '')
        filename = attachment.get('Filename Reference', '')
        date = attachment.get('Date', '')
        
        # Format date if it's a datetime object
        if isinstance(date, datetime):
            date = date.strftime('%Y-%m-%d %I:%M %p')
        
        # Add cover page class (first one needs special handling)
        cover_class = "first-cover-page" if i == 0 else "cover-page"
        
        html += f"""
        <div id="cover-{attachment_num}" class="{cover_class}">
            <h2>Attachment {attachment_num}</h2>
            <h3>{title}</h3>
        """
        
        # Add metadata fields if they exist
        if date:
            html += f'<div class="metadata"><span class="field-label">Date:</span> {date}</div>'
        
        if page_count:
            html += f'<div class="metadata"><span class="field-label">Pages:</span> {page_count}</div>'
        
        if filename:
            html += f'<div class="metadata"><span class="field-label">File:</span> {filename}</div>'
        
        if remarks:
            html += f'<div class="metadata"><span class="field-label">Remarks:</span> {remarks}</div>'
        
        if body:
            html += f'<div class="metadata"><span class="field-label">Description:</span></div>'
            html += f'<div>{body}</div>'
        
        html += '</div>'
    
    # Close HTML
    html += """
    </body>
    </html>
    """
    
    return html

def main():
    """
    Main function to run the script.
    """
    # Ensure output directory exists
    os.makedirs('output-files', exist_ok=True)
    
    try:
        # Read data from Excel
        attachments = read_attachment_data()
        
        # Generate HTML
        html_content = generate_html(attachments)
        
        # Convert HTML to PDF using WeasyPrint
        print(f"Generating PDF: {OUTPUT_PDF}")
        HTML(string=html_content).write_pdf(OUTPUT_PDF)
        
        print(f"PDF successfully generated at {OUTPUT_PDF}")
        
    except Exception as e:
        print(f"Error: {e}")
        return 1
    
    return 0

if __name__ == "__main__":
    exit(main()) 