#!/usr/bin/env python3
"""
Merge PDF Attachments

This script merges the table of contents and cover pages PDF with the actual PDF attachments.
"""

import os
import openpyxl
import fitz  # PyMuPDF

# File paths
EXCEL_FILE = os.path.join('input-files', 'input-pdfs.xlsx')
TOC_PDF = os.path.join('output-files', 'weasyoutput.pdf')
MERGED_PDF = os.path.join('output-files', 'merged-attachments.pdf')

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
        'Filename Reference': ['Filename Reference', 'Filename', 'File'],
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
    required_fields = ['Attachment Number', 'Filename Reference']
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

def merge_pdfs(attachments):
    """
    Merge the table of contents PDF with the attachment PDFs.
    
    Args:
        attachments (list): List of attachment data dictionaries
    """
    print(f"Merging PDFs into: {MERGED_PDF}")
    
    # Check if TOC PDF exists
    if not os.path.exists(TOC_PDF):
        raise FileNotFoundError(f"Table of contents PDF not found: {TOC_PDF}")
    
    # Create a new PDF document
    merged_pdf = fitz.open()
    
    # Add the table of contents PDF
    toc_pdf = fitz.open(TOC_PDF)
    merged_pdf.insert_pdf(toc_pdf)
    
    # Create bookmarks/outline for the PDF
    toc = []
    current_page = len(merged_pdf)  # Start after the TOC pages
    
    # Add each attachment PDF
    for attachment in attachments:
        attachment_num = attachment.get('Attachment Number', '')
        language = attachment.get('Language', 'EN')
        title = attachment.get('Title', 'Untitled')
        filename = attachment.get('Filename Reference', '')
        
        if not filename:
            print(f"Warning: No filename for Attachment {attachment_num}, skipping")
            continue
        
        # Create the full path to the attachment file
        attachment_path = os.path.join('input-files/'+language.lower(), filename)
        
        if not os.path.exists(attachment_path):
            print(f"Warning: Attachment file not found: {attachment_path}, skipping")
            continue
        
        try:
            # Open the attachment PDF
            attachment_pdf = fitz.open(attachment_path)
            
            # Add to bookmarks
            toc.append([1, f"Attachment {attachment_num}: {title}", current_page])
            
            # Add to merged PDF
            merged_pdf.insert_pdf(attachment_pdf)
            
            # Update current page for next bookmark
            current_page += len(attachment_pdf)
            
            print(f"Added: Attachment {attachment_num} - {filename}")
            
        except Exception as e:
            print(f"Error adding attachment {attachment_num}: {e}")
    
    # Set the table of contents
    merged_pdf.set_toc(toc)
    
    # Save the merged PDF
    merged_pdf.save(MERGED_PDF)
    merged_pdf.close()
    
    print(f"Merged PDF created at: {MERGED_PDF}")

def main():
    """
    Main function to run the script.
    """
    # Ensure output directory exists
    os.makedirs('output-files', exist_ok=True)
    
    try:
        # Read data from Excel
        attachments = read_attachment_data()
        
        # Merge PDFs
        merge_pdfs(attachments)
        
    except Exception as e:
        print(f"Error: {e}")
        return 1
    
    return 0

if __name__ == "__main__":
    exit(main()) 