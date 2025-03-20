#!/usr/bin/env python3
"""
Merge PDF Attachments

This script merges the table of contents and cover pages PDF with the actual PDF attachments.
"""

import os
import sys
import json

# Add the src directory to the Python path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

# Import from modules
import fitz  # PyMuPDF
from src.config.paths import OUTPUT_PDF, MERGED_PDF
from src.config.constants import TITLE_PAGE, FOREWORD_PAGE
from src.pdf.pdf_merger import merge_pdfs
from src.excel.excel_reader import load_attachments_from_excel
from src.utils.logger import setup_logger

# Setup logging
logger = setup_logger()

def main():
    # Load attachments from Excel file
    try:
        attachments = load_attachments_from_excel()
        
        if not attachments:
            print("No attachments found in Excel file")
            return 1
            
        print(f"Found {len(attachments)} attachments")
        
        # Check if TOC PDF exists
        if not os.path.exists(OUTPUT_PDF):
            print(f"Table of contents PDF not found: {OUTPUT_PDF}")
            return 1
        
        # Merge PDFs
        merge_pdfs(OUTPUT_PDF, attachments, MERGED_PDF)
        
        return 0
    
    except Exception as e:
        print(f"Error merging PDFs: {e}")
        import traceback
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    sys.exit(main()) 