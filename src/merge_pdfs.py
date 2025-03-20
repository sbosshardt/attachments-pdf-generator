#!/usr/bin/env python3
"""
Main script for merging PDFs.
"""

import os
import sys
import traceback

# Add the parent directory to sys.path to allow imports from src
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from src.excel.excel_reader import read_attachment_data
from src.pdf.pdf_merger import merge_pdfs
from src.config.paths import OUTPUT_DIR

def main():
    """
    Main function to run the script.
    """
    # Ensure output directory exists
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    try:
        # Read data from Excel
        attachments = read_attachment_data(for_toc=False)
        
        # Merge PDFs
        merge_pdfs(attachments)
        
    except Exception as e:
        print(f"Error: {e}")
        traceback.print_exc()
        return 1
    
    return 0

if __name__ == "__main__":
    sys.exit(main()) 