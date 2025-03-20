#!/usr/bin/env python3
"""
Script to check bookmarks in the merged PDF file.
"""

import os
import sys

# Get the absolute path to the virtual environment's Python interpreter
venv_python = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'venv', 'bin', 'python')
if os.path.exists(venv_python):
    # If we're not already using the venv Python, re-execute this script with it
    if sys.executable != venv_python:
        print(f"Switching to virtual environment Python: {venv_python}")
        os.execl(venv_python, venv_python, __file__)

# Add the src directory to the Python path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

# Import from the existing modules
from src.config.paths import MERGED_PDF

try:
    # Try various imports to handle PyMuPDF
    try:
        import fitz
        print("Using direct 'import fitz'")
    except ImportError:
        try:
            from pymupdf import fitz
            print("Using 'from pymupdf import fitz'")
        except ImportError:
            try:
                from PyMuPDF import fitz
                print("Using 'from PyMuPDF import fitz'")
            except ImportError:
                print("Cannot import fitz module. Make sure PyMuPDF is installed.")
                sys.exit(1)
    
    # Open the merged PDF
    pdf_path = MERGED_PDF
    if not os.path.exists(pdf_path):
        print(f"PDF file not found: {pdf_path}")
        sys.exit(1)
    
    doc = fitz.open(pdf_path)
    toc = doc.get_toc()
    
    print(f"\nNumber of bookmarks: {len(toc)}")
    print("\nFirst 5 bookmarks:")
    for i, item in enumerate(toc[:5]):
        level, title, page = item
        print(f"Bookmark {i+1}: Level {level}, Title: '{title}', Page: {page+1}")
    
    print("\nForeword bookmark:")
    foreword_found = False
    for item in toc:
        if 'Foreword' in item[1]:
            level, title, page = item
            print(f"Level {level}, Title: '{title}', Page: {page+1}")
            foreword_found = True
    
    if not foreword_found:
        print("No foreword bookmark found!")
    
    doc.close()

except Exception as e:
    print(f"Error: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1) 