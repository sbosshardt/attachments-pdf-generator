#!/usr/bin/env python3
# Check bookmarks in a PDF file

import sys
import os

# Get the absolute path to the virtual environment's Python interpreter
venv_python = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'venv', 'bin', 'python')
if os.path.exists(venv_python):
    # If we're not already using the venv Python, re-execute this script with it
    if sys.executable != venv_python:
        print(f"Switching to virtual environment Python: {venv_python}")
        os.execl(venv_python, venv_python, __file__)

# Get the PDF file path
if len(sys.argv) < 2:
    pdf_path = os.path.join('output-files', 'merged-attachments.pdf')
    print(f"No PDF file specified, using default: {pdf_path}")
else:
    pdf_path = sys.argv[1]

if not os.path.exists(pdf_path):
    print(f"PDF file not found: {pdf_path}")
    sys.exit(1)

try:
    # Using pikepdf instead of PyMuPDF
    import pikepdf
    from pikepdf import Pdf

    # Open the PDF and get bookmarks
    pdf = Pdf.open(pdf_path)
    
    # Extract bookmarks (outlines)
    def extract_bookmarks(pdf):
        bookmarks = []
        
        with pdf.open_outline() as outline:
            if not outline.root:
                return bookmarks
                
            def process_bookmark(bookmark, level=1):
                if bookmark.title and bookmark.destination:
                    if isinstance(bookmark.destination, int):
                        # Direct page number
                        page_num = bookmark.destination
                    elif isinstance(bookmark.destination, list) and len(bookmark.destination) > 0:
                        # Array destination, first element is the page reference
                        page_ref = bookmark.destination[0]
                        if hasattr(page_ref, 'objgen'):
                            # Convert page reference to page number
                            try:
                                page_num = pdf.pages.index(pdf.get_object(page_ref.objgen[0])) + 1
                            except:
                                page_num = 0
                        else:
                            page_num = 0
                    else:
                        page_num = 0
                        
                    bookmarks.append((level, bookmark.title, page_num))
                
                # Process children
                for child in bookmark.children:
                    process_bookmark(child, level + 1)
            
            # Process top-level bookmarks
            for item in outline.root:
                process_bookmark(item)
                
        return bookmarks
    
    toc = extract_bookmarks(pdf)
    
    print(f'\nNumber of bookmarks: {len(toc)}')
    print('\nAll bookmarks:')
    for i, item in enumerate(toc):
        level, title, page = item
        # Page numbers are already 1-based in pikepdf
        print(f'  {title} -> page {page}')
    
    print('\nForeword bookmark:')
    foreword_found = False
    for item in toc:
        if 'Foreword' in item[1]:
            level, title, page = item
            print(f'  {title} -> page {page}')
            foreword_found = True
    
    if not foreword_found:
        print("No foreword bookmark found!")
    
    # Close the document
    pdf.close()

except Exception as e:
    print(f"Error: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1) 