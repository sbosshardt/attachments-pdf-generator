#!/usr/bin/env python3
"""
Configuration file containing paths and constants used across the application.
"""

import os

# Input file paths
EXCEL_FILE = os.path.join('input-files', 'input-pdfs.xlsx')
TITLE_PAGE = os.path.join('input-files', 'title-page.pdf')
FOREWORD_PAGE = os.path.join('input-files', 'foreword.pdf')

# Output file paths
OUTPUT_DIR = 'output-files'
OUTPUT_TOC = os.path.join(OUTPUT_DIR, 'toc-coverpage.pdf')
OUTPUT_HTML = os.path.join(OUTPUT_DIR, 'toc-debug.html')
OUTPUT_PDF = os.path.join(OUTPUT_DIR, 'weasyoutput.pdf')
MERGED_PDF = os.path.join(OUTPUT_DIR, 'merged-attachments.pdf')

# Sheet name in Excel
SHEET_NAME = "Attachments Prep" 