#!/usr/bin/env python3
"""
Configuration file containing constants used across the application.
"""

import os

# Input file paths (referencing from paths.py to maintain consistency)
from src.config.paths import TITLE_PAGE, FOREWORD_PAGE

# PDF related constants
PAGE_SIZE = (8.5, 11)  # Letter size in inches
MARGIN = 1  # 1 inch margins
FONT_SIZE = 12  # Default font size
HEADING_FONT_SIZE = 18  # Heading font size

# TOC Constants
TOC_ENTRIES_PER_PAGE = 25  # Approximate number of TOC entries per page

# PDF Bookmark Levels
BOOKMARK_LEVEL_1 = 1  # Top level bookmark
BOOKMARK_LEVEL_2 = 2  # Second level bookmark 