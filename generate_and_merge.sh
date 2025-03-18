#!/bin/bash

# Script to generate the TOC+coverpage PDF and merge with attachments

# Print colored output
GREEN='\033[0;32m'
RED='\033[0;31m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

echo -e "${BLUE}[1/2] Generating Table of Contents and Cover Pages PDF${NC}"
python3 generate_toc_coverpage.py
if [ $? -ne 0 ]; then
    echo -e "${RED}Error generating Table of Contents PDF. Exiting.${NC}"
    exit 1
fi

echo -e "${BLUE}[2/2] Merging with attachment PDFs${NC}"
python3 merge_pdfs.py
if [ $? -ne 0 ]; then
    echo -e "${RED}Error merging PDFs. Exiting.${NC}"
    exit 1
fi

echo -e "${GREEN}Process completed successfully!${NC}"
echo -e "Output files are in the output-files directory:"
echo -e "  - weasyoutput.pdf (Table of Contents and Cover Pages)"
echo -e "  - merged-attachments.pdf (Final merged document)" 