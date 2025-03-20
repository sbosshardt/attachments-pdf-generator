#!/bin/bash

# Script to generate the TOC+coverpage PDF and merge with attachments

# Print colored output
GREEN='\033[0;32m'
RED='\033[0;31m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# Activate virtual environment
if [ -d "venv" ]; then
    echo -e "${BLUE}Activating virtual environment...${NC}"
    source venv/bin/activate
else
    echo -e "${RED}Virtual environment not found. Please create one first with 'python -m venv venv' and install requirements.${NC}"
    exit 1
fi

echo -e "${BLUE}[1/2] Generating Table of Contents and Cover Pages PDF${NC}"
python src/generate_toc.py
if [ $? -ne 0 ]; then
    echo -e "${RED}Error generating Table of Contents PDF. Exiting.${NC}"
    exit 1
fi

echo -e "${BLUE}[2/2] Merging with attachment PDFs${NC}"
python src/merge_pdfs.py
if [ $? -ne 0 ]; then
    echo -e "${RED}Error merging PDFs. Exiting.${NC}"
    exit 1
fi

echo -e "${GREEN}Process completed successfully!${NC}"
echo -e "Output files are in the output-files directory:"
echo -e "  - weasyoutput.pdf (Table of Contents and Cover Pages)"
echo -e "  - merged-attachments.pdf (Final merged document)" 