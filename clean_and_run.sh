#!/bin/bash

# Script to clean output directory and run the PDF generation process

# Define colors for output
GREEN='\033[0;32m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

echo -e "${BLUE}Setting up environment...${NC}"

# Create output directory if it doesn't exist
mkdir -p output-files
echo "Created output-files directory (if it didn't exist)"

# Clean the output directory without prompting, but preserve .gitignore
if [ -n "$(ls -A output-files 2>/dev/null)" ]; then
    echo "Cleaning output-files directory (preserving .gitignore)..."
    # First check if .gitignore exists
    if [ -f "output-files/.gitignore" ]; then
        # Delete all files except .gitignore
        find output-files -type f -not -name ".gitignore" -delete
    else
        # If no .gitignore, delete all files
        find output-files -type f -delete
    fi
else
    echo "Output directory is already empty"
fi

# Activate virtual environment
echo -e "${BLUE}Activating virtual environment...${NC}"
source ./venv/bin/activate

# Run the main script
echo -e "${BLUE}Starting PDF generation...${NC}"
./generate_and_merge.sh

echo -e "${GREEN}Process completed!${NC}"
echo "Final output files should be in the output-files directory" 