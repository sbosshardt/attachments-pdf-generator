# attachments-pdf-generator

Create unified PDF content with table of contents and cover pages (relevant documents) from a data source (e.g. Excel) and already-generated PDF files.

## Description
This repository holds tools and scripts to enable you to
seamlessly generate:

- A Table of Contents PDF page (with anchor-links)
- Cover pages for each attachment
- Merge those into a single professional-looking PDF
- Include optional PDF outline (bookmarks)

Optional steps like OCRmyPDF or link notations can be annexed via free software.

## Folder Structure

This repo has the following folders/files:

- `input-files/` - Directory for your incoming Excel data and original attachment PDFs
- `output-files/` - Directory for the final products (merged PDFs, table of contents, etc.)
- `prompts/` - Prompt files and other management files
- `generate_toc_coverpage.py` - Script to generate table of contents and cover pages
- `merge_pdfs.py` - Script to merge the TOC/cover pages with attachment PDFs
- `generate_and_merge.sh` - Shell script to run both Python scripts in sequence

## Requirements & Dependencies

We use Python 3. Minimal tested with Linux Mint 22. We also rely on these libraries:

- [weasyprint](https://weasyprint.org/) - For HTML to PDF conversion
- [openpyxl](https://openpyxl.readthedocs.io/en/stable/) - For reading Excel files
- [pymupdf](https://pymupdf.readthedocs.io/en/latest/) - For PDF manipulation and merging


Depending on your system, you may also need to install native libraries like cairo, pango, libffi, etc., which are required by WeasyPrint. On Ubuntu-based systems, you can typically run the following to install the relevant system dependencies:
```bash
sudo apt-get install python3-dev python3-pip python3-setuptools python3-wheel python3-cffi libcairo2 libpango-1.0-0 libpangocairo-1.0-0 libgdk-pixbuf2.0-0 libffi-dev shared-mime-info
```

## Installation

```bash
git clone https://github.com/sbosshardt/attachments-pdf-generator.git
cd attachments-pdf-generator
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

This creates an isolated Python environment in the current directory in the folder venv.

## Usage

1. Prepare your Excel file:
   - Create a file named `input-pdfs.xlsx` in the `input-files` directory
   - Include sheet named "Attachments Prep" with the following columns:
     - Attachment Number
     - Title
     - Page count (optional)
     - Additional Remarks about File (optional)
     - Body (Description) (optional)
     - Filename Reference (required, points to the PDF file)
     - Date (time Pacific) (optional)
     - Language (optional, specifies language subfolder)
     - Exclude (optional, set to true to exclude an attachment)

2. Place your attachment PDFs:
   - Put your PDF files in the `input-files` directory or language subdirectories (e.g., `input-files/en/`)
   - Make sure the filenames match those in the "Filename Reference" column

3. Run the script:
   ```bash
   ./generate_and_merge.sh
   ```

4. View the output files in the `output-files` directory:
   - `weasyoutput.pdf` - Table of Contents and Cover Pages only
   - `merged-attachments.pdf` - Complete document with attachments

## How It Works

1. `generate_toc_coverpage.py`:
   - Reads attachment data from the Excel file
   - Generates HTML with table of contents and cover pages
   - Converts HTML to PDF using WeasyPrint

2. `merge_pdfs.py`:
   - Reads attachment data from the Excel file
   - Opens the TOC/cover pages PDF
   - Merges it with all attachment PDFs
   - Adds PDF bookmarks/outline
   - Creates final merged PDF

3. `generate_and_merge.sh`:
   - Runs both scripts in sequence with error handling

## Author
Created by Samuel Bosshardt

Licensed under MIT License or the license of your choice.

