# attachments-pdf-generator

Create unified PDF content with table of contents and cover pages (relevant documents) from a data source (e.g. Excel) and already-generated PDF files.

## Description
This repository holds tools and scripts to enable you to
seamlessly generate:

- A title page
- A Table of Contents PDF page (with anchor-links and page numbers)
- Cover pages for each attachment
- Merge those into a single professional-looking PDF with each attachment immediately following its cover page
- Include optional PDF outline (bookmarks)

Optional steps like OCRmyPDF or link notations can be annexed via free software.

## Folder Structure

This repo has the following folders/files:

- `input-files/` - Directory for your incoming Excel data and original attachment PDFs
  - `title-page.pdf` - Title page to include at the beginning of the document
  - `input-pdfs.xlsx` - Excel file with attachment information
  - Language subdirectories (e.g., `en/`) - For PDFs in different languages
- `output-files/` - Directory for the final products
  - `toc-coverpage.pdf` - Generated table of contents and cover pages only
  - `weasyoutput.pdf` - Title page + TOC + cover pages
  - `merged-attachments.pdf` - Complete document with all attachments
- `src/` - Source code organized in a modular structure
  - `config/` - Configuration constants and file paths
  - `excel/` - Excel data reading functionality
  - `pdf/` - PDF generation and merging operations
  - `utils/` - Common utility functions
  - `generate_toc.py` - Entry point for generating table of contents and cover pages
  - `merge_pdfs.py` - Entry point for merging with attachment PDFs
- `prompts/` - Prompt files and other management files
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
   - Create a title page PDF named `title-page.pdf` in the `input-files` directory
   - Put your PDF files in language subdirectories (e.g., `input-files/en/`)
   - Make sure the filenames match those in the "Filename Reference" column

3. Run the script:
   ```bash
   ./generate_and_merge.sh
   ```

4. View the output files in the `output-files` directory:
   - `toc-coverpage.pdf` - Table of Contents and Cover Pages only
   - `weasyoutput.pdf` - Title page + TOC + Cover Pages
   - `merged-attachments.pdf` - Complete document with attachments

## How It Works

The process is organized into modular components for better maintainability:

1. Excel Processing:
   - `src/excel/excel_reader.py` reads attachment data from the Excel file and normalizes values

2. Table of Contents & Cover Pages Generation:
   - `src/pdf/toc_generator.py` handles the generation of HTML content
   - Converts the HTML to PDF using WeasyPrint
   - Adds internal links and processes the title page if present

3. PDF Merging:
   - `src/pdf/pdf_merger.py` takes the generated TOC/cover page PDF
   - Scans for cover pages and inserts the corresponding attachment PDFs
   - Updates page numbers, internal links, and bookmarks

4. Main Entry Points:
   - `src/generate_toc.py` - Handles the TOC generation process
   - `src/merge_pdfs.py` - Manages the PDF merging process
   
5. Orchestration:
   - `generate_and_merge.sh` activates the virtual environment and runs both scripts in sequence

## Author
Created by Samuel Bosshardt with assistance from Cursor/Claude/ChatGPT.

Licensed under MIT License.

