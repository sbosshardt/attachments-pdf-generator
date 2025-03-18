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

- app/lib folder *input-files* with your incoming excel data, original attachment PDFs, etc.
- app/lib folder *output-files* for the final products (merged PDFs, title page, table of contents, etc)
- prompts /other management files like Python scripts or master source configs.

## Requirements & Dependencies

We use Python 3. Minimal tested with Linux Mint 22. We also rely on these libraries:

- [weasyprint](https://weasyprint.org/)
- [openpyxl](https://openpyxl.readthedocs.io/en/stable/)
- [pymupdf](https://pymupdf.readthedocs.io/en/latest/)


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

1. Gather your attachments in the input-files folder.
2. Run the script that generates the title page, table of contents, cover pages, etc.
3. View the produced PDF file (weasyoutput.pdf) in output-files or as you prefer.
4. (Optional) run an OCR or other modification script on the produced pdf.

## Author
Created by Samuel Bosshardt

Licensed under MIT License or the license of your choice.

