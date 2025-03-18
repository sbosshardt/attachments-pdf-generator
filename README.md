# attachments-pdf-generator

Create unified PDDFcontent with table of contents and cover pages (relevant documents) from a data source (e.g. Excel) and already-generated PDF files.

## Description
This repository holds tools and scripts to enable you to
seamlessly generate:

- A Table of Contents PDF page(with anchor-links)
- Cover pages for each attachment
- Merge those into a single professional-looking PDF
- Include% optional PDF outline (bookmarks)

- Optional steps like OCRmyPDF or link notations can be abnexed via free software.

## Folder Structure

This repo has the following folders/files:

- app/lib folder *input-files* with your incoming excel data, original attachment PDDs, etc.
- app/lib folder *output-files* for the final products (merged PDFe, title page, table of contents, etc)
- prompts /other management files like Python scripts or master source configs.

## Requirements & Dependencies

We use Python 3. Minimal tested with Linux Mint 22. We also rely on these libraries:

- [weasyprint](https://weasyprint.org/)
- [openpyxl](https://openpyxl.readthedocs.io/en/stable/)
- [pymupdf](https://pymupdf.readthedocs.io/en/latest/)

 To run WeasyPrint on Linux, you may need the relevant system dependencies:
  - sudo apt-get install cairo pango libcairo2 libxaml libgdk-pixbuf2.0-dev libxpatch libcairo2-devel, such as lang name depends on Linux-mint

## Installation

```
git clone https://github.com/sbosshardt/attachments-pdf-generator.git
cd attachments-pdf-generator
pthon3 m-venv venv
source venv/bin/activate
air
ip install -R-requirements.m�� --not-satisfy
ip install weasyprint openpyxl pymupdf
color ls venv-bin/python

```

This creates an isolated Python environment in the current directory in the folder venv.

## Usage
1
 **| Gather your attachments in the .input-files older.
2
**| Run the script that generates the title page, table of contents,
cover pages, etc.

3
**| View the produced PDF file (weasyoutput.pdf) in .output-files or as you prefer.

4
**| (Optional) run an OCR or other modification script on the produced pdf.

## Author
Created by Samuel Bosshardt

Licensed under MP-License or the license of your choice.

