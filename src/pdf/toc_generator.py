#!/usr/bin/env python3
"""
Module for generating Table of Contents and cover pages.
"""

from weasyprint import HTML
import fitz
import os
from src.config.paths import OUTPUT_HTML, OUTPUT_TOC, OUTPUT_PDF, TITLE_PAGE
from src.excel.excel_reader import normalize_attachment_number, normalize_page_count

def generate_html(attachments):
    """
    Generate HTML content for the PDF.
    
    Args:
        attachments: List of dictionaries containing attachment data
        
    Returns:
        str: HTML content for the PDF
    """
    sorted_data = sorted(attachments, key=lambda x: float(x.get('Attachment Number', 0)))
    
    # CSS styling for the document
    css = """
        @page {
            size: 8.5in 11in;
            margin: 1in 1in 1in 1in;
            @bottom-center {
                content: counter(page);
                font-family: Arial, sans-serif;
                font-size: 10pt;
            }
        }
        
        @page :first {
            @bottom-center {
                content: '';  /* No page number on title page */
            }
        }
        
        @page toc-page {
            @bottom-center {
                content: counter(page);
                font-family: Arial, sans-serif;
                font-size: 10pt;
            }
        }
        
        body {
            font-family: Arial, sans-serif;
            font-size: 12pt;
            line-height: 1.5;
            counter-reset: page 1;
        }
        
        h1 {
            font-size: 18pt;
            font-weight: bold;
            text-align: center;
            margin-top: 0.5in;
            margin-bottom: 0.25in;
        }
        
        .toc-container {
            page: toc-page;
            break-after: page;
            width: 100%;
        }
        
        /* TOC Table Styling */
        table.toc-table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 0.25in;
            margin-bottom: 0.25in;
            table-layout: fixed;
        }
        
        /* Column widths */
        table.toc-table td.attachment-num {
            width: 1.5in;
            vertical-align: top;
            padding-bottom: 0.15in;
        }
        
        table.toc-table td.attachment-title {
            vertical-align: top;
            padding-bottom: 0.15in;
        }
        
        table.toc-table td.page-num {
            width: 0.5in;
            text-align: right;
            vertical-align: top;
            padding-bottom: 0.15in;
        }
        
        /* Leader line */
        table.toc-table td.attachment-title {
            position: relative;
        }
        
        table.toc-table td.attachment-title::after {
            content: "";
            position: absolute;
            bottom: 0.5em;
            left: 0;
            width: 100%;
            border-bottom: 1px dotted black;
            z-index: -1;
        }
        
        /* Link styling */
        a.toc-link {
            color: blue;
            text-decoration: underline;
            white-space: nowrap;
        }
        
        /* Cover page styling */
        .cover-page {
            page-break-before: always;
            page-break-after: always;
            text-align: center;
        }
        
        .cover-title {
            font-size: 16pt;
            font-weight: bold;
            margin-top: 3in;
            margin-bottom: 0.25in;
        }
        
        .cover-number {
            font-size: 14pt;
            margin-bottom: 0.25in;
        }
        
        .cover-info {
            font-size: 12pt;
            margin-bottom: 0.15in;
        }
    """
    
    # Start building the HTML document
    html = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Table of Contents and Cover Pages</title>
    <style>{css}</style>
</head>
<body>
    """
    
    # Add Table of Contents
    html += """
    <div class="toc-container">
        <h1>Table of Contents</h1>
        <table class="toc-table">
    """
    
    # Calculate page numbers
    page_map = calculate_page_map(sorted_data)
    
    # Add a TOC entry for each attachment
    for attachment in sorted_data:
        attachment_num = normalize_attachment_number(attachment.get('Attachment Number', ''))
        title = attachment.get('Title', 'Untitled')
        page_num = page_map.get(attachment_num, 0)
        
        # Add the TOC entry with table structure
        html += f"""
        <tr id="toc-entry-{attachment_num}">
            <td class="attachment-num">
                <a class="toc-link" href="#cover-{attachment_num}">Attachment {attachment_num}</a>
            </td>
            <td class="attachment-title">{title}</td>
            <td class="page-num">{page_num}</td>
        </tr>
        """
    
    # Close the TOC section
    html += """
        </table>
    </div>
    """
    
    # Add cover pages for each attachment
    for attachment in sorted_data:
        attachment_num = normalize_attachment_number(attachment.get('Attachment Number', ''))
        title = attachment.get('Title', 'Untitled')
        page_number = page_map.get(attachment_num, 0)
        
        html += f"""
    <div class="cover-page" id="cover-{attachment_num}">
        <div class="cover-title">{title}</div>
        <div class="cover-number">Attachment {attachment_num}</div>
        <div class="cover-info">Page {page_number}</div>
    </div>
    """
    
    # Close the HTML document
    html += """
</body>
</html>
    """
    
    # Save the HTML to a file for debugging
    with open(OUTPUT_HTML, 'w') as f:
        f.write(html)
    print(f"Saved HTML to {OUTPUT_HTML} for debugging")
    
    return html

def calculate_page_map(attachments):
    """
    Calculate the page number for each attachment.
    
    Args:
        attachments: List of dictionaries containing attachment data
        
    Returns:
        dict: Mapping from attachment number to page number
    """
    # Estimate TOC pages based on number of entries
    toc_entries = len(attachments)
    toc_pages = max(1, min(3, (toc_entries + 24) // 25))  # Estimate 25 entries per page
    
    # Start page is calculated based on Title page + TOC pages + first cover
    start_page = 1 + toc_pages + 1  # Title page + TOC pages + first cover page
    
    # Track page counts for each attachment
    current_page = start_page
    page_map = {}
    
    for attachment in attachments:
        attachment_num = normalize_attachment_number(attachment.get('Attachment Number', ''))
        
        # Store the page number for this attachment
        page_map[attachment_num] = current_page
        
        # Get the number of content pages for this attachment
        page_count = normalize_page_count(attachment.get('Page count', 1))
        
        # Move to the next cover page
        current_page += page_count + 1  # Add content pages plus next cover
    
    return page_map

def add_manual_toc_links(pdf_doc):
    """
    Adds manual links to the TOC if none were found.
    
    Args:
        pdf_doc: PyMuPDF document object
        
    Returns:
        bool: True if links were added, False otherwise
    """
    links_found = 0
    for page_num in range(pdf_doc.page_count):
        page = pdf_doc[page_num]
        links = page.get_links()
        if links:
            links_found += len(links)
    
    if links_found > 0:
        return False  # Links already exist
    
    print("WARNING: No links found in TOC PDF. Adding links manually...")
    
    # Find TOC page 
    toc_page_idx = -1
    for page_idx in range(pdf_doc.page_count):
        page = pdf_doc[page_idx]
        text = page.get_text()
        if "Table of Contents" in text:
            toc_page_idx = page_idx
            print(f"Found TOC on page {page_idx+1}")
            break
    
    if toc_page_idx < 0:
        return False  # No TOC page found
    
    # Extract attachment numbers from page IDs
    attachment_nums = []
    for page_idx in range(pdf_doc.page_count):
        if page_idx == toc_page_idx:
            continue  # Skip TOC page
        
        page = pdf_doc[page_idx]
        text = page.get_text()
        
        if "Attachment " in text and "Page " in text:
            parts = text.split("Attachment ")
            if len(parts) > 1:
                try:
                    att_num = parts[1].split()[0].rstrip(":")
                    attachment_nums.append(att_num)
                except Exception:
                    pass
    
    # For each attachment, find text "Attachment X" on TOC page and create a link
    toc_page = pdf_doc[toc_page_idx]
    links_added = 0
    
    for att_num in attachment_nums:
        # Find text on the page
        search_text = f"Attachment {att_num}"
        rects = toc_page.search_for(search_text)
        
        if rects:
            # The first occurrence is the one in the TOC
            rect = rects[0]
            
            # Find the target page (simplistic approach)
            target_page = -1
            for page_idx in range(pdf_doc.page_count):
                if page_idx == toc_page_idx:
                    continue
                
                page = pdf_doc[page_idx]
                text = page.get_text()
                if search_text in text and "Page " in text:
                    target_page = page_idx
                    break
            
            if target_page >= 0:
                # Create a new internal link
                link = {
                    "kind": fitz.LINK_GOTO,
                    "from": rect,  # rectangle containing found text
                    "page": target_page,
                    "to": fitz.Point(0, 0),  # top of target page
                    "zoom": 0,  # default zoom
                }
                
                toc_page.insert_link(link)
                links_added += 1
                print(f"Added link for {search_text} pointing to page {target_page+1}")
    
    if links_added > 0:
        print(f"Added {links_added} links to the TOC")
        return True
    
    return False

def create_bookmarks(merged_pdf, attachment_pages, attachments):
    """
    Create bookmarks/outline for the PDF.
    
    Args:
        merged_pdf: PyMuPDF document
        attachment_pages: Dictionary mapping attachment numbers to page indices
        attachments: List of attachment data dictionaries
        
    Returns:
        None
    """
    bookmarks = []
    
    # Add a bookmark for Table of Contents
    toc_page_idx = -1
    for page_idx in range(merged_pdf.page_count):
        page = merged_pdf[page_idx]
        text = page.get_text()
        if "Table of Contents" in text:
            toc_page_idx = page_idx
            bookmarks.append([1, "Table of Contents", page_idx])
            break
    
    # Create a mapping of attachment numbers to their data
    attachment_map = {}
    for attachment in attachments:
        att_num = normalize_attachment_number(attachment.get('Attachment Number', ''))
        attachment_map[att_num] = attachment
    
    # Add bookmarks for each attachment
    for att_num, page_idx in sorted(attachment_pages.items(), key=lambda x: float(x[0]) if x[0].replace('.', '', 1).isdigit() else float('inf')):
        attachment = attachment_map.get(att_num)
        if not attachment:
            continue
            
        title = attachment.get('Title', 'Untitled')
        bookmarks.append([1, f"Attachment {att_num}: {title}", page_idx])
    
    # Set the bookmarks
    if bookmarks:
        try:
            print(f"Adding {len(bookmarks)} bookmarks to PDF")
            merged_pdf.set_toc(bookmarks)
        except Exception as e:
            print(f"Warning: Failed to add bookmarks: {e}")

def fix_pdf_links(pdf_doc, attachment_pages):
    """
    Fix links in the PDF to point to the correct pages.
    
    Args:
        pdf_doc: PyMuPDF document
        attachment_pages: Dictionary mapping attachment numbers to page indices
        
    Returns:
        int: Number of links fixed
    """
    links_fixed = 0
    
    try:
        # Find the TOC page
        toc_page_idx = -1
        for i in range(pdf_doc.page_count):
            page = pdf_doc[i]
            text = page.get_text()
            if "Table of Contents" in text:
                toc_page_idx = i
                break
        
        if toc_page_idx < 0:
            return 0  # No TOC page found
        
        # First fix links on the TOC page
        toc_page = pdf_doc[toc_page_idx]
        links = toc_page.get_links()
        
        for link in links:
            if 'uri' in link and link['uri'].startswith('#cover-'):
                attachment_id = link['uri'][7:]  # Remove '#cover-'
                
                if attachment_id in attachment_pages:
                    target_page = attachment_pages[attachment_id]
                    
                    # Create a new goto link
                    new_link = {
                        'kind': fitz.LINK_GOTO,
                        'from': link['from'],
                        'page': target_page,
                        'to': fitz.Point(0, 0),
                        'zoom': 0
                    }
                    
                    # Remove old link and add new one
                    toc_page.delete_link(link)
                    toc_page.insert_link(new_link)
                    links_fixed += 1
        
        # Now fix links on other pages
        for page_num in range(pdf_doc.page_count):
            if page_num == toc_page_idx:
                continue  # Skip TOC page
                
            page = pdf_doc[page_num]
            links = page.get_links()
            
            for link in links:
                if 'uri' in link and link['uri'].startswith('#cover-'):
                    attachment_id = link['uri'][7:]
                    
                    if attachment_id in attachment_pages:
                        target_page = attachment_pages[attachment_id]
                        
                        # Create a new goto link
                        new_link = {
                            'kind': fitz.LINK_GOTO,
                            'from': link['from'],
                            'page': target_page,
                            'to': fitz.Point(0, 0),
                            'zoom': 0
                        }
                        
                        # Remove old link and add new one
                        page.delete_link(link)
                        page.insert_link(new_link)
                        links_fixed += 1
                        
    except Exception as e:
        print(f"Warning: Error fixing links: {e}")
        import traceback
        traceback.print_exc()
    
    return links_fixed

def generate_toc_pdf(attachments):
    """
    Generate the TOC and cover pages PDF.
    
    Args:
        attachments: List of attachment data dictionaries
        
    Returns:
        dict: Mapping of attachment numbers to actual page indices
    """
    # Generate HTML content
    html_content = generate_html(attachments)
    
    # Convert HTML to PDF using WeasyPrint
    print(f"Generating TOC and cover pages: {OUTPUT_TOC}")
    HTML(string=html_content).write_pdf(OUTPUT_TOC)
    
    # Check the generated PDF for links
    toc_doc = fitz.open(OUTPUT_TOC)
    
    # Add manual links if needed
    if add_manual_toc_links(toc_doc):
        toc_doc.save(OUTPUT_TOC)
    
    toc_doc.close()
    
    # If title page exists, add it to the TOC PDF
    if os.path.exists(TITLE_PAGE):
        print(f"Adding title page from: {TITLE_PAGE}")
        add_title_page_to_toc()
    else:
        # Just use the TOC PDF as the output
        print(f"Title page not found at {TITLE_PAGE}, using TOC only")
        import shutil
        shutil.copy(OUTPUT_TOC, OUTPUT_PDF)
    
    # Map the actual locations of cover pages
    return map_cover_pages()

def add_title_page_to_toc():
    """
    Add the title page to the TOC PDF.
    
    Returns:
        None
    """
    # Create a new PDF with title page followed by TOC and cover pages
    merged_pdf = fitz.open()
    
    # Add the title page
    title_pdf = fitz.open(TITLE_PAGE)
    merged_pdf.insert_pdf(title_pdf)
    
    # Add the TOC and cover pages
    toc_pdf = fitz.open(OUTPUT_TOC)
    merged_pdf.insert_pdf(toc_pdf)
    
    # Save the merged PDF
    merged_pdf.save(OUTPUT_PDF)
    merged_pdf.close()

def map_cover_pages():
    """
    Map the actual locations of cover pages in the final PDF.
    
    Returns:
        dict: Mapping of attachment numbers to page indices
    """
    pdf_doc = fitz.open(OUTPUT_PDF)
    actual_cover_pages = {}
    
    print("Mapping actual cover page locations...")
    for page_idx in range(pdf_doc.page_count):
        page = pdf_doc[page_idx]
        text = page.get_text()
        
        # Skip TOC page
        if "Table of Contents" in text:
            continue
            
        # Look for text that matches "Attachment X" pattern
        if "Attachment " in text and "Page " in text:
            # Extract the attachment number
            parts = text.split("Attachment ")
            if len(parts) > 1:
                try:
                    attachment_num = parts[1].split()[0].rstrip(":")
                    actual_cover_pages[attachment_num] = page_idx
                    print(f"Found Attachment {attachment_num} on page {page_idx+1}")
                except Exception:
                    pass
    
    # Fix links in the PDF
    links_fixed = fix_pdf_links(pdf_doc, actual_cover_pages)
    
    if links_fixed > 0:
        print(f"Fixed {links_fixed} links in the document")
        # Use temporary file to avoid "save to original must be incremental" error
        import tempfile
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            pdf_doc.save(tmp.name)
            pdf_doc.close()
            import shutil
            shutil.copy(tmp.name, OUTPUT_PDF)
            os.unlink(tmp.name)
        pdf_doc = fitz.open(OUTPUT_PDF)
    
    # Add bookmarks to the PDF
    create_bookmarks(pdf_doc, actual_cover_pages, [])
    
    # Use temporary file to avoid "save to original must be incremental" error
    import tempfile
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        pdf_doc.save(tmp.name)
        pdf_doc.close()
        import shutil
        shutil.copy(tmp.name, OUTPUT_PDF)
        os.unlink(tmp.name)
    
    return actual_cover_pages 