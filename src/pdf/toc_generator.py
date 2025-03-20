#!/usr/bin/env python3
"""
Module for generating Table of Contents and cover pages.
"""

from weasyprint import HTML
import fitz
import os
from src.config.paths import OUTPUT_HTML, OUTPUT_TOC, OUTPUT_PDF, TITLE_PAGE, FOREWORD_PAGE
from src.excel.excel_reader import normalize_attachment_number, normalize_page_count

def generate_html(attachments):
    """
    Generate HTML content for the PDF.
    
    Args:
        attachments: List of dictionaries containing attachment data
        
    Returns:
        str: HTML content for the PDF
    """
    sorted_data = sorted(attachments, key=lambda x: float(normalize_attachment_number(x.get('Attachment Number', 0))))
    
    # CSS styling for the document
    css = """
        @page {
            size: 8.5in 11in;
            margin: 1in 1in 1in 1in;
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
        
        /* Special sections */
        .toc-section {
            margin-top: 1em;
            margin-bottom: 0.5em;
        }
        
        .toc-section-title {
            font-weight: bold;
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
    <div class="toc-container" id="table-of-contents">
        <h1>Table of Contents</h1>
        <table class="toc-table">
    """
    
    # Check if title page and foreword exist
    title_page_count = 0
    if os.path.exists(TITLE_PAGE):
        title_pdf = fitz.open(TITLE_PAGE)
        title_page_count = title_pdf.page_count
        title_pdf.close()
        
    foreword_exists = os.path.exists(FOREWORD_PAGE)
    foreword_page_count = 0
    if foreword_exists:
        foreword_pdf = fitz.open(FOREWORD_PAGE)
        foreword_page_count = foreword_pdf.page_count
        foreword_pdf.close()
    
    # Calculate TOC position
    toc_start_page = 1 + title_page_count + foreword_page_count
    
    # Add title page entry to TOC
    if os.path.exists(TITLE_PAGE):
        html += f"""
        <tr>
            <td class="attachment-num">
                <a class="toc-link" href="#title-page">Title Page</a>
            </td>
            <td class="attachment-title">Title Page</td>
            <td class="page-num">1</td>
        </tr>
        """
    
    # Add foreword entry to TOC
    if foreword_exists:
        foreword_page = title_page_count + 1  # Foreword comes after title page
        html += f"""
        <tr>
            <td class="attachment-num">
                <a class="toc-link" href="#foreword">Foreword</a>
            </td>
            <td class="attachment-title">Foreword</td>
            <td class="page-num">{foreword_page}</td>
        </tr>
        """
    
    # Add Table of Contents entry to itself
    toc_page = title_page_count + foreword_page_count + 1  # TOC comes after title and foreword
    html += f"""
    <tr>
        <td class="attachment-num">
            <a class="toc-link" href="#table-of-contents">Table of Contents</a>
        </td>
        <td class="attachment-title">Table of Contents</td>
        <td class="page-num">{toc_page}</td>
    </tr>
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
    
    # Add title page placeholder if it exists
    if os.path.exists(TITLE_PAGE):
        html += """
    <div class="cover-page" id="title-page">
        <!-- Title page placeholder, will be replaced with actual title page PDF -->
    </div>
    """
    
    # Add foreword placeholder if it exists
    if foreword_exists:
        html += """
    <div class="cover-page" id="foreword">
        <!-- Foreword placeholder, will be replaced with actual foreword PDF -->
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
    # Check if title page and foreword exist and get their page counts
    title_page_count = 0
    if os.path.exists(TITLE_PAGE):
        title_pdf = fitz.open(TITLE_PAGE)
        title_page_count = title_pdf.page_count
        title_pdf.close()
        
    foreword_page_count = 0
    if os.path.exists(FOREWORD_PAGE):
        foreword_pdf = fitz.open(FOREWORD_PAGE)
        foreword_page_count = foreword_pdf.page_count
        foreword_pdf.close()
    
    # Estimate TOC pages based on number of entries
    toc_entries = len(attachments) + (1 if os.path.exists(FOREWORD_PAGE) else 0)  # Add entry for foreword
    toc_pages = max(1, min(3, (toc_entries + 24) // 25))  # Estimate 25 entries per page
    
    # Each attachment has a cover page and content
    # Title page (1) + Foreword (1) + TOC (1) + Cover page (1) + offset (1) = 5
    # The "+1" offset is because the first attachment's content starts after its cover page
    start_page = 1 + title_page_count + foreword_page_count + toc_pages + 1  
    print(f"DEBUG: Calculating attachment pages starting from page {start_page}")
    print(f"DEBUG: Title={title_page_count}, Foreword={foreword_page_count}, TOC={toc_pages}")
    
    # Track page counts for each attachment
    current_page = start_page
    page_map = {}
    
    for attachment in sorted(attachments, key=lambda x: float(normalize_attachment_number(x.get('Attachment Number', 0)))):
        attachment_num = normalize_attachment_number(attachment.get('Attachment Number', ''))
        
        # Store the page number for this attachment
        page_map[attachment_num] = current_page
        print(f"DEBUG: Setting Attachment {attachment_num} to page {current_page}")
        
        # Get the number of content pages for this attachment
        page_count = normalize_page_count(attachment.get('Page count', 1))
        
        # Cover page is already counted, just add content for the next attachment
        current_page += (page_count + 1)  # Add content pages plus next cover
    
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
    Create bookmarks for the PDF.
    
    Args:
        merged_pdf (fitz.Document): The merged PDF document
        attachment_pages (dict): Mapping of attachment numbers to page indices
        attachments (list): List of attachment data dictionaries
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        # Create bookmarks based on actual page locations
        toc_bookmarks = []
        
        # Create mapping of attachment numbers to data
        attachment_map = {}
        for attachment in attachments:
            att_num = normalize_attachment_number(attachment.get('Attachment Number', ''))
            attachment_map[att_num] = attachment
        
        # Find TOC page and other special pages
        toc_page_idx = -1
        title_page_idx = 0  # Title page is always page 0 (first page)
        foreword_page_idx = 1  # Foreword is always page 1 (second page)
        
        # Find the Table of Contents
        for page_idx in range(merged_pdf.page_count):
            page = merged_pdf[page_idx]
            text = page.get_text()
            if "Table of Contents" in text:
                toc_page_idx = page_idx
                print(f"Found TOC on page {toc_page_idx+1}")
                break
        
        # Hard-code the page indices for the special pages to avoid any issues
        title_page_exists = os.path.exists(TITLE_PAGE)
        foreword_exists = os.path.exists(FOREWORD_PAGE)
        
        # Add title page bookmark
        if title_page_exists:
            toc_bookmarks.append([1, "Title Page", title_page_idx])
            print(f"Adding bookmark: Title Page -> page {title_page_idx+1}")
        
        # Add foreword bookmark
        if foreword_exists:
            toc_bookmarks.append([1, "Foreword", foreword_page_idx])
            print(f"Adding bookmark: Foreword -> page {foreword_page_idx+1}")
        
        # Add Table of Contents bookmark
        if toc_page_idx >= 0:
            toc_bookmarks.append([1, "Table of Contents", toc_page_idx])
            print(f"Adding bookmark: Table of Contents -> page {toc_page_idx+1}")
        
        # Add bookmarks for each attachment
        for attachment_num, page_idx in sorted(attachment_pages.items(), key=lambda x: float(x[0])):
            attachment = attachment_map.get(attachment_num)
            if not attachment:
                continue
            
            title = attachment.get('Title', 'Untitled')
            
            # Add bookmark
            toc_bookmarks.append([1, f"Attachment {attachment_num}: {title}", page_idx])
            print(f"Adding bookmark for Attachment {attachment_num} to page {page_idx+1}")
        
        # Set the bookmarks
        if toc_bookmarks:
            print(f"Setting {len(toc_bookmarks)} bookmarks with 1-based page positions for PDF viewers")
            merged_pdf.set_toc(toc_bookmarks)
            
            print(f"Successfully set {len(toc_bookmarks)} bookmarks:")
            for i, bookmark in enumerate(toc_bookmarks):
                level, title, page = bookmark
                print(f"  {title} -> page {page+1}")
            
            return True
        else:
            print("No bookmarks created")
            return False
    
    except Exception as e:
        print(f"Error creating bookmarks: {e}")
        import traceback
        traceback.print_exc()
        return False

def fix_pdf_links(pdf_doc, attachment_pages):
    """
    Fix links in the PDF document.
    
    Args:
        pdf_doc: PyMuPDF document
        attachment_pages: Dictionary mapping attachment numbers to page indices
    
    Returns:
        int: Number of links fixed
    """
    links_fixed = 0
    
    try:
        # First locate TOC page
        toc_page_idx = -1
        title_page_idx = 0  # Title page is always the first page
        foreword_page_idx = 1  # Foreword is always the second page
        
        for page_idx in range(pdf_doc.page_count):
            page = pdf_doc[page_idx]
            text = page.get_text()
            if "Table of Contents" in text:
                toc_page_idx = page_idx
                print(f"Found TOC page at page {page_idx+1}")
                break
        
        # Fix links starting from TOC page
        if toc_page_idx >= 0:
            toc_page = pdf_doc[toc_page_idx]
            links = toc_page.get_links()
            
            for link in links:
                # Fix Title Page link
                if 'uri' in link and link['uri'] == '#title-page':
                    # Create a new goto link to the first page
                    new_link = {
                        'kind': fitz.LINK_GOTO,
                        'from': link['from'],
                        'page': title_page_idx,  # Title page is always page 0 (the first page)
                        'to': fitz.Point(0, 0),
                        'zoom': 0
                    }
                    # Remove old link and add new one
                    toc_page.delete_link(link)
                    toc_page.insert_link(new_link)
                    links_fixed += 1
                    print(f"Fixed link to Title Page -> page {title_page_idx+1}")
                    
                # Fix Foreword link
                elif 'uri' in link and link['uri'] == '#foreword':
                    # Create a new goto link
                    new_link = {
                        'kind': fitz.LINK_GOTO,
                        'from': link['from'],
                        'page': foreword_page_idx,  # Foreword is always page 1 (the second page)
                        'to': fitz.Point(0, 0),
                        'zoom': 0
                    }
                    # Remove old link and add new one
                    toc_page.delete_link(link)
                    toc_page.insert_link(new_link)
                    links_fixed += 1
                    print(f"Fixed link to Foreword -> page {foreword_page_idx+1}")
                    
                # Fix Table of Contents link (self-reference)
                elif 'uri' in link and link['uri'] == '#table-of-contents':
                    # Create a new goto link
                    new_link = {
                        'kind': fitz.LINK_GOTO,
                        'from': link['from'],
                        'page': toc_page_idx,
                        'to': fitz.Point(0, 0),
                        'zoom': 0
                    }
                    # Remove old link and add new one
                    toc_page.delete_link(link)
                    toc_page.insert_link(new_link)
                    links_fixed += 1
                    print(f"Fixed link to Table of Contents -> page {toc_page_idx+1}")
                    
                # Fix attachment links
                elif 'uri' in link and link['uri'].startswith('#cover-'):
                    attachment_id = link['uri'][7:]  # Remove '#cover-'
                    
                    # If we have a page number for this attachment
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
                        print(f"Fixed link to Attachment {attachment_id} -> page {target_page+1}")
            
            # Check other pages for links too
            for page_num in range(pdf_doc.page_count):
                if page_num == toc_page_idx:
                    continue  # Skip the TOC page, already processed
                    
                page = pdf_doc[page_num]
                links = page.get_links()
                
                for link in links:
                    # If this is a named destination link
                    if 'uri' in link and link['uri'].startswith('#cover-'):
                        attachment_id = link['uri'][7:]  # Remove '#cover-'
                        
                        # If we have a page number for this attachment
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
    Generate the Table of Contents and cover pages PDF.
    
    Args:
        attachments: List of dictionaries containing attachment data
    
    Returns:
        str: Path to the final PDF
    """
    # Estimate page counts and ordering
    total_content_pages = 0
    attachment_pages = {}
    
    # Count attachment pages
    for attachment in attachments:
        attachment_num = normalize_attachment_number(attachment.get('Attachment Number', ''))
        page_count = normalize_page_count(attachment.get('Page count', 1))
        
        attachment_pages[attachment_num] = page_count
        total_content_pages += page_count
        print(f"Attachment {attachment_num}: {page_count} page(s)")
    
    # Calculate total pages including cover pages
    total_attachment_pages = total_content_pages + len(attachments)  # Content + cover pages
    
    # Check if title page exists
    title_page_count = 0
    if os.path.exists(TITLE_PAGE):
        title_pdf = fitz.open(TITLE_PAGE)
        title_page_count = title_pdf.page_count
        title_pdf.close()
    
    # Check if foreword exists
    foreword_page_count = 0
    if os.path.exists(FOREWORD_PAGE):
        foreword_pdf = fitz.open(FOREWORD_PAGE)
        foreword_page_count = foreword_pdf.page_count
        foreword_pdf.close()
    
    # Estimate TOC pages based on number of entries
    toc_entries = len(attachments)
    toc_pages = max(1, min(3, (toc_entries + 24) // 25))  # Estimate 25 entries per page
    
    # Calculate starting page number for the first attachment
    start_page = 1 + title_page_count + foreword_page_count + toc_pages  # Title + Foreword + TOC pages
    
    # Calculate page numbers for displaying on cover pages
    current_page = start_page
    page_map = {}
    
    # Generate page map for TOC display
    for attachment in sorted(attachments, key=lambda x: float(normalize_attachment_number(x.get('Attachment Number', 0)))):
        attachment_num = normalize_attachment_number(attachment.get('Attachment Number', ''))
        page_map[attachment_num] = current_page
        
        # Get page count for this attachment
        page_count = attachment_pages.get(attachment_num, 1)
        
        # Move to next cover page position
        current_page += page_count + 1  # Current content + next cover
    
    # Generate HTML for TOC and cover pages
    html_content = generate_html(attachments)
    
    # Save HTML for debugging
    with open(OUTPUT_HTML, 'w') as f:
        f.write(html_content)
    print(f"Saved HTML to {OUTPUT_HTML} for debugging")
    
    # Convert HTML to PDF
    print(f"Generating TOC and cover pages: {OUTPUT_TOC}")
    HTML(string=html_content).write_pdf(OUTPUT_TOC)
    
    # Add title page if it exists
    final_pdf = add_title_page_to_toc()
    
    # Map the actual cover page positions in the final PDF
    attachment_pages = map_cover_pages()
    
    # Create bookmarks
    create_bookmarks(final_pdf, attachment_pages, attachments)
    
    # Fix links in the PDF
    links_fixed = fix_pdf_links(final_pdf, attachment_pages)
    if links_fixed > 0:
        print(f"Fixed {links_fixed} links in the PDF")
        final_pdf.save(OUTPUT_PDF)
    
    final_pdf.close()
    
    return OUTPUT_PDF

def add_title_page_to_toc():
    """
    Add the title page to the Table of Contents PDF.
    
    Returns:
        fitz.Document: The merged PDF with title page and TOC
    """
    if not os.path.exists(OUTPUT_TOC):
        raise FileNotFoundError(f"TOC PDF not found: {OUTPUT_TOC}")
    
    # Check if title page exists
    if os.path.exists(TITLE_PAGE):
        print(f"Adding title page from: {TITLE_PAGE}")
        
        # Create a new PDF with title page followed by TOC and cover pages
        merged_pdf = fitz.open()
        
        # Add the title page
        title_pdf = fitz.open(TITLE_PAGE)
        title_page_count = title_pdf.page_count
        merged_pdf.insert_pdf(title_pdf)
        print(f"Title page has {title_page_count} page(s)")
        title_pdf.close()
        
        # Add the foreword page if it exists
        foreword_page_count = 0
        if os.path.exists(FOREWORD_PAGE):
            print(f"Adding foreword from: {FOREWORD_PAGE}")
            foreword_pdf = fitz.open(FOREWORD_PAGE)
            foreword_page_count = foreword_pdf.page_count
            merged_pdf.insert_pdf(foreword_pdf)
            print(f"Foreword has {foreword_page_count} page(s)")
            foreword_pdf.close()
        
        # Add the TOC and cover pages
        toc_pdf = fitz.open(OUTPUT_TOC)
        
        # Count actual pages with content (no blank pages)
        content_pages = 0
        blank_pages = []
        for i in range(toc_pdf.page_count):
            page = toc_pdf[i]
            text = page.get_text().strip()
            if not text:
                blank_pages.append(i)
                print(f"WARNING: Blank page detected in TOC PDF at index {i}")
            else:
                content_pages += 1
        
        # Insert only non-blank pages
        for i in range(toc_pdf.page_count):
            if i not in blank_pages:
                merged_pdf.insert_pdf(toc_pdf, from_page=i, to_page=i)
        
        print(f"Added {content_pages} content pages from TOC PDF (skipped {len(blank_pages)} blank pages)")
        toc_pdf.close()
        
        # Save merged PDF
        merged_pdf.save(OUTPUT_PDF)
        
        return merged_pdf
    else:
        # If no title page, just use the TOC PDF as the output
        print(f"Title page not found at {TITLE_PAGE}, using TOC only")
        
        # Check for and remove blank pages
        toc_pdf = fitz.open(OUTPUT_TOC)
        content_pages = 0
        blank_pages = []
        for i in range(toc_pdf.page_count):
            page = toc_pdf[i]
            text = page.get_text().strip()
            if not text:
                blank_pages.append(i)
                print(f"WARNING: Blank page detected in TOC PDF at index {i}")
            else:
                content_pages += 1
        
        if blank_pages:
            print(f"Removing {len(blank_pages)} blank pages from TOC PDF")
            non_blank_pdf = fitz.open()
            for i in range(toc_pdf.page_count):
                if i not in blank_pages:
                    non_blank_pdf.insert_pdf(toc_pdf, from_page=i, to_page=i)
            non_blank_pdf.save(OUTPUT_PDF)
            non_blank_pdf.close()
            toc_pdf.close()
        else:
            # Just copy the file if no blank pages
            import shutil
            shutil.copy(OUTPUT_TOC, OUTPUT_PDF)
            toc_pdf.close()
            
        print(f"PDF generated at: {OUTPUT_PDF}")
        return fitz.open(OUTPUT_PDF)

def map_cover_pages():
    """
    Map the actual locations of cover pages in the final PDF.
    
    Returns:
        dict: Mapping of attachment numbers to page indices
    """
    pdf_doc = fitz.open(OUTPUT_PDF)
    actual_cover_pages = {}
    blank_pages_found = 0
    
    print("Mapping actual cover page locations...")
    for page_idx in range(pdf_doc.page_count):
        page = pdf_doc[page_idx]
        text = page.get_text().strip()
        
        # Detect blank pages
        if not text:
            blank_pages_found += 1
            print(f"WARNING: Found blank page at index {page_idx} (page {page_idx+1})")
            continue
        
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
    
    # Remove blank pages if found
    if blank_pages_found > 0:
        print(f"WARNING: Detected {blank_pages_found} blank pages in the document")
        
        # We'll handle this by properly mapping pages in subsequent steps
        # and ensuring bookmarks point to the correct pages
    
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