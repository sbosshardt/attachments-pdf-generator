#!/usr/bin/env python3
"""
Module for merging PDFs with attachments.
"""

import os
import fitz
from src.config.paths import OUTPUT_PDF, MERGED_PDF, TITLE_PAGE, FOREWORD_PAGE
from src.excel.excel_reader import normalize_attachment_number

def build_attachment_map(attachments):
    """
    Build a mapping of attachment numbers to their data.
    
    Args:
        attachments (list): List of attachment data dictionaries
        
    Returns:
        dict: Dictionary mapping attachment numbers to attachment data
    """
    attachment_map = {}
    for attachment in attachments:
        attachment_num = normalize_attachment_number(attachment.get('Attachment Number', ''))
        attachment_map[attachment_num] = attachment
    return attachment_map

def locate_toc_page(pdf_doc):
    """
    Locate all Table of Contents pages in the PDF.
    
    Args:
        pdf_doc: PyMuPDF document
        
    Returns:
        tuple: (list of TOC page indices, dict of TOC links)
    """
    first_toc_page = -1
    toc_page_indices = []
    toc_links = {}
    
    # TOC is only on pages 2 and 3 (index 1 and 2)
    for page_num in range(pdf_doc.page_count):
        page = pdf_doc[page_num]
        text = page.get_text()
        
        # First page with "Table of Contents" is the main TOC page
        if "Table of Contents" in text and "Attachment " in text:
            first_toc_page = page_num
            toc_page_indices.append(page_num)
            print(f"Found main TOC page at page {page_num+1}")
            
            # Extract all links from this TOC page
            links = page.get_links()
            print(f"Found {len(links)} links on TOC page {page_num+1}")
            print(f"Links: {links}")
            for link in links:
                if 'uri' in link and link['uri'].startswith('#cover-'):
                    attachment_id = link['uri'][7:]  # Remove '#cover-'
                    toc_links[attachment_id] = link
                    print(f"Found TOC link to Attachment {attachment_id}")
        
        # The next page after the TOC main page is continuation (if it has attachment entries)
        elif first_toc_page > -1 and page_num not in toc_page_indices and "Attachment " in text:
            # Find two adjacent numbers (like "14 50") which indicates this is TOC formatting
            import re
            if re.search(r'Attachment\s+\d+\s+\d+\s*$', text, re.MULTILINE):
                toc_page_indices.append(page_num)
                print(f"Found possible TOC continuation page at page {page_num+1}")
                
                # Extract all links from this TOC page
                links = page.get_links()
                for link in links:
                    if 'uri' in link and link['uri'].startswith('#cover-'):
                        attachment_id = link['uri'][7:]  # Remove '#cover-'
                        toc_links[attachment_id] = link
                        print(f"Found TOC link to Attachment {attachment_id}")
            else:
                break
    
    print(f"Found {len(toc_page_indices)} TOC pages with {len(toc_links)} total links")
    return first_toc_page, toc_links, toc_page_indices

def locate_cover_pages(pdf_doc, attachment_map):
    """
    Locate all cover pages in the PDF.
    
    Args:
        pdf_doc: PyMuPDF document
        attachment_map: Dictionary mapping attachment numbers to attachment data
        
    Returns:
        dict: Dictionary mapping attachment numbers to page indices
    """
    cover_page_indices = {}
    
    for page_num in range(pdf_doc.page_count):
        page = pdf_doc[page_num]
        text = page.get_text()
        
        if "Table of Contents" in text:
            continue  # Skip TOC page
            
        for attachment_id in attachment_map:
            if f"Attachment {attachment_id}" in text:
                # Make sure this isn't just a mention in another cover page
                if "Page" in text and f"Attachment {attachment_id}" in text.split('\n')[:5]:
                    cover_page_indices[attachment_id] = page_num
                    print(f"Found cover page for Attachment {attachment_id} on page {page_num+1}")
                    break
    
    return cover_page_indices

def insert_attachments(merged_pdf, attachments, attachment_map, cover_page_indices):
    """
    Insert all attachment PDFs after their respective cover pages.
    
    Args:
        merged_pdf: PyMuPDF document for merged PDF
        attachments: List of attachment data dictionaries
        attachment_map: Dictionary mapping attachment numbers to attachment data
        cover_page_indices: Dictionary mapping attachment numbers to page indices
        
    Returns:
        tuple: (updated page mapping, bookmark positions)
    """
    # Keep track of mapping from original page to new page
    page_mapping = {}
    current_page = 0
    
    # Add all the TOC pages first
    toc_pdf = fitz.open(OUTPUT_PDF)
    for i in range(toc_pdf.page_count):
        merged_pdf.insert_pdf(toc_pdf, from_page=i, to_page=i)
        page_mapping[i] = current_page
        current_page += 1
    
    # Track positions for bookmarks and links
    bookmark_positions = {}
    
    # Now add each attachment right after its cover page
    for attachment_num, page_idx in sorted(cover_page_indices.items(), key=lambda x: float(x[0]) if x[0].replace('.', '', 1).isdigit() else float('inf')):
        attachment = attachment_map.get(attachment_num)
        if not attachment:
            print(f"Warning: No data found for Attachment {attachment_num}, skipping")
            continue
            
        filename = attachment.get('Filename Reference', '')
        language = attachment.get('Language', 'EN')
        
        # Skip if no filename
        if not filename:
            print(f"Warning: No filename for Attachment {attachment_num}, skipping")
            continue
        
        # We already added the cover page in the first pass
        # Find its position in the merged PDF
        cover_page_pos = page_mapping[page_idx]
        bookmark_positions[attachment_num] = cover_page_pos
        
        # Create the full path to the attachment file
        attachment_path = os.path.join('input-files', language.lower(), filename)
        
        # Check if attachment exists
        if not os.path.exists(attachment_path):
            print(f"Warning: Attachment file not found: {attachment_path}, skipping")
            continue
        
        try:
            # Calculate where to insert
            insert_pos = cover_page_pos + 1
            
            # Open the attachment PDF
            attachment_pdf = fitz.open(attachment_path)
            attachment_pages = len(attachment_pdf)
            
            print(f"Inserting {attachment_pages} pages for Attachment {attachment_num} after position {insert_pos}")
            
            # Insert after the cover page
            merged_pdf.insert_pdf(attachment_pdf, start_at=insert_pos)
            
            # Update page mapping for all subsequent pages
            for orig_page in sorted(page_mapping.keys()):
                if page_mapping[orig_page] >= insert_pos:
                    page_mapping[orig_page] += attachment_pages
            
            # Also update bookmark positions for all subsequent attachments
            for att_num in bookmark_positions:
                if bookmark_positions[att_num] >= insert_pos:
                    bookmark_positions[att_num] += attachment_pages
            
            # Update current page counter
            current_page += attachment_pages
            
            print(f"Added: Attachment {attachment_num} - {filename} ({attachment_pages} pages)")
            
        except Exception as e:
            print(f"Error adding attachment {attachment_num}: {e}")
    
    return page_mapping, bookmark_positions

def create_bookmarks(doc, toc_pdf, cover_page_info, title_page_exists=True, foreword_exists=True):
    """Create bookmarks in the merged PDF document."""
    if not doc or not toc_pdf:
        return

    # Initialize bookmarks list
    bookmarks = []
    
    # Add title page bookmark
    title_page_index = 1  # Title page is at page 1
    bookmarks.append([1, "Title Page", title_page_index])
    print(f"Adding bookmark: Title Page -> page {title_page_index}")
    
    # Add foreword bookmark if it exists
    foreword_index = 2  # Foreword starts at page 2
    if foreword_exists:
        bookmarks.append([1, "Foreword", foreword_index])
        print(f"Adding bookmark: Foreword -> page {foreword_index}")
    
    # Add TOC bookmark
    toc_index = 3
    bookmarks.append([1, "Table of Contents", toc_index])
    print(f"Adding bookmark: Table of Contents -> page {toc_index}")
    
    # Add attachment bookmarks
    for info in cover_page_info:
        if 'attachment_num' in info and 'merged_page' in info and 'title' in info:
            attachment_num = info['attachment_num']
            merged_page = info['merged_page']
            title = info['title']
            
            # Create bookmark title
            bookmark_title = f"Attachment {attachment_num}: {title}"
            
            # Create bookmark with page offset
            bookmarks.append([1, bookmark_title, merged_page+1])
            print(f"Adding bookmark for Attachment {attachment_num} to page {merged_page+1}")
    
    # Set bookmarks in the document
    if bookmarks:
        print(f"Setting {len(bookmarks)} bookmarks with adjusted page positions for PDF viewers")
        doc.set_toc(bookmarks)
        
        # Print confirmation of bookmarks
        print("Successfully set", len(bookmarks), "bookmarks:")
        for bookmark in bookmarks:
            level, title, page = bookmark
            # Display page number as 1-based for easier understanding
            print(f"  {title} -> page {page}")
    
    return bookmarks

def merge_pdfs(toc_pdf_path, attachments, output_file=MERGED_PDF):
    """
    Merge the TOC PDF and attachment PDFs into a single PDF
    
    Args:
        toc_pdf_path: Path to the TOC PDF file
        attachments: List of dictionaries with attachment information
        output_file: Path to the output PDF file
    """
    print(f"Merging PDFs into: {output_file}")
    
    # Create attachments map for easier lookups
    attachment_map = {a['Number']: a for a in attachments}
    
    # Open TOC PDF
    toc_pdf = fitz.open(toc_pdf_path)
    print(f"TOC PDF has {toc_pdf.page_count} pages")
    
    # Scan the TOC PDF to find the main TOC page, and map pages to attachments
    print("Scanning TOC PDF to locate cover pages and links...")
    title_page_found = False
    foreword_found = False
    toc_page_indices = []
    
    # Looking for title page, foreword, TOC page
    title_page_index = None
    foreword_page_index = None
    main_toc_page_index = None
    
    # Track found cover pages to avoid duplicates
    found_cover_pages = set()
    cover_page_info = []
    
    # First, check for title page and foreword
    for i in range(toc_pdf.page_count):
        page_text = toc_pdf[i].get_text()
        
        # Check for title page (usually first page)
        if i == 0:
            title_page_index = i
            title_page_found = True
            print(f"Found title page at page {i+1}")
        
        # Check for foreword (usually second page)
        if i == 1 and "foreword" in page_text.lower():
            foreword_page_index = i
            foreword_found = True
            print(f"Found foreword on page {i+1}")

    # Get the main TOC page index and the TOC page indices
    main_toc_page_index, toc_links, tp_indices = locate_toc_page(toc_pdf)
    toc_page_indices = toc_page_indices + tp_indices

    # Now look for all attachment cover pages
    for i in range(toc_pdf.page_count):
        page_text = toc_pdf[i].get_text()
        
        # Skip TOC pages - they also contain attachment references
        if i in toc_page_indices or "Table of Contents" in page_text:
            continue
            
        # Check for attachment cover pages - ensure it's a cover page, not just a mention
        if "Attachment " in page_text and "Page " in page_text:
            # Extract the attachment number from the page text
            import re
            attachment_matches = re.findall(r'Attachment\s+(\d+\.?\d*)', page_text)
            
            if attachment_matches:
                attachment_num = attachment_matches[0]
                
                # Skip if we already found this attachment's cover page
                if attachment_num in found_cover_pages:
                    continue
                    
                # Only process if this attachment number is in our list from Excel
                if attachment_num in attachment_map:
                    found_cover_pages.add(attachment_num)
                    attachment_title = attachment_map[attachment_num].get('Title', 'Untitled')
                    print(f"Found cover page for Attachment {attachment_num} on page {i+1}")
                    
                    cover_page_info.append({
                        'attachment_num': attachment_num,
                        'toc_page': i,
                        'merged_page': i,  # Initial value, will be updated as we insert pages
                        'title': attachment_title
                    })
    
    # Report cover pages found
    print(f"Found {len(cover_page_info)} cover pages")
    
    # Create a merged PDF document starting with the TOC PDF
    merged_pdf = fitz.open()
    
    # Add TOC PDF pages to the merged PDF
    merged_pdf.insert_pdf(toc_pdf)
    
    # Track page offsets for bookmarks and links
    page_offset = 0
    page_mapping = {}  # Maps toc_page -> merged_page
    bookmark_positions = {}  # Maps attachment_num -> page_index
    
    # Sort cover page info by attachment number to ensure consistent order
    cover_page_info.sort(key=lambda x: float(x['attachment_num']) if x['attachment_num'].replace('.', '', 1).isdigit() else float('inf'))
    
    # Process each attachment
    for i, info in enumerate(cover_page_info):
        attachment_num = info['attachment_num']
        attachment = attachment_map.get(attachment_num, {})
        filepath = attachment.get('FilePath')
        
        if not filepath or not os.path.exists(filepath):
            print(f"Warning: File for Attachment {attachment_num} not found: {filepath}")
            continue
        
        # Open the attachment PDF
        try:
            attachment_pdf = fitz.open(filepath)
            num_pages = attachment_pdf.page_count
            
            # Get the position to insert after (the cover page)
            insert_after = info['toc_page']
            
            # Define where to insert this attachment's pages after its cover page
            insert_position = insert_after + 1 + page_offset
            print(f"Inserting {num_pages} pages for Attachment {attachment_num} after position {insert_position}")
            
            # Insert the attachment PDF after its cover page
            merged_pdf.insert_pdf(attachment_pdf, from_page=0, to_page=num_pages-1, start_at=insert_position)
            
            # Record bookmark position
            bookmark_positions[attachment_num] = insert_after + page_offset
            
            # Update page_mapping
            toc_page = info['toc_page']
            page_mapping[toc_page] = toc_page + page_offset
            
            # Update the merged_page in cover_page_info
            info['merged_page'] = toc_page + page_offset
            
            # Update the page offset for subsequent insertions
            page_offset += num_pages
            
            # Close the attachment PDF
            attachment_pdf.close()
            
            # Print confirmation
            print(f"Added: {os.path.basename(filepath)} ({num_pages} pages)")
            
        except Exception as e:
            print(f"Error adding Attachment {attachment_num}: {e}")
    
    # Create bookmarks for the merged PDF
    create_bookmarks(merged_pdf, toc_pdf, cover_page_info, 
                    title_page_exists=(title_page_index is not None),
                    foreword_exists=(foreword_page_index is not None))
    
    # Fix links in the merged PDF - look for all TOC pages
    first_toc_page, toc_links, toc_pages_merged = locate_toc_page(merged_pdf)

    # Fix links on all TOC pages
    for toc_page in toc_pages_merged:
        page = merged_pdf[toc_page]
        links = page.get_links()
        print(f"Fixing links on TOC page (page {toc_page+1})")
        print(f"Found {len(links)} links on TOC page {toc_page+1}")
        
        # Look for potential attachment references in the TOC
        text = page.get_text()
        
        # Extract text line by line for more accurate processing
        lines = text.split('\n')
        
        # Look for potential attachment entries in various formats
        import re
        potential_attachments = set()
        
        for line in lines:
            # Search for standard format "Attachment X" or "Attachment X:"
            std_matches = re.findall(r"Attachment\s+(\d+\.?\d*)", line)
            if std_matches:
                for match in std_matches:
                    potential_attachments.add(match)
        
        if potential_attachments:
            print(f"Found potential attachments on page {toc_page+1}: {', '.join(sorted(potential_attachments))}")
            
            # Create links for attachments found on this page
            for attachment_id in potential_attachments:
                if attachment_id in bookmark_positions:
                    target_page = bookmark_positions[attachment_id]
                    
                    # Find the text in the TOC page
                    search_text = f"Attachment {attachment_id}"
                    rects = page.search_for(search_text)
                    
                    if rects:
                        # Create a new link
                        new_link = {
                            'kind': fitz.LINK_GOTO,
                            'from': rects[0],  # use the first occurrence
                            'page': target_page,
                            'to': fitz.Point(0, 0),
                            'zoom': 0
                        }
                        
                        # Add the new link
                        page.insert_link(new_link)
                        print(f"Created link for {search_text} pointing to page {target_page+1}")
    
    # Save the merged PDF
    try:
        merged_pdf.save(output_file)
        page_count = merged_pdf.page_count
        merged_pdf.close()
        
        print(f"Merged PDF created at: {output_file} ({page_count} pages)")
        
        # Print debug info about bookmarks in the final PDF
        print("\nDEBUG - Bookmarks in final PDF:")
        doc = fitz.open(output_file)
        toc = doc.get_toc()
        print(f"Total bookmarks: {len(toc)}")
        print("First few bookmarks:")
        for level, title, page in toc[:5]:
            print(f"  {title} -> page {page+1}")
        
        # Check for Foreword bookmark specifically
        for item in toc:
            if 'Foreword' in item[1]:
                level, title, page = item
                print(f"  Foreword bookmark -> page {page+1}")
                break
                
        doc.close()
    except Exception as e:
        print(f"Error saving PDF: {e}")
        import traceback
        traceback.print_exc() 