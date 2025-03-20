#!/usr/bin/env python3
"""
Module for merging PDFs with attachments.
"""

import os
import fitz
from src.config.paths import OUTPUT_PDF, MERGED_PDF
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
    Locate the Table of Contents page in the PDF.
    
    Args:
        pdf_doc: PyMuPDF document
        
    Returns:
        tuple: (page index, dict of TOC links)
    """
    toc_page_index = -1
    toc_links = {}
    
    for page_num in range(pdf_doc.page_count):
        page = pdf_doc[page_num]
        text = page.get_text()
        if "Table of Contents" in text:
            toc_page_index = page_num
            print(f"Found TOC on page {page_num+1}")
            # Extract all links from TOC
            links = page.get_links()
            for link in links:
                if 'uri' in link and link['uri'].startswith('#cover-'):
                    attachment_id = link['uri'][7:]  # Remove '#cover-'
                    toc_links[attachment_id] = link
                    print(f"Found TOC link to Attachment {attachment_id}")
            break
    
    return toc_page_index, toc_links

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

def create_bookmarks(merged_pdf, toc_page_index, bookmark_positions, attachment_map, page_mapping):
    """
    Create bookmarks/outline for the merged PDF.
    
    Args:
        merged_pdf: PyMuPDF document
        toc_page_index: Page index of the TOC
        bookmark_positions: Dictionary mapping attachment numbers to page indices
        attachment_map: Dictionary mapping attachment numbers to attachment data
        page_mapping: Dictionary mapping original page indices to new page indices
        
    Returns:
        None
    """
    toc = []
    
    # First add TOC bookmark if we found the TOC page
    if toc_page_index >= 0 and toc_page_index in page_mapping:
        toc.append([1, "Table of Contents", page_mapping[toc_page_index]])
    
    # Add bookmarks for each attachment that we found a position for
    for attachment_num in sorted(bookmark_positions.keys(), key=lambda x: float(x) if x.replace('.', '', 1).isdigit() else float('inf')):
        attachment = attachment_map.get(attachment_num)
        if not attachment:
            continue
            
        title = attachment.get('Title', 'Untitled')
        
        # Add bookmark for this attachment's cover page
        toc.append([1, f"Attachment {attachment_num}: {title}", bookmark_positions[attachment_num]])
    
    # Set the table of contents/bookmarks
    if toc:
        print(f"Setting {len(toc)} bookmarks in PDF")
        try:
            merged_pdf.set_toc(toc)
        except Exception as e:
            print(f"Warning: Failed to set bookmarks: {e}")

def fix_links(merged_pdf, bookmark_positions):
    """
    Fix all internal links in the merged PDF.
    
    Args:
        merged_pdf: PyMuPDF document
        bookmark_positions: Dictionary mapping attachment numbers to page indices
        
    Returns:
        int: Number of links fixed
    """
    links_fixed = 0
    
    try:
        # First, focus on the TOC page for fixing links
        toc_page_pos = -1
        for i in range(merged_pdf.page_count):
            text = merged_pdf[i].get_text()
            if "Table of Contents" in text:
                toc_page_pos = i
                break
        
        if toc_page_pos >= 0:
            print(f"Fixing links on TOC page (page {toc_page_pos+1})")
            toc_page = merged_pdf[toc_page_pos]
            links = toc_page.get_links()
            print(f"Found {len(links)} links on TOC page")
            
            # Fix each link
            for link in links:
                if 'uri' in link and link['uri'].startswith('#cover-'):
                    attachment_id = link['uri'][7:]  # Remove '#cover-'
                    
                    # If we know the position of this attachment's cover page
                    if attachment_id in bookmark_positions:
                        target_page = bookmark_positions[attachment_id]
                        
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
        
        # Now check all pages for any other internal links that need fixing
        for page_num in range(merged_pdf.page_count):
            if page_num == toc_page_pos:
                continue  # Skip TOC page as we've already processed it
                
            page = merged_pdf[page_num]
            links = page.get_links()
            
            for link in links:
                # If this is a named destination link
                if 'uri' in link and link['uri'].startswith('#cover-'):
                    attachment_id = link['uri'][7:]  # Remove '#cover-'
                    
                    # If we know the position of this attachment's cover page
                    if attachment_id in bookmark_positions:
                        target_page = bookmark_positions[attachment_id]
                        
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

def merge_pdfs(attachments):
    """
    Merge the TOC PDF with the attachment PDFs, placing each attachment after its cover page.
    
    Args:
        attachments (list): List of attachment data dictionaries
    """
    print(f"Merging PDFs into: {MERGED_PDF}")
    
    # Check if TOC PDF exists
    if not os.path.exists(OUTPUT_PDF):
        raise FileNotFoundError(f"Table of contents PDF not found: {OUTPUT_PDF}")
    
    # Load the TOC and cover pages PDF
    toc_pdf = fitz.open(OUTPUT_PDF)
    print(f"TOC PDF has {len(toc_pdf)} pages")
    
    # Create a new document for the final merged PDF
    merged_pdf = fitz.open()
    
    # Build a mapping of attachment numbers to their data
    attachment_map = build_attachment_map(attachments)
    
    # First, scan the TOC PDF to locate all cover pages and links
    print("Scanning TOC PDF to locate cover pages and links...")
    toc_page_index, toc_links = locate_toc_page(toc_pdf)
    print(f"Found {len(toc_links)} TOC links")
    
    # Next, identify all cover pages by text pattern
    cover_page_indices = locate_cover_pages(toc_pdf, attachment_map)
    print(f"Found {len(cover_page_indices)} cover pages")
    
    # Now process the TOC PDF and insert attachments
    page_mapping, bookmark_positions = insert_attachments(
        merged_pdf, attachments, attachment_map, cover_page_indices
    )
    
    # Create bookmarks/outline for the PDF
    create_bookmarks(merged_pdf, toc_page_index, bookmark_positions, attachment_map, page_mapping)
    
    # Fix links - update all links in the TOC to point to the correct pages
    links_fixed = fix_links(merged_pdf, bookmark_positions)
    
    # Save the merged PDF
    print(f"Saving merged PDF with {links_fixed} fixed links")
    merged_pdf.save(MERGED_PDF)
    final_page_count = len(merged_pdf)
    merged_pdf.close()
    
    print(f"Merged PDF created at: {MERGED_PDF} ({final_page_count} pages)") 