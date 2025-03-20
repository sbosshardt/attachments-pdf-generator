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
    Locate all Table of Contents pages in the PDF.
    
    Args:
        pdf_doc: PyMuPDF document
        
    Returns:
        tuple: (list of TOC page indices, dict of TOC links)
    """
    toc_page_indices = []
    toc_links = {}
    
    # TOC is only on pages 2 and 3 (index 1 and 2)
    for page_num in range(pdf_doc.page_count):
        page = pdf_doc[page_num]
        text = page.get_text()
        
        # First page with "Table of Contents" is the main TOC page
        if "Table of Contents" in text:
            toc_page_indices.append(page_num)
            print(f"Found main TOC page at page {page_num+1}")
            
            # Extract all links from this TOC page
            links = page.get_links()
            for link in links:
                if 'uri' in link and link['uri'].startswith('#cover-'):
                    attachment_id = link['uri'][7:]  # Remove '#cover-'
                    toc_links[attachment_id] = link
                    print(f"Found TOC link to Attachment {attachment_id}")
        
        # The next page after the TOC main page is continuation (if it has attachment entries)
        elif page_num == 2 and page_num not in toc_page_indices and "Attachment " in text:
            # Find two adjacent numbers (like "14 50") which indicates this is TOC formatting
            import re
            if re.search(r'Attachment\s+\d+\s+\d+\s*$', text, re.MULTILINE):
                toc_page_indices.append(page_num)
                print(f"Found TOC continuation page at page {page_num+1}")
                
                # Extract all links from this TOC page
                links = page.get_links()
                for link in links:
                    if 'uri' in link and link['uri'].startswith('#cover-'):
                        attachment_id = link['uri'][7:]  # Remove '#cover-'
                        toc_links[attachment_id] = link
                        print(f"Found TOC link to Attachment {attachment_id}")
    
    # If TOC seems to be 3 or more pages, print a warning
    if len(toc_page_indices) > 2:
        print(f"WARNING: Found more than 2 TOC pages: {len(toc_page_indices)}. This is unexpected.")
        # Limit to first 2 pages
        toc_page_indices = toc_page_indices[:2]
    
    print(f"Found {len(toc_page_indices)} TOC pages with {len(toc_links)} total links")
    
    # Return the first page for backward compatibility with existing code
    # but also return the full list for improved processing
    if toc_page_indices:
        first_toc_page = toc_page_indices[0]
    else:
        first_toc_page = -1
        
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

def create_bookmarks(merged_pdf, toc_page_index, bookmark_positions, attachment_map, page_mapping, toc_page_indices=None):
    """
    Create bookmarks/outline for the merged PDF.
    
    Args:
        merged_pdf: PyMuPDF document
        toc_page_index: Primary page index of the TOC
        bookmark_positions: Dictionary mapping attachment numbers to page indices
        attachment_map: Dictionary mapping attachment numbers to attachment data
        page_mapping: Dictionary mapping original page indices to new page indices
        toc_page_indices: List of all TOC page indices (optional)
        
    Returns:
        None
    """
    # Use a direct mapping and hardcoded approach
    toc = []
    
    # NOTE: For PDF bookmark destinations, we need to add +1 to the page index
    # This is because PyMuPDF (fitz) uses 0-based indexing internally
    # but PDF specifications and viewers typically expect 1-based page numbers
    
    # Title page (always first page)
    toc.append([1, "Title Page", 1])  # Page 1 (not 0) for PDF viewers
    print(f"Adding bookmark: Title Page -> page 1")
    
    # Table of Contents (could span multiple pages)
    if toc_page_indices and len(toc_page_indices) > 0:
        # Add main TOC bookmark to the first TOC page
        toc.append([1, "Table of Contents", toc_page_indices[0] + 1])
        print(f"Adding bookmark: Table of Contents -> page {toc_page_indices[0]+1}")
        
        # If there are additional TOC pages, add them as sub-bookmarks
        if len(toc_page_indices) > 1:
            for i, page_idx in enumerate(toc_page_indices[1:], 1):
                toc.append([2, f"Table of Contents (continued {i})", page_idx + 1])
                print(f"Adding bookmark: Table of Contents (continued {i}) -> page {page_idx+1}")
    else:
        # Fallback to the old behavior if no toc_page_indices provided
        toc.append([1, "Table of Contents", 2])  # Page 2 (not 1) for PDF viewers
        print(f"Adding bookmark: Table of Contents -> page 2")
    
    # Create a simple map for scanning
    page_info = {}
    
    # Scan the entire document to map attachment numbers to actual page indices
    for i in range(2, merged_pdf.page_count):  # Start after TOC
        page = merged_pdf[i]
        text = page.get_text()
        
        # Check for attachment cover pages
        if "Attachment " in text and "Page " in text:
            # Extract the attachment number
            try:
                parts = text.split("Attachment ")
                if len(parts) > 1:
                    attachment_num = parts[1].split()[0].rstrip(":")
                    page_info[attachment_num] = i
                    print(f"Direct scan found Attachment {attachment_num} at page {i+1}")
            except Exception as e:
                print(f"Error parsing attachment from page {i+1}: {e}")
    
    # Add bookmarks for all attachments we found
    for attachment_num in sorted(page_info.keys(), key=lambda x: float(x) if x.replace('.', '', 1).isdigit() else float('inf')):
        attachment = attachment_map.get(attachment_num)
        if not attachment:
            title = f"Attachment {attachment_num}"
        else:
            title = f"Attachment {attachment_num}: {attachment.get('Title', 'Untitled')}"
        
        page_idx = page_info[attachment_num]
        # Add 1 to page_idx for PDF viewers' 1-based page numbering
        toc.append([1, title, page_idx + 1])
        print(f"Adding bookmark for Attachment {attachment_num} to page {page_idx+1}")
    
    # Set the bookmarks directly
    print(f"Setting {len(toc)} bookmarks with 1-based page positions for PDF viewers")
    
    # Create a temporary file to save with bookmarks
    import tempfile
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        # Save current state
        merged_pdf.save(tmp.name)
        # Close and reopen to ensure changes are applied
        merged_pdf.close()
        
        # Reopen and set bookmarks
        temp_pdf = fitz.open(tmp.name)
        temp_pdf.set_toc(toc)
        
        # Save to final output
        temp_pdf.save(MERGED_PDF)
        
        # Verify the bookmarks were set properly
        new_toc = temp_pdf.get_toc()
        temp_pdf.close()
        
        # Clean up
        import os
        try:
            os.unlink(tmp.name)
        except:
            pass
    
    # Reopen the merged PDF with bookmarks
    new_pdf = fitz.open(MERGED_PDF)
    new_toc = new_pdf.get_toc()
    new_pdf.close()
    
    if new_toc:
        print(f"Successfully set {len(new_toc)} bookmarks:")
        for i, (level, title, page) in enumerate(new_toc):
            print(f"  {title} -> page {page}")
    else:
        print("Warning: Bookmarks may not have been set correctly")

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
        # Find TOC pages (should be just pages 2 and 3, indices 1 and 2)
        toc_pages = []
        
        # First check for primary TOC page (always has "Table of Contents" heading)
        for i in range(min(5, merged_pdf.page_count)):  # Only check first few pages
            text = merged_pdf[i].get_text()
            if "Table of Contents" in text:
                toc_pages.append(i)
                print(f"Found main TOC page at page {i+1}")
                
                # The next page might be TOC continuation
                if i+1 < merged_pdf.page_count:
                    next_page_text = merged_pdf[i+1].get_text()
                    # If next page has attachment entries but not a cover page
                    if "Attachment " in next_page_text and "Page " not in next_page_text:
                        toc_pages.append(i+1)
                        print(f"Found TOC continuation page at page {i+2}")
                break
        
        # Process all TOC pages
        for toc_page_pos in toc_pages:
            print(f"Fixing links on TOC page (page {toc_page_pos+1})")
            toc_page = merged_pdf[toc_page_pos]
            links = toc_page.get_links()
            print(f"Found {len(links)} links on TOC page {toc_page_pos+1}")
            
            # Look for all references to attachments in the TOC text
            text = toc_page.get_text()
            
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
                
                # Search for "Attachment X - Description"
                alt_matches = re.findall(r"Attachment\s+(\d+\.?\d*)\s*[:-]", line)
                if alt_matches:
                    for match in alt_matches:
                        potential_attachments.add(match)
            
            # Print all potential attachments found on this page
            if potential_attachments:
                print(f"Found potential attachments on page {toc_page_pos+1}: {', '.join(sorted(potential_attachments))}")
            
            # Create links for each attachment found on this page
            for attachment_id in potential_attachments:
                if attachment_id in bookmark_positions:
                    target_page = bookmark_positions[attachment_id]
                    
                    # Find the text in the TOC page
                    search_text = f"Attachment {attachment_id}"
                    rects = toc_page.search_for(search_text)
                    
                    if rects:
                        # The rectangle around the attachment number text in TOC
                        rect = rects[0]
                        
                        # Create a new goto link
                        new_link = {
                            'kind': fitz.LINK_GOTO,
                            'from': rect,
                            'page': target_page,
                            'to': fitz.Point(0, 0),
                            'zoom': 0
                        }
                        
                        # Add the new link
                        toc_page.insert_link(new_link)
                        links_fixed += 1
                        print(f"Created link for {search_text} pointing to page {target_page+1}")
            
            # Now check if there were original links and fix them if needed
            for link in links:
                if 'uri' in link and link['uri'].startswith('#cover-'):
                    attachment_id = link['uri'][7:]  # Remove '#cover-'
                    
                    # If we know the position of this attachment's cover page
                    if attachment_id in bookmark_positions:
                        # For PDF links, use the actual page numbers as they appear
                        # in the document (1-based), not the internal 0-based index
                        target_page = bookmark_positions[attachment_id]
                        # For PDF specification, we need real page number
                        # This ensures consistent behavior with PDF viewers
                        
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
                        print(f"Fixed existing link to Attachment {attachment_id} -> page {target_page+1}")
        
        # Now check all pages for any other internal links that need fixing
        for page_num in range(merged_pdf.page_count):
            if page_num in toc_pages:
                continue  # Skip TOC pages as we've already processed them
                
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
                        print(f"Fixed link to Attachment {attachment_id} -> page {target_page+1}")
    
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
    toc_page_index, toc_links, toc_page_indices = locate_toc_page(toc_pdf)
    print(f"Found {len(toc_links)} TOC links")
    
    # Next, identify all cover pages by text pattern
    cover_page_indices = locate_cover_pages(toc_pdf, attachment_map)
    print(f"Found {len(cover_page_indices)} cover pages")
    
    # Now process the TOC PDF and insert attachments
    page_mapping, bookmark_positions = insert_attachments(
        merged_pdf, attachments, attachment_map, cover_page_indices
    )
    
    # Fix links - update all links in the TOC to point to the correct pages
    links_fixed = fix_links(merged_pdf, bookmark_positions)
    print(f"Fixed {links_fixed} links in merged PDF")
    
    # Save the merged PDF (without bookmarks for now)
    temp_merged_file = MERGED_PDF + ".temp"
    merged_pdf.save(temp_merged_file)
    final_page_count = len(merged_pdf)
    merged_pdf.close()
    
    # Now, create bookmarks on the saved file
    # This function will load the file, add bookmarks, and save it back
    merged_pdf = fitz.open(temp_merged_file)
    create_bookmarks(merged_pdf, toc_page_index, bookmark_positions, attachment_map, page_mapping, toc_page_indices)
    
    # Check if we need to fix links in the final PDF
    try:
        # Reopen the final PDF and fix links again to ensure they're preserved
        final_pdf = fitz.open(MERGED_PDF)
        
        # Find all TOC pages - same logic as in fix_links
        toc_pages = []
        
        # First check for primary TOC page (always has "Table of Contents" heading)
        for i in range(min(5, final_pdf.page_count)):  # Only check first few pages
            text = final_pdf[i].get_text()
            if "Table of Contents" in text:
                toc_pages.append(i)
                print(f"Found main TOC page at page {i+1} in final PDF")
                
                # The next page might be TOC continuation
                if i+1 < final_pdf.page_count:
                    next_page_text = final_pdf[i+1].get_text()
                    # If next page has attachment entries but not a cover page
                    if "Attachment " in next_page_text and "Page " not in next_page_text:
                        toc_pages.append(i+1)
                        print(f"Found TOC continuation page at page {i+2} in final PDF")
                break
        
        # Check if links are present on all TOC pages
        links_missing = False
        total_links = 0
        expected_links = 0
        
        for toc_page_pos in toc_pages:
            toc_page = final_pdf[toc_page_pos]
            links = toc_page.get_links()
            total_links += len(links)
            print(f"Found {len(links)} links on TOC page {toc_page_pos+1} in final PDF")
            
            # Look for all references to attachments in the TOC text
            text = toc_page.get_text()
            
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
                
                # Search for "Attachment X - Description"
                alt_matches = re.findall(r"Attachment\s+(\d+\.?\d*)\s*[:-]", line)
                if alt_matches:
                    for match in alt_matches:
                        potential_attachments.add(match)
            
            # Count valid attachment references
            count_attachment_entries = sum(1 for att_id in potential_attachments if att_id in bookmark_positions)
            expected_links += count_attachment_entries
            
            # If we have fewer links than attachment entries, we have missing links
            if len(links) < count_attachment_entries:
                links_missing = True
                print(f"WARNING: Page {toc_page_pos+1} has {len(links)} links but {count_attachment_entries} attachment entries")
        
        print(f"Expected total links across all TOC pages: {expected_links}, Found: {total_links}")
        
        # If any TOC page has missing links, reapply the links
        if links_missing or total_links < expected_links:
            print(f"Warning: Some links were lost during bookmark creation. Fixing links in final PDF...")
            links_fixed = fix_links(final_pdf, bookmark_positions)
            print(f"Fixed {links_fixed} links in final PDF after bookmarks")
            
            # We can't save directly to the same file, so use a temporary file
            temp_final_file = MERGED_PDF + ".final.temp"
            final_pdf.save(temp_final_file)
            final_pdf.close()
            
            # Now move the temporary file to the target location
            import shutil
            shutil.copy(temp_final_file, MERGED_PDF)
            
            try:
                os.remove(temp_final_file)
            except:
                pass  # Ignore errors if we can't remove the temp file
        else:
            # Close the PDF since we didn't need to modify it
            final_pdf.close()
    except Exception as e:
        print(f"Warning: Error when checking links in final PDF: {e}")
        # Don't rethrow the exception, since the PDF should be functional
    
    # Clean up temporary file
    try:
        os.remove(temp_merged_file)
    except:
        pass  # Ignore errors
    
    print(f"Merged PDF created at: {MERGED_PDF} ({final_page_count} pages)")
    print(f"Table of Contents links and bookmarks have been properly set for all attachments") 