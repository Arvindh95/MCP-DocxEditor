"""
Docx MCP Server using FastMCP and python-docx.

This server allows ChatGPT to read, search, and modify DOCX files.
Supports:
- Dynamic document switching with fuzzy search
- Reading and searching paragraphs
- Inserting content at arbitrary positions with formatting (bold, italic, alignment, styles)
- Template placeholders (<<Name>>, {{Date}}, etc.)
- Find and replace text for surgical edits
- Automatic table detection and conversion (markdown & tab-delimited)
- Full table manipulation (read, update cells, add/delete rows)
"""

import os
import re
from difflib import SequenceMatcher
from docx import Document
from docx.shared import Pt
from docx.table import Table
from docx.enum.text import WD_ALIGN_PARAGRAPH
from fastmcp import FastMCP

# Define the default document path (now in documents/ folder)
DEFAULT_DOCX_PATH = os.environ.get("DOCX_PATH", os.path.join("documents", "MCP.docx"))

# Global variable to track current active document
CURRENT_DOCX_PATH = DEFAULT_DOCX_PATH

# Create the MCP server
mcp = FastMCP(name="Docx Editor")


def get_document():
    """Helper to load the document or create a new one if it doesn't exist."""
    global CURRENT_DOCX_PATH
    if os.path.exists(CURRENT_DOCX_PATH):
        return Document(CURRENT_DOCX_PATH)
    return Document()


def save_document(doc):
    """Save document to the current active path."""
    global CURRENT_DOCX_PATH
    doc.save(CURRENT_DOCX_PATH)


def find_document_by_name(query: str, search_dir: str = ".") -> list:
    """
    Find .docx files by fuzzy name matching in the specified directory.
    Returns list of (file_path, filename, score) tuples sorted by score.
    """
    results = []

    # Walk through directory to find all .docx files
    for root, dirs, files in os.walk(search_dir):
        for file in files:
            if file.endswith('.docx') and not file.startswith('~$'):  # Skip temp files
                file_path = os.path.join(root, file)
                filename = os.path.splitext(file)[0]  # Remove .docx extension

                # Calculate similarity
                score = similarity(query.lower(), filename.lower())

                # Also check if query is contained in filename
                if query.lower() in filename.lower():
                    score = max(score, 0.8)

                if score >= 0.3:  # Minimum threshold
                    results.append((file_path, file, score))

    # Sort by score descending
    results.sort(key=lambda x: x[2], reverse=True)
    return results


def similarity(a: str, b: str) -> float:
    """Calculate similarity ratio between two strings using SequenceMatcher."""
    return SequenceMatcher(None, a.lower(), b.lower()).ratio()


def find_paragraph_by_text(doc, query: str, threshold: float = 0.5):
    """
    Find a paragraph by fuzzy text matching.
    Returns (index, paragraph, score) or None if not found.
    """
    best_match = None
    best_score = 0

    for idx, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue

        # Check if query is contained in paragraph
        if query.lower() in text.lower():
            score = 0.9 + (0.1 * similarity(query, text))
            if score > best_score:
                best_score = score
                best_match = (idx, para, score)
        else:
            # Use fuzzy matching
            score = similarity(query, text)
            if score > best_score and score >= threshold:
                best_score = score
                best_match = (idx, para, score)

    return best_match


def find_placeholders(doc) -> list:
    """Find all placeholders in the document (<<...>> or {{...>>), including inside tables."""
    placeholders = []
    pattern = re.compile(r'(<<[^<>]+>>|\{\{[^{}]+\}\})')

    # Search in paragraphs
    for idx, para in enumerate(doc.paragraphs):
        matches = pattern.findall(para.text)
        for match in matches:
            placeholders.append({
                "placeholder": match,
                "location_type": "paragraph",
                "paragraph_index": idx,
                "context": para.text[:100] + "..." if len(para.text) > 100 else para.text
            })

    # Search in tables
    for table_idx, table in enumerate(doc.tables):
        for row_idx, row in enumerate(table.rows):
            for col_idx, cell in enumerate(row.cells):
                cell_text = cell.text
                matches = pattern.findall(cell_text)
                for match in matches:
                    placeholders.append({
                        "placeholder": match,
                        "location_type": "table",
                        "table_index": table_idx,
                        "row": row_idx,
                        "column": col_idx,
                        "context": cell_text[:100] + "..." if len(cell_text) > 100 else cell_text
                    })

    return placeholders


def replace_text_in_paragraph(paragraph, old_text: str, new_text: str) -> bool:
    """
    Replace text in a paragraph, handling the case where text might be split across runs.
    Returns True if replacement was made.
    """
    # First try simple replacement in the full text
    if old_text in paragraph.text:
        # Try to find and replace in individual runs first
        for run in paragraph.runs:
            if old_text in run.text:
                run.text = run.text.replace(old_text, new_text)
                return True

        # If not found in individual runs, the text is split across runs
        # We need to rebuild the paragraph
        full_text = paragraph.text
        new_full_text = full_text.replace(old_text, new_text)

        # Clear all runs and add the new text
        for run in paragraph.runs:
            run.text = ""

        if paragraph.runs:
            paragraph.runs[0].text = new_full_text
        else:
            paragraph.add_run(new_full_text)

        return True

    return False


def detect_table_format(text: str) -> tuple[str, list[list[str]]]:
    """
    Detect if text is a table and parse it.
    Returns (format_type, table_data) where format_type is 'markdown', 'tab', or None.
    table_data is a list of rows, where each row is a list of cell values.
    """
    lines = [line.strip() for line in text.strip().split('\n') if line.strip()]

    if not lines:
        return None, []

    # Check for markdown table format (| col1 | col2 |)
    if lines[0].startswith('|') and lines[0].endswith('|'):
        table_data = []
        for line in lines:
            # Skip separator lines like |------|------| or |:-----|-----:|
            # Check if line contains only dashes, colons, spaces, and pipes
            if re.match(r'^\s*\|[\s\-:|]+\|\s*$', line):
                # Further check: make sure it's mostly dashes (not actual content)
                content = line.replace('|', '').replace('-', '').replace(':', '').replace(' ', '')
                if len(content) == 0:
                    continue

            # Parse cells
            cells = [cell.strip() for cell in line.split('|')[1:-1]]

            # Skip if all cells are just dashes (another way separator lines appear)
            if all(re.match(r'^[\s\-:]+$', cell) for cell in cells if cell):
                continue

            if cells:
                table_data.append(cells)

        if len(table_data) >= 1:  # At least header (changed from 2 to 1)
            return 'markdown', table_data

    # Check for tab-delimited format
    tab_rows = []
    for line in lines:
        if '\t' in line:
            cells = [cell.strip() for cell in line.split('\t')]
            tab_rows.append(cells)

    if tab_rows and len(tab_rows) >= 2:
        # Check if all rows have similar column count
        col_counts = [len(row) for row in tab_rows]
        if max(col_counts) - min(col_counts) <= 1:  # Allow 1 column variance
            return 'tab', tab_rows

    return None, []


def apply_paragraph_formatting(paragraph, bold: bool = False, italic: bool = False,
                               alignment: str = None, style: str = None):
    """
    Apply formatting to a paragraph.

    Args:
        paragraph: The paragraph object to format
        bold: Make text bold
        italic: Make text italic
        alignment: Text alignment - "left", "center", "right", "justify"
        style: Word style name - "Heading 1", "Title", etc.
    """
    # Apply text formatting to all runs
    if bold or italic:
        for run in paragraph.runs:
            if bold:
                run.bold = True
            if italic:
                run.italic = True

    # Apply alignment
    if alignment:
        alignment_map = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "justify": WD_ALIGN_PARAGRAPH.JUSTIFY
        }
        if alignment.lower() in alignment_map:
            paragraph.alignment = alignment_map[alignment.lower()]

    # Apply style
    if style:
        try:
            paragraph.style = style
        except KeyError:
            # Style doesn't exist, ignore
            pass


def create_word_table(doc, table_data: list[list[str]], has_header: bool = True):
    """
    Create a Word table from parsed table data.
    Returns the created table object.
    """
    if not table_data or not table_data[0]:
        return None

    rows = len(table_data)
    cols = max(len(row) for row in table_data)

    # Create table
    table = doc.add_table(rows=rows, cols=cols)

    # Apply the most common table style, fall back to default if not available
    try:
        table.style = 'Table Grid'
    except KeyError:
        # If style not available, leave as default (no borders)
        pass

    # Fill in data
    for i, row_data in enumerate(table_data):
        row = table.rows[i]
        for j, cell_value in enumerate(row_data):
            if j < len(row.cells):
                row.cells[j].text = cell_value

                # Make header row bold
                if i == 0 and has_header:
                    for paragraph in row.cells[j].paragraphs:
                        for run in paragraph.runs:
                            run.bold = True

    return table


# ============================================
# Document Management Tools
# ============================================

@mcp.tool()
async def get_current_document() -> dict:
    """
    Get the name and path of the currently active document.
    """
    global CURRENT_DOCX_PATH
    return {
        "current_document": os.path.basename(CURRENT_DOCX_PATH),
        "full_path": os.path.abspath(CURRENT_DOCX_PATH),
        "exists": os.path.exists(CURRENT_DOCX_PATH)
    }


@mcp.tool()
async def list_documents(search_dir: str = ".") -> dict:
    """
    List all .docx files in the specified directory and subdirectories.

    Args:
        search_dir: Directory to search in (default is current directory)
    """
    documents = []

    for root, dirs, files in os.walk(search_dir):
        for file in files:
            if file.endswith('.docx') and not file.startswith('~$'):
                file_path = os.path.join(root, file)
                rel_path = os.path.relpath(file_path, search_dir)
                documents.append({
                    "filename": file,
                    "path": rel_path,
                    "full_path": os.path.abspath(file_path)
                })

    return {
        "total": len(documents),
        "documents": documents[:50]  # Limit to 50 results
    }


@mcp.tool()
async def switch_document(query: str, search_dir: str = ".") -> dict:
    """
    SWITCH to a different document by name using fuzzy search.
    This changes which document all other tools will operate on.

    Args:
        query: The document name to search for (fuzzy matching)
        search_dir: Directory to search in (default is current directory)
    """
    global CURRENT_DOCX_PATH

    # Find matching documents
    results = find_document_by_name(query, search_dir)

    if not results:
        return {
            "error": f"No .docx files found matching '{query}' in '{search_dir}'",
            "suggestion": "Use list_documents to see available files"
        }

    # Use the best match
    best_match = results[0]
    file_path, filename, score = best_match

    # Update the current document path
    CURRENT_DOCX_PATH = file_path

    # Return results with top matches
    top_matches = [
        {"filename": fn, "score": round(sc, 2), "path": fp}
        for fp, fn, sc in results[:5]
    ]

    return {
        "status": "success",
        "message": f"Switched to document: {filename}",
        "current_document": filename,
        "full_path": os.path.abspath(file_path),
        "match_score": round(score, 2),
        "other_matches": top_matches[1:] if len(top_matches) > 1 else []
    }


# ============================================
# ChatGPT Required Tools: search and fetch
# ============================================

@mcp.tool()
async def search(query: str) -> dict:
    """
    Search for paragraphs in the DOCX document using fuzzy text matching.
    Returns a list of matching paragraphs with IDs for retrieval.

    Args:
        query: The search query string
    """
    doc = get_document()
    results = []

    for idx, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue

        # Check containment first
        if query.lower() in text.lower():
            score = 0.9
        else:
            score = similarity(query, text)

        if score >= 0.3:  # Lower threshold for more results
            results.append({
                "id": f"para-{idx}",
                "score": round(score, 2),
                "text": text[:200] + "..." if len(text) > 200 else text
            })

    # Sort by score descending
    results.sort(key=lambda x: x["score"], reverse=True)

    return {"results": results[:10]}


@mcp.tool()
async def fetch(id: str) -> dict:
    """
    Retrieve the full content of a paragraph by its ID.
    Use IDs returned from the search tool.

    Args:
        id: The paragraph ID (e.g., "para-5")
    """
    doc = get_document()

    # Parse the index from the ID
    if not id.startswith("para-"):
        return {"error": f"Invalid ID format: {id}. Expected 'para-N'"}

    try:
        idx = int(id.replace("para-", ""))
    except ValueError:
        return {"error": f"Invalid ID format: {id}"}

    if idx < 0 or idx >= len(doc.paragraphs):
        return {"error": f"Paragraph {id} not found"}

    para = doc.paragraphs[idx]

    return {
        "id": id,
        "index": idx,
        "text": para.text,
        "style": para.style.name if para.style else None
    }


# ============================================
# Reading Tools
# ============================================

@mcp.tool()
async def read_document() -> dict:
    """
    Read the full text content of the document.
    """
    doc = get_document()
    full_text = []
    for para in doc.paragraphs:
        if para.text.strip():
            full_text.append(para.text)
    return {"text": "\n".join(full_text)}


@mcp.tool()
async def get_paragraphs(limit: int = 50, start_index: int = 0) -> dict:
    """
    Get a list of paragraphs from the document with their IDs.

    Args:
        limit: Maximum number of paragraphs to return (default 50)
        start_index: Index to start reading from (default 0)
    """
    doc = get_document()
    paragraphs = []

    for idx, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if text:
            paragraphs.append({
                "id": f"para-{idx}",
                "text": text[:100] + "..." if len(text) > 100 else text
            })

    # Slice the result
    sliced = paragraphs[start_index : start_index + limit]

    return {
        "total_paragraphs": len(paragraphs),
        "start_index": start_index,
        "showing": len(sliced),
        "paragraphs": sliced
    }


# ============================================
# Editing Tools
# ============================================

@mcp.tool()
async def add_paragraph(
    text: str,
    bold: bool = False,
    italic: bool = False,
    alignment: str = None,
    style: str = None
) -> dict:
    """
    Append a new paragraph to the end of the document with optional formatting.

    Args:
        text: Text content to add
        bold: Make text bold (default False)
        italic: Make text italic (default False)
        alignment: Text alignment - "left", "center", "right", "justify" (default left)
        style: Word style name - "Heading 1", "Heading 2", "Title", "Subtitle", etc.
    """
    doc = get_document()
    para = doc.add_paragraph(text)

    # Apply formatting
    apply_paragraph_formatting(para, bold=bold, italic=italic, alignment=alignment, style=style)

    save_document(doc)
    return {"status": "success", "message": "Paragraph added to end of document."}


@mcp.tool()
async def update_paragraph(
    id: str,
    text: str,
    bold: bool = None,
    italic: bool = None,
    alignment: str = None,
    style: str = None
) -> dict:
    """
    Update the text content of a paragraph by its ID with optional formatting.
    This replaces the entire paragraph text.

    Args:
        id: The paragraph ID to update (e.g., "para-5")
        text: The new text content for the paragraph
        bold: Make text bold (None = don't change)
        italic: Make text italic (None = don't change)
        alignment: Text alignment - "left", "center", "right", "justify" (None = don't change)
        style: Word style name - "Heading 1", "Title", etc. (None = don't change)
    """
    doc = get_document()

    # Parse the index from the ID
    if not id.startswith("para-"):
        return {"error": f"Invalid ID format: {id}. Expected 'para-N'"}

    try:
        idx = int(id.replace("para-", ""))
    except ValueError:
        return {"error": f"Invalid ID format: {id}"}

    if idx < 0 or idx >= len(doc.paragraphs):
        return {"error": f"Paragraph {id} not found"}

    para = doc.paragraphs[idx]

    # Clear existing runs and set new text
    for run in para.runs:
        run.text = ""
    if para.runs:
        para.runs[0].text = text
    else:
        para.add_run(text)

    # Apply formatting if specified (using False instead of None to actually apply)
    if bold is not None or italic is not None or alignment or style:
        apply_paragraph_formatting(
            para,
            bold=bold if bold is not None else False,
            italic=italic if italic is not None else False,
            alignment=alignment,
            style=style
        )

    save_document(doc)
    return {"status": "success", "message": f"Paragraph {id} updated."}


@mcp.tool()
async def insert_after_text(
    query: str,
    text: str,
    threshold: float = 0.5,
    bold: bool = False,
    italic: bool = False,
    alignment: str = None,
    style: str = None
) -> dict:
    """
    INSERT CONTENT AT ANY POSITION in the document with optional formatting.
    This is the PRIMARY tool for adding content at a specific location.
    Finds any paragraph by text search and inserts new content immediately after it.

    Use this tool whenever you need to:
    - Insert after a specific section
    - Add content after certain text
    - Place new paragraphs at arbitrary positions (not just at the end)

    Args:
        query: The text to search for - finds the paragraph containing or matching this text
        text: The text content to insert as a new paragraph after the found location
        threshold: Minimum similarity score 0-1 (default 0.5)
        bold: Make text bold (default False)
        italic: Make text italic (default False)
        alignment: Text alignment - "left", "center", "right", "justify" (default left)
        style: Word style name - "Heading 1", "Title", etc.
    """
    doc = get_document()

    match = find_paragraph_by_text(doc, query, threshold)

    if not match:
        return {"error": f"No paragraph found matching: '{query}'"}

    idx, para, score = match

    # Insert a new paragraph after the found one
    # python-docx doesn't have a direct insert_after, so we need to work with the XML
    new_para = doc.add_paragraph(text)

    # Move the new paragraph to the correct position
    para._element.addnext(new_para._element)

    # Apply formatting
    apply_paragraph_formatting(new_para, bold=bold, italic=italic, alignment=alignment, style=style)

    save_document(doc)

    return {
        "status": "success",
        "message": f"Content inserted after paragraph {idx} (match score: {score:.2f})",
        "matched_text": para.text[:100] + "..." if len(para.text) > 100 else para.text
    }


@mcp.tool()
async def insert_after_heading(heading_text: str, text: str) -> dict:
    """
    INSERT CONTENT AFTER A SECTION/HEADING.
    Finds a heading by text and inserts new content at the end of that section
    (before the next heading of same or higher level).

    Args:
        heading_text: The heading/section title to find
        text: The text content to insert after that section
    """
    doc = get_document()

    # Find the heading
    heading_idx = None
    heading_level = None

    for idx, para in enumerate(doc.paragraphs):
        if para.style and para.style.name.startswith('Heading'):
            if heading_text.lower() in para.text.lower():
                heading_idx = idx
                # Extract heading level (e.g., "Heading 1" -> 1)
                try:
                    heading_level = int(para.style.name.replace('Heading ', ''))
                except ValueError:
                    heading_level = 1
                break

    if heading_idx is None:
        # Try fuzzy matching on all paragraphs if no heading style found
        match = find_paragraph_by_text(doc, heading_text, 0.6)
        if match:
            heading_idx = match[0]
            heading_level = None

    if heading_idx is None:
        return {"error": f"No heading found matching: '{heading_text}'"}

    # Find the end of this section (next heading of same or higher level)
    insert_idx = heading_idx

    if heading_level:
        for idx in range(heading_idx + 1, len(doc.paragraphs)):
            para = doc.paragraphs[idx]
            if para.style and para.style.name.startswith('Heading'):
                try:
                    level = int(para.style.name.replace('Heading ', ''))
                    if level <= heading_level:
                        insert_idx = idx - 1
                        break
                except ValueError:
                    pass
            insert_idx = idx

    # Insert the new paragraph
    target_para = doc.paragraphs[insert_idx]
    new_para = doc.add_paragraph(text)
    target_para._element.addnext(new_para._element)

    save_document(doc)

    return {
        "status": "success",
        "message": f"Content inserted after section '{heading_text}' (at paragraph {insert_idx})"
    }


# ============================================
# Template/Placeholder Tools
# ============================================

@mcp.tool()
async def list_placeholders() -> dict:
    """
    TEMPLATE SUPPORT - CALL THIS FIRST when working with templates.
    Find all placeholders in the document like <<Name>>, <<Poem>>, {{Date}}, etc.

    IMPORTANT: If a user mentions any placeholder (text with << >> or {{ }}),
    ALWAYS call this tool first to discover placeholders, then use replace_placeholder to fill them in.
    """
    doc = get_document()
    placeholders = find_placeholders(doc)

    # Get unique placeholder names
    unique = list(set(p["placeholder"] for p in placeholders))

    if not unique:
        return {
            "found": 0,
            "message": "No placeholders found. Placeholders should be formatted as <<Name>> or {{Name}}."
        }

    return {
        "found": len(unique),
        "placeholders": unique,
        "details": placeholders
    }


@mcp.tool()
async def replace_placeholder(placeholder: str, value: str) -> dict:
    """
    TEMPLATE SUPPORT - USE THIS TO FILL PLACEHOLDERS.
    Replaces a placeholder like <<Poem>> or {{Name}} with actual content IN-PLACE.
    This REMOVES the placeholder and puts the new content exactly where it was.
    Works in both paragraphs AND table cells.

    ALWAYS use this tool (not insert_after_text) when working with template placeholders.

    Args:
        placeholder: The EXACT placeholder to replace, including brackets (e.g., "<<Poem>>")
        value: The content to put in place of the placeholder
    """
    doc = get_document()
    count = 0

    # Replace in paragraphs
    for para in doc.paragraphs:
        if placeholder in para.text:
            if replace_text_in_paragraph(para, placeholder, value):
                count += 1

    # Replace in table cells
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if placeholder in cell.text:
                    cell.text = cell.text.replace(placeholder, value)
                    count += 1

    if count == 0:
        return {"error": f"Placeholder '{placeholder}' was not found in the document."}

    save_document(doc)

    return {
        "status": "success",
        "message": f"Replaced {count} occurrence(s) of '{placeholder}' with the provided value.",
        "replacements": count
    }


@mcp.tool()
async def replace_placeholders(replacements: dict) -> dict:
    """
    TEMPLATE SUPPORT: Replace multiple placeholders at once.
    Provide a mapping of placeholder -> value pairs to fill in a template document efficiently.
    Works in both paragraphs AND table cells.

    Args:
        replacements: Object mapping placeholders to their values, e.g., {"<<Name>>": "John", "<<Date>>": "2024-01-15"}
    """
    doc = get_document()
    total_count = 0
    results = {}

    for placeholder, value in replacements.items():
        count = 0

        # Replace in paragraphs
        for para in doc.paragraphs:
            if placeholder in para.text:
                if replace_text_in_paragraph(para, placeholder, value):
                    count += 1

        # Replace in table cells
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if placeholder in cell.text:
                        cell.text = cell.text.replace(placeholder, value)
                        count += 1

        results[placeholder] = count
        total_count += count

    save_document(doc)

    return {
        "status": "success",
        "message": f"Replaced {total_count} placeholder occurrence(s) across {len(replacements)} unique placeholder(s).",
        "details": results
    }


# ============================================
# Table Tools
# ============================================

@mcp.tool()
async def insert_table(text: str, has_header: bool = True, query: str = None) -> dict:
    """
    INSERT A PROPER WORD TABLE into the document.
    Automatically detects and converts markdown or tab-delimited table text into a real Word table.

    Supports two formats:
    1. Markdown: | Column1 | Column2 |
    2. Tab-delimited: Column1\tColumn2

    Args:
        text: The table content (markdown or tab-delimited format)
        has_header: Whether the first row is a header (default True)
        query: Optional - text to search for to insert table after. If not provided, adds to end.
    """
    doc = get_document()

    # Detect and parse table format
    format_type, table_data = detect_table_format(text)

    if not format_type:
        return {
            "error": "Could not detect table format. Please use markdown (| col |) or tab-delimited format.",
            "received_text": text[:200] + "..." if len(text) > 200 else text
        }

    # If query provided, find insertion point
    insert_after_para = None
    if query:
        match = find_paragraph_by_text(doc, query, 0.5)
        if not match:
            return {"error": f"Could not find paragraph matching: '{query}'"}
        insert_after_para = match[1]

    # Create the table
    table = create_word_table(doc, table_data, has_header)

    if not table:
        return {"error": "Failed to create table"}

    # Move table to correct position if query was provided
    if insert_after_para:
        # Add space before table
        space_before = doc.add_paragraph()
        insert_after_para._element.addnext(space_before._element)

        # Add table after the space
        space_before._element.addnext(table._element)

        # Add space after table
        space_after = doc.add_paragraph()
        table._element.addnext(space_after._element)
    else:
        # If adding to end of document, add spaces too
        doc.add_paragraph()  # Space before
        # Table is already at the end
        doc.add_paragraph()  # Space after

    save_document(doc)

    return {
        "status": "success",
        "message": f"Inserted {len(table_data)}x{len(table_data[0])} table ({format_type} format detected)",
        "rows": len(table_data),
        "columns": len(table_data[0]),
        "format_detected": format_type
    }


@mcp.tool()
async def convert_text_to_table(query: str, has_header: bool = True, threshold: float = 0.5) -> dict:
    """
    CONVERT EXISTING TEXT PARAGRAPH(S) into a proper Word table.
    Finds paragraph(s) containing table-like text and replaces them with a real Word table.

    Use this when GPT has already inserted table text (markdown or tab-delimited) and you want to
    convert it to a proper formatted table.

    Args:
        query: Text to search for to find the paragraph(s) containing the table
        has_header: Whether the first row is a header (default True)
        threshold: Minimum similarity score for text matching (0-1, default 0.5)
    """
    doc = get_document()

    # Find the paragraph containing table-like text
    match = find_paragraph_by_text(doc, query, threshold)

    if not match:
        return {"error": f"No paragraph found matching: '{query}'"}

    idx, para, score = match

    # Collect consecutive paragraphs that might be part of the table
    table_paragraphs = [para]
    table_indices = [idx]

    # Check if following paragraphs are also table rows
    for i in range(idx + 1, min(idx + 20, len(doc.paragraphs))):
        next_para = doc.paragraphs[i]
        text = next_para.text.strip()

        if not text:
            break  # Empty line signals end of table

        # Check if it looks like a table row
        if '|' in text or '\t' in text:
            table_paragraphs.append(next_para)
            table_indices.append(i)
        else:
            break

    # Combine all paragraphs into one text block
    combined_text = '\n'.join(p.text for p in table_paragraphs)

    # Detect and parse table format
    format_type, table_data = detect_table_format(combined_text)

    if not format_type:
        return {
            "error": "Could not detect table format in the matched paragraph(s).",
            "matched_text": combined_text[:200] + "..." if len(combined_text) > 200 else combined_text
        }

    # Create the Word table
    table = create_word_table(doc, table_data, has_header)

    if not table:
        return {"error": "Failed to create table"}

    # Add space before table
    space_before = doc.add_paragraph()
    table_paragraphs[0]._element.addprevious(space_before._element)

    # Insert table after the space
    space_before._element.addnext(table._element)

    # Add space after table
    space_after = doc.add_paragraph()
    table._element.addnext(space_after._element)

    # Delete all the old text paragraphs
    for p in table_paragraphs:
        p._element.getparent().remove(p._element)

    save_document(doc)

    return {
        "status": "success",
        "message": f"Converted {len(table_paragraphs)} paragraph(s) to {len(table_data)}x{len(table_data[0])} table",
        "rows": len(table_data),
        "columns": len(table_data[0]),
        "format_detected": format_type,
        "paragraphs_removed": len(table_paragraphs)
    }


# ============================================
# Table Manipulation Tools
# ============================================

@mcp.tool()
async def list_tables() -> dict:
    """
    List all tables in the document with their dimensions and preview of first row.
    Returns table indices that can be used with other table tools.
    """
    doc = get_document()
    tables_info = []

    for idx, table in enumerate(doc.tables):
        rows = len(table.rows)
        cols = len(table.columns) if rows > 0 else 0

        # Get first row as preview
        first_row = []
        if rows > 0:
            first_row = [cell.text.strip() for cell in table.rows[0].cells]

        tables_info.append({
            "table_index": idx,
            "rows": rows,
            "columns": cols,
            "first_row_preview": first_row[:5]  # First 5 cells
        })

    return {
        "total_tables": len(tables_info),
        "tables": tables_info
    }


@mcp.tool()
async def read_table(table_index: int) -> dict:
    """
    Read the complete contents of a table by its index.

    Args:
        table_index: The index of the table (from list_tables, starting at 0)
    """
    doc = get_document()

    if table_index < 0 or table_index >= len(doc.tables):
        return {"error": f"Table index {table_index} not found. Document has {len(doc.tables)} table(s)."}

    table = doc.tables[table_index]
    table_data = []

    for row in table.rows:
        row_data = [cell.text.strip() for cell in row.cells]
        table_data.append(row_data)

    return {
        "table_index": table_index,
        "rows": len(table.rows),
        "columns": len(table.columns),
        "data": table_data
    }


@mcp.tool()
async def update_table_cell(table_index: int, row: int, column: int, text: str) -> dict:
    """
    Update a specific cell in a table by row and column index.
    This is the PRIMARY tool for editing table cells directly.

    Args:
        table_index: The index of the table (from list_tables, starting at 0)
        row: The row index (starting at 0)
        column: The column index (starting at 0)
        text: The new text content for the cell
    """
    doc = get_document()

    if table_index < 0 or table_index >= len(doc.tables):
        return {"error": f"Table index {table_index} not found. Document has {len(doc.tables)} table(s)."}

    table = doc.tables[table_index]

    if row < 0 or row >= len(table.rows):
        return {"error": f"Row index {row} out of range. Table has {len(table.rows)} rows."}

    if column < 0 or column >= len(table.columns):
        return {"error": f"Column index {column} out of range. Table has {len(table.columns)} columns."}

    # Update the cell
    cell = table.rows[row].cells[column]
    cell.text = text

    save_document(doc)

    return {
        "status": "success",
        "message": f"Updated cell at table {table_index}, row {row}, column {column}",
        "new_value": text
    }


@mcp.tool()
async def add_table_row(table_index: int, row_data: list = None) -> dict:
    """
    Add a new row to the end of a table, optionally with data.

    Args:
        table_index: The index of the table (from list_tables, starting at 0)
        row_data: Optional list of cell values for the new row
    """
    doc = get_document()

    if table_index < 0 or table_index >= len(doc.tables):
        return {"error": f"Table index {table_index} not found. Document has {len(doc.tables)} table(s)."}

    table = doc.tables[table_index]
    new_row = table.add_row()

    # Fill in data if provided
    if row_data:
        for col_idx, value in enumerate(row_data):
            if col_idx < len(new_row.cells):
                new_row.cells[col_idx].text = str(value)

    save_document(doc)

    return {
        "status": "success",
        "message": f"Added row to table {table_index}",
        "new_row_index": len(table.rows) - 1,
        "total_rows": len(table.rows)
    }


@mcp.tool()
async def update_table_row(table_index: int, row: int, row_data: list) -> dict:
    """
    Update an entire row in a table with new data.

    Args:
        table_index: The index of the table (from list_tables, starting at 0)
        row: The row index (starting at 0)
        row_data: List of cell values for the row
    """
    doc = get_document()

    if table_index < 0 or table_index >= len(doc.tables):
        return {"error": f"Table index {table_index} not found. Document has {len(doc.tables)} table(s)."}

    table = doc.tables[table_index]

    if row < 0 or row >= len(table.rows):
        return {"error": f"Row index {row} out of range. Table has {len(table.rows)} rows."}

    # Update each cell in the row
    table_row = table.rows[row]
    for col_idx, value in enumerate(row_data):
        if col_idx < len(table_row.cells):
            table_row.cells[col_idx].text = str(value)

    save_document(doc)

    return {
        "status": "success",
        "message": f"Updated row {row} in table {table_index}",
        "cells_updated": min(len(row_data), len(table_row.cells))
    }


@mcp.tool()
async def delete_table_row(table_index: int, row: int) -> dict:
    """
    Delete a row from a table.

    Args:
        table_index: The index of the table (from list_tables, starting at 0)
        row: The row index to delete (starting at 0)
    """
    doc = get_document()

    if table_index < 0 or table_index >= len(doc.tables):
        return {"error": f"Table index {table_index} not found. Document has {len(doc.tables)} table(s)."}

    table = doc.tables[table_index]

    if row < 0 or row >= len(table.rows):
        return {"error": f"Row index {row} out of range. Table has {len(table.rows)} rows."}

    # Delete the row using XML manipulation
    table._element.remove(table.rows[row]._element)

    save_document(doc)

    return {
        "status": "success",
        "message": f"Deleted row {row} from table {table_index}",
        "remaining_rows": len(table.rows)
    }


# ============================================
# Formatting Tools
# ============================================

@mcp.tool()
async def format_paragraph(
    query: str,
    bold: bool = None,
    italic: bool = None,
    underline: bool = None,
    alignment: str = None,
    font_size: int = None,
    style: str = None,
    threshold: float = 0.5
) -> dict:
    """
    Apply formatting to an existing paragraph found by fuzzy text search.
    Use this to format paragraphs that have already been added to the document.

    Args:
        query: Text to search for to find the paragraph (fuzzy matching)
        bold: Make text bold (True/False, None = don't change)
        italic: Make text italic (True/False, None = don't change)
        underline: Make text underlined (True/False, None = don't change)
        alignment: Text alignment - "left", "center", "right", "justify" (None = don't change)
        font_size: Font size in points, e.g., 12, 14, 16 (None = don't change)
        style: Word style name - "Heading 1", "Title", etc. (None = don't change)
        threshold: Minimum similarity score for text matching (0-1, default 0.5)
    """
    doc = get_document()

    # Find the paragraph
    match = find_paragraph_by_text(doc, query, threshold)

    if not match:
        return {"error": f"No paragraph found matching: '{query}'"}

    idx, para, score = match

    # Apply text formatting (bold, italic, underline)
    if bold is not None or italic is not None or underline is not None:
        for run in para.runs:
            if bold is not None:
                run.bold = bold
            if italic is not None:
                run.italic = italic
            if underline is not None:
                run.underline = underline

    # Apply font size
    if font_size:
        for run in para.runs:
            run.font.size = Pt(font_size)

    # Apply alignment
    if alignment:
        alignment_map = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "justify": WD_ALIGN_PARAGRAPH.JUSTIFY
        }
        if alignment.lower() in alignment_map:
            para.alignment = alignment_map[alignment.lower()]

    # Apply style
    if style:
        try:
            para.style = style
        except KeyError:
            pass  # Style doesn't exist, ignore

    save_document(doc)

    return {
        "status": "success",
        "message": f"Applied formatting to paragraph {idx} (match score: {score:.2f})",
        "matched_text": para.text[:100] + "..." if len(para.text) > 100 else para.text
    }


# ============================================
# Text Replacement Tool
# ============================================

@mcp.tool()
async def replace_text(old_text: str, new_text: str) -> dict:
    """
    FIND AND REPLACE any text anywhere in the document.
    This is useful for editing generated content without replacing entire paragraphs.

    Use this tool to:
    - Fix typos or errors in previously generated text
    - Update specific words or phrases throughout the document
    - Make surgical edits to GPT-generated content

    Args:
        old_text: The exact text to find and replace
        new_text: The text to replace it with
    """
    doc = get_document()
    count = 0
    affected_paragraphs = []

    for idx, para in enumerate(doc.paragraphs):
        if old_text in para.text:
            if replace_text_in_paragraph(para, old_text, new_text):
                count += 1
                affected_paragraphs.append({
                    "id": f"para-{idx}",
                    "preview": para.text[:100] + "..." if len(para.text) > 100 else para.text
                })

    if count == 0:
        return {
            "error": f"Text '{old_text}' was not found in the document.",
            "suggestion": "Make sure the text matches exactly (case-sensitive)."
        }

    save_document(doc)

    return {
        "status": "success",
        "message": f"Replaced {count} occurrence(s) of '{old_text}' with '{new_text}'.",
        "replacements": count,
        "affected_paragraphs": affected_paragraphs
    }


# ============================================
# Save Tool
# ============================================

@mcp.tool()
async def save_document_as(filename: str) -> dict:
    """
    Save the document to a new filename.
    Use this to create a copy or save with a different name.

    Args:
        filename: The filename to save as (e.g., "output.docx")
    """
    doc = get_document()

    # Ensure .docx extension
    if not filename.endswith('.docx'):
        filename += '.docx'

    doc.save(filename)

    return {
        "status": "success",
        "message": f"Document saved to: {filename}",
        "path": os.path.abspath(filename)
    }


if __name__ == "__main__":
    # Ensure the default document exists or create a blank one
    if not os.path.exists(CURRENT_DOCX_PATH):
        print(f"Creating new document at {CURRENT_DOCX_PATH}")
        doc = Document()
        doc.save(CURRENT_DOCX_PATH)

    print(f"Starting Docx MCP Server...")
    print(f"Current document: {os.path.abspath(CURRENT_DOCX_PATH)}")
    print(f"Use 'switch_document' to work on a different file")
    print(f"\nAvailable tools:")
    print("  - get_current_document, list_documents, switch_document")
    print("  - search, fetch (ChatGPT required)")
    print("  - read_document, get_paragraphs")
    print("  - add_paragraph, update_paragraph")
    print("  - insert_after_text, insert_after_heading")
    print("  - list_placeholders, replace_placeholder, replace_placeholders")
    print("  - replace_text (find and replace any text)")
    print("  - insert_table, convert_text_to_table")
    print("  - list_tables, read_table, update_table_cell, add_table_row, update_table_row, delete_table_row")
    print("  - format_paragraph (apply formatting to existing text)")
    print("  - save_document_as")
    print()

    mcp.run(transport="sse", host="0.0.0.0", port=8000)
