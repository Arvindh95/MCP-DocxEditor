"""
Docx MCP Server using FastMCP and python-docx.

This server allows ChatGPT to read, search, and modify DOCX files.
Supports:
- Reading and searching paragraphs
- Inserting content at arbitrary positions
- Template placeholders (<<Name>>, {{Date}}, etc.)
"""

import os
import re
from difflib import SequenceMatcher
from docx import Document
from docx.shared import Pt
from fastmcp import FastMCP

# Define the path to the document
DOCX_PATH = os.environ.get("DOCX_PATH", "MCP.docx")

# Create the MCP server
mcp = FastMCP(name="Docx Editor")


def get_document():
    """Helper to load the document or create a new one if it doesn't exist."""
    if os.path.exists(DOCX_PATH):
        return Document(DOCX_PATH)
    return Document()


def save_document(doc):
    """Save document to the default path."""
    doc.save(DOCX_PATH)


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
    """Find all placeholders in the document (<<...>> or {{...}})."""
    placeholders = []
    pattern = re.compile(r'(<<[^<>]+>>|\{\{[^{}]+\}\})')

    for idx, para in enumerate(doc.paragraphs):
        matches = pattern.findall(para.text)
        for match in matches:
            placeholders.append({
                "placeholder": match,
                "paragraph_index": idx,
                "context": para.text[:100] + "..." if len(para.text) > 100 else para.text
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
async def add_paragraph(text: str) -> dict:
    """
    Append a new paragraph to the end of the document.

    Args:
        text: Text content to add
    """
    doc = get_document()
    doc.add_paragraph(text)
    save_document(doc)
    return {"status": "success", "message": "Paragraph added to end of document."}


@mcp.tool()
async def update_paragraph(id: str, text: str) -> dict:
    """
    Update the text content of a paragraph by its ID.
    This replaces the entire paragraph text.

    Args:
        id: The paragraph ID to update (e.g., "para-5")
        text: The new text content for the paragraph
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

    save_document(doc)
    return {"status": "success", "message": f"Paragraph {id} updated."}


@mcp.tool()
async def insert_after_text(query: str, text: str, threshold: float = 0.5) -> dict:
    """
    INSERT CONTENT AT ANY POSITION in the document.
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

    ALWAYS use this tool (not insert_after_text) when working with template placeholders.

    Args:
        placeholder: The EXACT placeholder to replace, including brackets (e.g., "<<Poem>>")
        value: The content to put in place of the placeholder
    """
    doc = get_document()
    count = 0

    for para in doc.paragraphs:
        if placeholder in para.text:
            if replace_text_in_paragraph(para, placeholder, value):
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

    Args:
        replacements: Object mapping placeholders to their values, e.g., {"<<Name>>": "John", "<<Date>>": "2024-01-15"}
    """
    doc = get_document()
    total_count = 0
    results = {}

    for placeholder, value in replacements.items():
        count = 0
        for para in doc.paragraphs:
            if placeholder in para.text:
                if replace_text_in_paragraph(para, placeholder, value):
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
    # Ensure the document exists or create a blank one
    if not os.path.exists(DOCX_PATH):
        print(f"Creating new document at {DOCX_PATH}")
        doc = Document()
        doc.save(DOCX_PATH)

    print(f"Starting Docx MCP Server...")
    print(f"Document: {os.path.abspath(DOCX_PATH)}")
    print(f"\nAvailable tools:")
    print("  - search, fetch (ChatGPT required)")
    print("  - read_document, get_paragraphs")
    print("  - add_paragraph, update_paragraph")
    print("  - insert_after_text, insert_after_heading")
    print("  - list_placeholders, replace_placeholder, replace_placeholders")
    print("  - save_document_as")
    print()

    mcp.run(transport="sse", host="0.0.0.0", port=8000)
