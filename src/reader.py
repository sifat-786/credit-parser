"""
reader.py
---------
Reads a .docx manuscript and extracts:
  1. The list of author names from the paper's author line
  2. The raw text paragraphs of the contributions section

Returns (authors, section_paragraphs).
Returns ([], []) for corrupted or unreadable files.
Returns (authors, []) when no contributions section exists.
"""

import re
from docx import Document
from docx.opc.exceptions import PackageNotFoundError

# Section headings we recognise as contributions sections (lowercase, for matching)
CONTRIBUTIONS_HEADINGS = [
    "author contributions",
    "author contribution",
    "author contributions statement",
    "contributorship statement",
    "contributorship",
    "credit taxonomy",
    "contributor roles",
]

# Section headings that mark the END of the contributions section
NEXT_SECTION_KEYWORDS = [
    "acknowledgement", "acknowledgment", "conflict of interest",
    "competing interest", "data availability", "references", "appendix",
    "supplementary", "ethics statement", "declaration",
]

# Standalone single-word section headings (only match if text IS just this word/phrase)
NEXT_SECTION_STANDALONE = [
    "funding", "references", "appendix",
]


def read_manuscript(filepath: str) -> tuple[list[str], list[str]]:
    """
    Open a .docx file and return (author_names, contribution_paragraphs).

    author_names: list of clean author name strings from the paper's byline.
    contribution_paragraphs: list of non-empty paragraph strings from the
        contributions section (empty list if none found).

    Never raises — returns ([], []) on any file error.
    """
    try:
        doc = Document(filepath)
    except (PackageNotFoundError, Exception) as e:
        print(f"[reader] Could not open {filepath}: {e}")
        return [], []

    paragraphs = [p.text.strip() for p in doc.paragraphs]
    # Drop blanks for processing, but keep index alignment
    non_empty = [(i, p) for i, p in enumerate(paragraphs) if p]

    if not non_empty:
        return [], []

    authors = _extract_authors(non_empty)
    section_paras = _extract_contributions_section(non_empty)

    return authors, section_paras


# ---------------------------------------------------------------------------
# Author extraction
# ---------------------------------------------------------------------------

def _extract_authors(non_empty: list[tuple[int, str]]) -> list[str]:
    """
    The author byline is typically the second non-empty paragraph.
    It looks like: "Jane Doe1, Anil Bose2" or "Maria Lopez1, Tom Reid2"
    We clean each name and return them.
    """
    if len(non_empty) < 2:
        return []

    # Try paragraphs 1–4 (index 1 is most common)
    for idx in range(1, min(5, len(non_empty))):
        _, text = non_empty[idx]
        names = _try_parse_author_line(text)
        if names:
            return names

    return []


def _try_parse_author_line(text: str) -> list[str]:
    """
    Return cleaned author names if the text looks like an author byline,
    otherwise return empty list.
    An author byline has comma-separated names, each being Title Case words
    optionally followed by superscript digits/symbols.
    """
    # Must look like "Name1 Name21, Name3 Name42" — capitalised words separated by commas
    # Reject lines that look like addresses (contain "University", "Department", etc.)
    if any(kw in text.lower() for kw in [
        "university", "department", "institute", "faculty", "school",
        "hospital", "center", "centre", "laboratory", "corresponding",
        "received", "revised", "accepted", "abstract", "keywords",
    ]):
        return []

    # Split on comma
    parts = [p.strip() for p in text.split(",") if p.strip()]
    if not parts:
        return []

    names = []
    for part in parts:
        name = _clean_author_name(part)
        if name and _looks_like_person_name(name):
            names.append(name)

    # Require at least 1 name; reject if most parts didn't parse as names
    if len(names) < 1 or (len(parts) > 1 and len(names) < len(parts) / 2):
        return []

    return names


def _looks_like_person_name(name: str) -> bool:
    """A person's name has 1–4 capitalised words."""
    words = name.split()
    if not words:
        return False
    if len(words) > 5:
        return False
    # Each word should start with a capital letter
    return all(w[0].isupper() for w in words if w)


# ---------------------------------------------------------------------------
# Contributions section extraction
# ---------------------------------------------------------------------------

def _extract_contributions_section(non_empty: list[tuple[int, str]]) -> list[str]:
    """
    Find the contributions section heading and collect all paragraphs
    that belong to it (stops at the next section heading).

    Returns the list of content paragraph strings (empty list if not found).
    """
    heading_pos = None
    for i, (_, text) in enumerate(non_empty):
        if _is_contributions_heading(text):
            heading_pos = i
            break

    if heading_pos is None:
        return []

    # Collect paragraphs after the heading until the next section
    section_paras = []
    for i in range(heading_pos + 1, len(non_empty)):
        _, text = non_empty[i]
        if _is_next_section_heading(text):
            break
        section_paras.append(text)

    return section_paras


def _is_contributions_heading(text: str) -> bool:
    """
    Returns True if text is a standalone contributions-section heading.
    Carefully avoids false positives from reference citations.
    """
    # Must be short (headings are never long paragraphs)
    if len(text) > 80:
        return False

    lower = text.lower().strip()

    # Must contain one of our known keywords
    if not any(kw in lower for kw in CONTRIBUTIONS_HEADINGS):
        return False

    # Reject reference citations: they start with [N] or contain year+journal patterns
    if re.match(r'^\[\d+\]', text):
        return False

    # Reject if it looks like a citation: contains year range like 2019;4:1-9
    if re.search(r'\d{4};\d+:\d+', text):
        return False

    # Reject if it contains typical citation words
    if re.search(r'\.\s+[A-Z]\s+\w+\.?\s+\d{4}', text):  # "Smith A. 2019"
        return False

    # The keyword should make up a significant part of the heading text
    # (for a heading, it IS the text; for a citation, it's a small fragment)
    for kw in CONTRIBUTIONS_HEADINGS:
        if kw in lower:
            if len(kw) >= len(lower) * 0.4:
                return True

    return False


def _is_next_section_heading(text: str) -> bool:
    """Returns True if text looks like the start of the section after contributions."""
    if len(text) > 100:
        return False
    stripped = text.strip()
    lower = stripped.lower()
    # Match keywords only as short standalone headings (not buried in prose sentences)
    if any(kw in lower for kw in NEXT_SECTION_KEYWORDS):
        # Must be short enough to be a heading, not a sentence
        if len(stripped) < 50:
            return True
    # Standalone single-word headings — must be an exact (or near-exact) match
    for kw in NEXT_SECTION_STANDALONE:
        if lower == kw or lower == kw + "s":
            return True
    return False


# ---------------------------------------------------------------------------
# Name cleaning (used by both author-line and contributions parsers)
# ---------------------------------------------------------------------------

def _clean_author_name(name: str) -> str:
    """
    Remove honorifics, trailing superscripts, and encoding artefacts.
    e.g. "Dr. Jane Doe1*" → "Jane Doe"
    e.g. "Prof. Anil Bose2\ufffd" → "Anil Bose"
    """
    # Strip honorifics at start
    name = re.sub(
        r'^(?:Dr|Prof|Professor|Mr|Ms|Mrs|Mx)\.?\s+',
        '', name, flags=re.IGNORECASE
    )
    # Strip trailing superscript digits, symbols, replacement chars
    name = re.sub(r'[\d*†‡§\ufffd\u2020\u2021\u00a7]+\s*$', '', name)
    name = name.strip(' ,;:()')
    # Do NOT strip trailing dot — dotted initials like "J.D." must keep their dot
    return name


# Expose for use in parser
clean_author_name = _clean_author_name
