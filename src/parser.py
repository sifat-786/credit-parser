"""
parser.py
---------
Parses author-contribution pairs from the raw text of a contributions section.

Given the section paragraphs and the list of known authors (from the paper's
byline), returns a list of (author_name, raw_role_text) tuples — one tuple per
individual role fragment.

Handles these formats:
  • Structured:    "J.D.: Conceptualization, Methodology. A.B.: Software."
  • Semicolon:     "Jane Doe: Investigation; Data curation."
  • Parenthetical: "A.B. (Conceptualization – lead; Methodology – equal)."
  • Prose:         "Jane Doe designed the study and analyzed the data."
  • Multi-line:    One author per paragraph (any of the above formats)
  • All-equal:     "All authors contributed equally to conceptualization, ..."
  • Multi-author:  "AB and CD wrote and edited the paper."
"""

import re
from src.reader import clean_author_name


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def parse_contributions(
    section_paragraphs: list[str],
    known_authors: list[str],
) -> list[tuple[str, str]]:
    """
    Parse contributions section into (author, role_fragment) pairs.

    Args:
        section_paragraphs: raw text paragraphs from the contributions section.
        known_authors: cleaned author names from the paper's byline.

    Returns:
        List of (clean_author_name, raw_role_text) tuples.
        raw_role_text is a single role fragment (not a comma list).
    """
    if not section_paragraphs:
        return []

    # Normalize all paragraphs
    paras = [_normalize(p) for p in section_paragraphs if p.strip()]
    if not paras:
        return []

    combined = " ".join(paras)

    # Special case: "all authors contributed equally"
    if re.search(r'\ball\s+authors?\b', combined, re.IGNORECASE):
        return _parse_all_authors_equally(combined, known_authors)

    results = []
    for para in paras:
        results.extend(_parse_paragraph(para, known_authors))

    return results


# ---------------------------------------------------------------------------
# Format dispatchers
# ---------------------------------------------------------------------------

def _parse_paragraph(text: str, known_authors: list[str]) -> list[tuple[str, str]]:
    """
    Detect which format a paragraph uses and dispatch accordingly.
    Tries formats in order of reliability.
    """
    # 1. Parenthetical: "Name (roles)."
    if _is_parenthetical(text):
        result = _parse_parenthetical(text)
        if result:
            return result

    # 2. Structured (colon-based): "Name: roles."
    if ":" in text:
        result = _parse_structured(text)
        if result:
            return result

    # 3. Prose (natural language)
    return _parse_prose(text, known_authors)


# ---------------------------------------------------------------------------
# Format: Parenthetical  →  "A.B. (Conceptualization – lead; Methodology)."
# ---------------------------------------------------------------------------

def _is_parenthetical(text: str) -> bool:
    """Check if text uses the parenthetical format."""
    # Pattern: something like "Name (role; role)" — parenthesis after a name
    return bool(re.search(r'[A-Z][\w.]*\s*\([^)]+\)', text))


def _parse_parenthetical(text: str) -> list[tuple[str, str]]:
    """
    Parse: "A.B. (Conceptualization – lead; Methodology – equal). J.D. (Writing)."
    """
    results = []
    # Match: Name (roles).  possibly with a period+space before the next block
    pattern = re.compile(
        r'((?:(?:Dr|Prof)\.?\s+)?[A-Z][\w.\s]*?)\s*\(([^)]+)\)',
        re.IGNORECASE
    )
    for m in pattern.finditer(text):
        author_raw = m.group(1).strip()
        roles_raw = m.group(2).strip()
        author = clean_author_name(author_raw)
        if not author:
            continue
        for role in _split_roles(roles_raw):
            results.append((author, role))
    return results


# ---------------------------------------------------------------------------
# Format: Structured  →  "Name: Role1, Role2. Name2: Role3."
# ---------------------------------------------------------------------------

def _parse_structured(text: str) -> list[tuple[str, str]]:
    """
    Parse colon-separated "Author: roles" blocks.
    Handles commas and semicolons as role separators.
    Handles multiple author blocks in one paragraph.
    """
    results = []
    blocks = _split_into_author_blocks(text)
    for block in blocks:
        block = block.strip().rstrip(".")
        if ":" not in block:
            continue
        colon_idx = block.index(":")
        author_raw = block[:colon_idx].strip()
        roles_raw = block[colon_idx + 1:].strip()
        author = clean_author_name(author_raw)
        if not author:
            continue
        for role in _split_roles(roles_raw):
            results.append((author, role))
    return results


def _split_into_author_blocks(text: str) -> list[str]:
    """
    Split a paragraph like "J.D.: Role1. A.B.: Role2." into
    individual author blocks ["J.D.: Role1", "A.B.: Role2"].

    We split on ". " when followed by something that looks like
    an author name + colon.
    """
    # Pattern: period+space followed by optional honorific and a capitalized token then colon
    author_start = (
        r'(?:'
        r'(?:Dr|Prof|Professor|Mr|Ms|Mrs)\.?\s+'  # optional honorific
        r')?'
        r'(?:[A-Z][a-z]*\.(?:\s*[A-Z]\.)*'        # dotted initials: J.D.
        r'|[A-Z]{2,4}'                             # bare initials: AB, EF
        r'|[A-Z][a-z]+(?:\s+[A-Z][a-z]+)*'        # full name: Jane Doe
        r')'
        r'(?:[\d*†‡§\ufffd]*)'                    # optional superscripts
        r'\s*:'                                    # must end with colon
    )
    split_pattern = re.compile(r'\.\s+(?=' + author_start + r')')
    blocks = split_pattern.split(text)
    return [b.strip() for b in blocks if b.strip()]


# ---------------------------------------------------------------------------
# Format: Prose  →  "Jane Doe designed the study. Tom Reid collected data."
# ---------------------------------------------------------------------------

def _parse_prose(text: str, known_authors: list[str]) -> list[tuple[str, str]]:
    """
    Parse prose sentences where each sentence starts with an author name.
    Handles single-author and multi-author sentences ("AB and CD wrote...").
    """
    sentences = _split_sentences(text)
    results = []

    # Build author lookup: all possible forms → clean name
    author_lookup = _build_author_lookup(known_authors)

    for sentence in sentences:
        sentence = sentence.strip()
        if not sentence:
            continue

        # Try to find one or more authors at the start of the sentence
        matched_authors, description = _extract_authors_and_desc(sentence, author_lookup)

        if matched_authors and description:
            description = description.strip(" ,;")
            if description:
                for author in matched_authors:
                    results.append((author, description))
        # If we couldn't match any known author, skip — don't invent rows

    return results


def _split_sentences(text: str) -> list[str]:
    """
    Split text into sentences on '. ' followed by a capital letter.
    Avoids splitting on abbreviations like 'Dr.', 'Prof.', dotted initials.
    """
    # Replace known abbreviations with placeholders temporarily
    text = re.sub(r'\b(Dr|Prof|Mr|Ms|Mrs|Fig|Eq|et al)\.\s+', r'\1__DOT__ ', text)
    # Split on '. ' followed by capital letter (author name start)
    sentences = re.split(r'\.\s+(?=[A-Z])', text)
    # Restore placeholders
    result = []
    for s in sentences:
        s = s.replace('__DOT__ ', '. ')
        result.append(s.strip())
    return [s for s in result if s]


def _build_author_lookup(known_authors: list[str]) -> dict[str, str]:
    """
    Build a dict mapping every recognisable form of each author name
    to the canonical clean name.

    Forms generated:
      "Jane Doe"    → "Jane Doe"
      "J.D."        → "Jane Doe"   (dotted initials)
      "JD"          → "Jane Doe"   (bare initials)
    """
    lookup = {}
    for author in known_authors:
        clean = clean_author_name(author)
        if not clean:
            continue
        lookup[clean.lower()] = clean
        lookup[clean.lower().replace(" ", "")] = clean

        # Generate initials
        words = clean.split()
        if len(words) >= 2:
            bare = "".join(w[0].upper() for w in words)
            dotted = ".".join(w[0].upper() for w in words) + "."
            lookup[bare.lower()] = clean
            lookup[dotted.lower()] = clean
            # Also without trailing dot
            lookup[dotted.lower().rstrip(".")] = clean

    return lookup


def _extract_authors_and_desc(
    sentence: str,
    author_lookup: dict[str, str],
) -> tuple[list[str], str]:
    """
    Try to find author name(s) at the start of a sentence.
    Returns (list_of_clean_author_names, description_text).
    Returns ([], "") if no known author found.

    Handles:
      "Jane Doe designed ..."           → (["Jane Doe"], "designed ...")
      "J.D. designed ..."               → (["Jane Doe"], "designed ...")
      "AB and CD wrote ..."             → (["Anil Bose", "Carla Diaz"], "wrote ...")
    """
    # Try to match one or more authors joined by " and "
    # Pattern: Author (and Author)* verb/description
    #
    # We try greedily: attempt to match the first N words as an author name,
    # then check if " and " follows and another author name follows that.

    tokens = sentence.split()
    if not tokens:
        return [], ""

    # Try to consume author names from the front
    matched_authors = []
    pos = 0  # word position

    while pos < len(tokens):
        # Try to match 1, 2, 3 words as an author name from current position
        found = False
        for length in [3, 2, 1]:
            if pos + length > len(tokens):
                continue
            candidate = " ".join(tokens[pos:pos + length])
            candidate_clean = _strip_name_suffix(candidate)
            key = candidate_clean.lower()
            # Try exact key
            if key in author_lookup:
                matched_authors.append(author_lookup[key])
                pos += length
                found = True
                break
            # Try without dots and spaces
            key2 = re.sub(r'[.\s]', '', key)
            if key2 in author_lookup:
                matched_authors.append(author_lookup[key2])
                pos += length
                found = True
                break

        if not found:
            break

        # Check if next token is "and" — if so, continue matching authors
        if pos < len(tokens) and tokens[pos].lower() == "and":
            pos += 1  # skip "and"
            continue
        else:
            break

    if not matched_authors:
        return [], ""

    description = " ".join(tokens[pos:]).strip()
    return matched_authors, description


def _strip_name_suffix(name: str) -> str:
    """Strip trailing superscript digits/symbols from a candidate author name.
    Preserves trailing dots so dotted initials like J.D. are kept intact."""
    return re.sub(r'[\d*†‡§\ufffd]+$', '', name).strip()


# ---------------------------------------------------------------------------
# Format: All-authors-equal
# ---------------------------------------------------------------------------

def _parse_all_authors_equally(text: str, known_authors: list[str]) -> list[tuple[str, str]]:
    """
    Handle: "All authors contributed equally to conceptualization,
             investigation and writing of the original draft."

    Extracts the listed roles and assigns each to every known author.
    """
    # Extract the roles after "to" or "in"
    m = re.search(
        r'\ball\s+authors?\b.{0,60}?\b(?:to|in)\s+(.+?)\.?\s*$',
        text, re.IGNORECASE
    )
    if not m:
        # Fallback: extract everything after "equally"
        m = re.search(r'\bequally\b\s+(.+?)\.?\s*$', text, re.IGNORECASE)

    if not m:
        return []

    roles_text = m.group(1)
    # Split on comma and "and"
    raw_roles = re.split(r',\s*|\s+and\s+', roles_text)
    raw_roles = [r.strip() for r in raw_roles if r.strip()]

    results = []
    clean_authors = [clean_author_name(a) for a in known_authors]
    clean_authors = [a for a in clean_authors if a]

    for author in clean_authors:
        for role in raw_roles:
            results.append((author, role))

    return results


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _normalize(text: str) -> str:
    """
    Normalize encoding artefacts and whitespace in raw paragraph text.
    The replacement character \ufffd often stands for an en-dash (–).
    """
    text = text.replace("\ufffd", "\u2013")  # replacement char → en-dash
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def _split_roles(roles_text: str) -> list[str]:
    """
    Split a role string on commas and semicolons.
    Each fragment is stripped and returned if non-empty.
    e.g. "Conceptualization, Writing – original draft; Methodology"
      → ["Conceptualization", "Writing – original draft", "Methodology"]
    """
    parts = re.split(r"[,;]", roles_text)
    return [p.strip(" .,;") for p in parts if p.strip(" .,;")]
