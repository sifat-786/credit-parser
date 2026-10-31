"""
mapper.py
---------
Maps a raw text description of a contribution to one or more official CRediT role names.

The mapping tries three approaches in order:
  1. Direct match — the text IS already a CRediT role (case/dash-insensitive)
  2. Keyword match — the text CONTAINS recognisable keywords for a role
  3. Fuzzy match  — the text is close enough to a role name (for typos/encodings)

Main entry point: build_mapper(credit_roles) → returns extract_all_roles function
  extract_all_roles(text) → list of matched role strings (or ["INVALID"])
"""

import re
from rapidfuzz import process, fuzz

INVALID = "INVALID"
FUZZY_THRESHOLD = 80  # minimum score (0–100) for fuzzy match acceptance


# ---------------------------------------------------------------------------
# Keyword mapping: prose phrases → CRediT role
# Keys use en-dash (–) as fetched from the site.
# Values are lowercase substrings to find inside the description text.
# ---------------------------------------------------------------------------
ROLE_KEYWORDS: dict[str, list[str]] = {
    "Conceptualization": [
        "conceptuali", "conceived", "designed the study", "research idea",
        "research design", "formulated the", "hypothesis", "overall design",
    ],
    "Data curation": [
        "data curation", "curated", "data management", "managed data",
        "organised data", "organized data", "annotated", "data cleaning",
        "cleaned data",
    ],
    "Formal analysis": [
        "formal analysis", "statistical analysis", "data analysis",
        "analyzed the data", "analysed the data", "performed the analysis",
        "performed the statistical", "conducted the analysis", "ran the analysis",
    ],
    "Funding acquisition": [
        "funding acquisition", "acquired funding", "obtained funding",
        "secured funding", "acquired the funding",
    ],
    "Investigation": [
        "investigation", "investigated", "collected data", "data collection",
        "conducted experiment", "performed experiment", "gathered data",
        "carried out experiment", "field work",
    ],
    "Methodology": [
        "methodology", "designed the method", "developed the method",
        "designed the protocol", "experimental design",
    ],
    "Project administration": [
        "project administration", "administered", "project management",
        "coordinated", "overall management", "managed the project",
        "administered the project",
    ],
    "Resources": [
        "provided resources", "provided materials", "provided equipment",
        "supplied materials", "provided facilities", "provided the",
    ],
    "Software": [
        "software", "developed the code", "wrote the code", "programmed",
        "implemented the", "scripted", "coding", "analysis code",
        "developed the analysis code",
    ],
    "Supervision": [
        "supervision", "supervised", "oversaw", "oversight", "mentored",
        "supervised the", "led the team",
    ],
    "Validation": [
        "validation", "validated", "verified", "reproducibility",
        "checked results", "confirmed results", "quality control",
    ],
    "Visualization": [
        "visualization", "visuali", "created figures", "made figures",
        "prepared figures", "generated figures", "plotted", "created the figures",
        "created the image", "made the figures",
    ],
    "Writing \u2013 original draft": [
        "original draft", "wrote the manuscript", "wrote the initial draft",
        "wrote the paper", "wrote the article", "drafted the manuscript",
        "drafted the paper", "wrote the first draft", "prepared the manuscript",
        "wrote initial", "writing of the original draft",
        "wrote and edited",  # partial — also maps to review & editing below
    ],
    "Writing \u2013 review & editing": [
        "review and editing", "review & editing", "edited the manuscript",
        "edited the paper", "revised the manuscript", "proofread",
        "wrote and edited",  # partial — shared with original draft above
        "reviewed the manuscript", "critical revision", "reviewed and edited",
        "edited the final text",
    ],
}


# ---------------------------------------------------------------------------
# Public factory
# ---------------------------------------------------------------------------

def build_mapper(credit_roles: list[str]):
    """
    Build and return an extract_all_roles function bound to the live-fetched roles.

    Args:
        credit_roles: list of role strings from credit.niso.org

    Returns:
        extract_all_roles(text: str) -> list[str]
            Maps a raw contribution description to a list of official CRediT
            role strings. Returns ["INVALID"] if no match is found.
    """
    # Normalised lookup: normalised_key → original fetched string
    norm_to_original: dict[str, str] = {_norm(r): r for r in credit_roles}
    norm_keys = list(norm_to_original.keys())

    # Align keyword map to actual fetched role names (handles encoding differences)
    keyword_map: dict[str, list[str]] = {}
    keyword_map_noart: dict[str, list[str]] = {}
    for kw_role, phrases in ROLE_KEYWORDS.items():
        matched = _match_to_fetched(kw_role, credit_roles)
        if matched:
            keyword_map[matched] = phrases
            # Article-free variants so "collected the data" matches "collected data".
            keyword_map_noart[matched] = [_strip_articles(p) for p in phrases]

    def _map_single(fragment: str) -> str:
        """Map one clean fragment to exactly one CRediT role or INVALID."""
        if not fragment.strip():
            return INVALID
        cleaned = _clean_fragment(fragment)
        normed = _norm(cleaned)

        # 1. Direct / near-direct match
        if normed in norm_to_original:
            return norm_to_original[normed]

        # 2. Keyword match (also try an article-stripped form of the text so
        #    phrasings like "collected the data" / "performed the experiments"
        #    match the keywords "collected data" / "performed experiment").
        lower = cleaned.lower()
        lower_noart = _strip_articles(lower)
        for role, phrases in keyword_map.items():
            if (any(phrase in lower for phrase in phrases)
                    or any(p in lower_noart for p in keyword_map_noart[role])):
                return role

        # 3. Fuzzy match
        result = process.extractOne(
            normed, norm_keys,
            scorer=fuzz.token_sort_ratio,
            score_cutoff=FUZZY_THRESHOLD,
        )
        if result:
            matched_norm, _score, _idx = result
            return norm_to_original[matched_norm]

        return INVALID

    def extract_all_roles(text: str) -> list[str]:
        """
        Extract all CRediT roles from a raw description string.

        The description might be:
          • A single role name:  "Conceptualization"
          • A prose phrase:      "designed the study and analyzed the data"
          • A parenthetical mix: "Developed the analysis code (Software)"

        Returns a deduplicated list of matched role strings,
        or ["INVALID"] if nothing matches.
        """
        if not text or not text.strip():
            return [INVALID]

        # Step 1: check for parenthetical role names — extract them first
        # e.g. "Developed the analysis code (Software) and checked results (Validation)"
        paren_roles = []
        text_without_parens = text
        for m in re.finditer(r'\(([^)]+)\)', text):
            inner = m.group(1).strip()
            role = _map_single(inner)
            if role != INVALID:
                paren_roles.append(role)
        if paren_roles:
            # Also check the text outside parentheses for additional roles
            text_without_parens = re.sub(r'\([^)]+\)', '', text).strip()
            extra = _map_text(text_without_parens, _map_single)
            all_roles = paren_roles + [r for r in extra if r not in paren_roles]
            all_roles = _dedupe(all_roles)
            return all_roles if all_roles else [INVALID]

        # Step 2: map the whole text
        roles = _map_text(text, _map_single)
        roles = _dedupe(roles)
        return roles if roles else [INVALID]

    return extract_all_roles


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _map_text(text: str, map_single) -> list[str]:
    """
    Extract all roles from a description by combining two passes and unioning them:
    1. Whole text as one role — catches role names/phrases that legitimately
       contain " and " (e.g. "reviewed and edited the manuscript").
    2. Each piece after splitting on " and " — catches multi-role conjunctions
       (e.g. "acquired funding and administered the project").

    Both passes contribute, so a second conjoined role is never dropped just
    because the first clause already matched.
    """
    roles = []

    # Pass 1: whole text.
    role = map_single(text)
    if role != INVALID:
        roles.append(role)
        # "wrote and edited" is shorthand for both writing roles.
        lower = text.lower()
        if "wrote and edited" in lower or "write and edit" in lower:
            for r in ["Writing \u2013 original draft", "Writing \u2013 review & editing"]:
                if r not in roles:
                    roles.append(r)
    # Pass 2: split on " and " and try each piece.
    # Avoid splitting the role name "Writing – review and editing".
    protected = re.sub(r'review\s+and\s+editing', 'review__AND__editing', text, flags=re.IGNORECASE)
    pieces = re.split(r'\s+and\s+', protected)
    if len(pieces) > 1:
        for piece in pieces:
            piece = piece.replace('__AND__', ' and ').strip()
            r = map_single(piece)
            if r != INVALID and r not in roles:
                roles.append(r)

    return roles


def _clean_fragment(text: str) -> str:
    """
    Light cleaning of a role fragment before matching:
    - Strip trailing qualifiers (– lead, – equal, – supporting)
    - Strip parenthetical qualifiers like (lead)
    - Strip trailing punctuation
    """
    # Strip parenthetical qualifiers
    text = re.sub(r'\s*\(\s*(?:lead|equal|supporting|co-\w+)\s*\)\s*$', '', text, flags=re.IGNORECASE)
    # Strip dash qualifiers: "Conceptualization – lead" or "Conceptualization - lead"
    text = re.sub(
        r'\s*[\u2013\u2014-]\s*(?:lead|equal|supporting|senior|junior|co[-\s]?\w+)\s*$',
        '', text, flags=re.IGNORECASE
    )
    text = text.strip(' .,;:')
    text = re.sub(r'\s+', ' ', text)
    return text


def _strip_articles(text: str) -> str:
    """Remove the articles 'the', 'a', 'an' so phrasings differing only by an
    article still match (e.g. 'collected the data' -> 'collected data')."""
    return re.sub(r'\b(?:the|a|an)\s+', '', text, flags=re.IGNORECASE)


def _norm(text: str) -> str:
    """
    Canonical form for comparison:
    - lowercase
    - all dashes → plain hyphen
    - \ufffd (replacement char) stripped
    - & → and
    - extra whitespace collapsed
    """
    text = text.replace('\ufffd', '-')
    text = re.sub(r'[\u2013\u2014\u2012\u2010]', '-', text)
    text = text.replace('&', 'and')
    text = text.lower().strip()
    text = re.sub(r'\s+', ' ', text)
    return text


def _match_to_fetched(kw_role: str, fetched: list[str]) -> str | None:
    """
    Find which fetched role name corresponds to a keyword-map entry.
    (They may differ in dash encoding etc.)
    """
    kw_norm = _norm(kw_role)
    for f in fetched:
        if _norm(f) == kw_norm:
            return f
    # Fuzzy fallback
    result = process.extractOne(
        kw_norm,
        [_norm(f) for f in fetched],
        scorer=fuzz.token_sort_ratio,
        score_cutoff=85,
    )
    if result:
        _, _, idx = result
        return fetched[idx]
    return None


def _dedupe(lst: list[str]) -> list[str]:
    seen = set()
    result = []
    for item in lst:
        if item not in seen:
            seen.add(item)
            result.append(item)
    return result
