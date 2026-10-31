"""
fetcher.py
----------
Fetches the 14 official CRediT contributor roles from https://credit.niso.org/ at runtime.
Hardcoding the roles is explicitly prohibited by the problem statement.

Returns a list of role name strings on success, or None on failure.
The caller is responsible for deciding what to do when None is returned (write INVALID).
"""

import requests
from bs4 import BeautifulSoup

CREDIT_URL = "https://credit.niso.org/"
REQUEST_TIMEOUT = 15  # seconds
EXPECTED_ROLE_COUNT = 14


def fetch_credit_roles() -> list[str] | None:
    """
    Scrape the official CRediT role names from credit.niso.org.

    Returns:
        A list of role name strings (e.g. ["Conceptualization", "Data curation", ...])
        or None if the fetch or parse fails for any reason.
    """
    try:
        response = requests.get(CREDIT_URL, timeout=REQUEST_TIMEOUT)
        response.raise_for_status()
        roles = _parse_roles_from_html(response.text)
        if roles:
            print(f"[fetcher] Successfully fetched {len(roles)} CRediT roles.")
        return roles
    except requests.RequestException as e:
        print(f"[fetcher] Network error fetching {CREDIT_URL}: {e}")
        return None
    except Exception as e:
        print(f"[fetcher] Unexpected error: {e}")
        return None


def _parse_roles_from_html(html: str) -> list[str] | None:
    """
    Parse the CRediT role names out of the page HTML.

    The credit.niso.org page (as of 2026) lists each role as a standalone <p> tag
    between the "14 Contributor Roles" heading and the copyright footer.
    We find that section and extract short, title-cased paragraphs as role names.
    """
    soup = BeautifulSoup(html, "html.parser")
    roles = []

    # Primary strategy: roles appear as standalone <p> tags.
    # They sit between the h3 "CRediT's 14 Contributor Roles" and the footer.
    # Collect all <p> text and filter to those that look like role names.
    for p in soup.find_all("p"):
        text = p.get_text(strip=True)
        # Normalize encoding artefacts (replacement char from UTF-8 issues)
        text = text.replace("\ufffd", "\u2013")  # treat as en-dash
        if _looks_like_role(text):
            roles.append(text)

    if len(roles) >= 10:
        return _deduplicate(roles)

    # Fallback: try h2/h3 tags
    roles = []
    for tag in soup.find_all(["h2", "h3", "h4"]):
        text = tag.get_text(strip=True)
        if _looks_like_role(text):
            roles.append(text)

    if len(roles) >= 10:
        return _deduplicate(roles)

    # Fallback: any element with short role-like text
    roles = []
    for tag in soup.find_all(["li", "dt", "td", "span", "div"]):
        # Only leaf-ish nodes — skip if they contain sub-elements with text
        child_texts = [c.get_text(strip=True) for c in tag.children
                       if hasattr(c, 'get_text')]
        text = tag.get_text(strip=True)
        if _looks_like_role(text):
            roles.append(text)

    if len(roles) >= 10:
        return _deduplicate(roles)

    print(f"[fetcher] Warning: could not reliably parse roles from page "
          f"(found {len(roles)} candidates). Will return None.")
    return None


# Role names that are short phrases we'd expect to NOT be navigation or boilerplate.
# Used only to skip obvious non-role text.
_SKIP_WORDS = {
    "home", "about", "contact", "search", "menu", "login", "register",
    "subscribe", "download", "copyright", "privacy", "terms", "niso",
    "share", "powered", "wordpress", "facebook", "linkedin", "twitter",
    "news", "views", "adopters", "benefits", "origins",
    "welcome", "approved", "find", "summary", "table",
}


def _looks_like_role(text: str) -> bool:
    """
    Heuristic: a CRediT role name is a short phrase (1–6 words),
    starts with a capital letter, and is not navigation/boilerplate.
    Known roles are 1–4 words; "Writing – original draft" is the longest at 4.
    """
    if not text:
        return False
    # Strip trailing punctuation for length check
    clean = text.strip(".")
    words = clean.split()
    if len(words) < 1 or len(words) > 6:
        return False
    if len(clean) > 50:
        return False
    # Must start with a capital letter
    if not clean[0].isupper():
        return False
    # Skip obvious non-role words
    if clean.lower() in _SKIP_WORDS:
        return False
    first_word = words[0].lower().rstrip(".,;:")
    if first_word in _SKIP_WORDS:
        return False
    # Skip if it looks like a sentence (has a verb indicator)
    if any(w in clean.lower() for w in [" is ", " are ", " was ", " were ",
                                         " has ", " have ", "http", "www"]):
        return False
    return True


def _deduplicate(roles: list[str]) -> list[str]:
    """Remove duplicates while preserving order."""
    seen = set()
    result = []
    for r in roles:
        key = r.lower()
        if key not in seen:
            seen.add(key)
            result.append(r)
    return result
