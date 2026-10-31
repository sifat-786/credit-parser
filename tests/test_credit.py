"""
test_credit.py
--------------
Automated test suite for the CRediT contributor-role parser.

Design notes
------------
* The suite is DATA-DRIVEN: every check records one row
  (checks, input, expected, actual, result) and all rows are written to
  ``test_report.csv`` in the project root. Running the module exits with a
  non-zero status if any check fails, so it works in CI as well.

* The suite runs FULLY OFFLINE and DETERMINISTICALLY:
    - Fetcher tests mock the HTTP layer (mocked HTML / mocked network errors).
    - Mapper / parser / pipeline tests build the mapper from a FIXED list of
      the 14 official role names used purely as a *test fixture*.
  This fixed list is NOT a taxonomy hardcode in the product: ``fetcher.py``
  still scrapes credit.niso.org live. A test simply needs a known-correct
  answer to compare against, and must not depend on live network.

Run it with:
    python3 tests/test_credit.py
"""

import csv
import os
import sys
import tempfile
from unittest import mock

# Make the project root importable regardless of the current directory.
PROJ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, PROJ)

import requests  # noqa: E402  (used for the mocked exception types)
from docx import Document  # noqa: E402

from src import fetcher, reader, parser, mapper  # noqa: E402
import main  # noqa: E402

SAMPLES = os.path.join(PROJ, "sample_manuscripts")
REPORT_PATH = os.path.join(PROJ, "test_report.csv")

# The 14 official role names, exactly as the live site presents them
# (en-dash in "Writing – ...", ampersand in "review & editing").
# TEST FIXTURE ONLY — the product fetches these dynamically.
CANONICAL_ROLES = [
    "Conceptualization", "Data curation", "Formal analysis", "Funding acquisition",
    "Investigation", "Methodology", "Project administration", "Resources",
    "Software", "Supervision", "Validation", "Visualization",
    "Writing – original draft", "Writing – review & editing",
]

RESULTS = []  # list of dicts: checks, input, expected, actual, result


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _fmt(value) -> str:
    """Canonical string form so expected/actual compare and read cleanly."""
    if isinstance(value, (list, tuple, set)):
        return " | ".join(sorted(str(v) for v in value))
    return str(value)


def check(name: str, input_desc, expected, actual) -> None:
    """Record one check row (expected == actual decides PASS/FAIL)."""
    exp, act = _fmt(expected), _fmt(actual)
    RESULTS.append({
        "checks": name,
        "input": _fmt(input_desc),
        "expected": exp,
        "actual": act,
        "result": "PASS" if exp == act else "FAIL",
    })


def _blank_docx() -> str:
    """Create a temporary .docx with only whitespace paragraphs (Scenario 14)."""
    path = os.path.join(tempfile.gettempdir(), "blank_test.docx")
    doc = Document()
    doc.add_paragraph("")
    doc.add_paragraph("   ")
    doc.save(path)
    return path


def _html_with(tag: str, roles) -> str:
    """Wrap each role in <tag> plus some navigation boilerplate."""
    body = "<nav><p>Home</p><p>About</p></nav>"
    body += "".join(f"<{tag}>{r}</{tag}>" for r in roles)
    return f"<html><body>{body}</body></html>"


# ---------------------------------------------------------------------------
# Area 1: Dynamic web scraping & resilience (Scenarios 1-8)
# ---------------------------------------------------------------------------

def test_fetcher():
    roles_ascii = [r.replace("–", "-") for r in CANONICAL_ROLES]

    # S1: happy-path HTML with 14 <p> roles.
    parsed = fetcher._parse_roles_from_html(_html_with("p", roles_ascii))
    check("S1 happy-path fetch parses 14 roles", "14 <p> roles",
          14, len(parsed or []))

    # S2: total network outage -> None.
    with mock.patch("src.fetcher.requests.get",
                    side_effect=requests.ConnectionError("dns")):
        check("S2 network outage -> None", "ConnectionError", None,
              fetcher.fetch_credit_roles())

    # S3: HTTP error status -> None.
    resp = mock.Mock()
    resp.raise_for_status.side_effect = requests.HTTPError("500")
    with mock.patch("src.fetcher.requests.get", return_value=resp):
        check("S3 HTTP 500 -> None", "HTTPError", None,
              fetcher.fetch_credit_roles())

    # S4: socket timeout -> None.
    with mock.patch("src.fetcher.requests.get",
                    side_effect=requests.Timeout("timeout")):
        check("S4 timeout -> None", "Timeout", None,
              fetcher.fetch_credit_roles())

    # S5: DOM restructuring, roles only in <li>.
    parsed = fetcher._parse_roles_from_html(_html_with("li", roles_ascii))
    check("S5 DOM restructuring (<li>)", "14 <li> roles", 14, len(parsed or []))

    # S6: partial taxonomy (<10 roles) -> None.
    parsed = fetcher._parse_roles_from_html(_html_with("p", roles_ascii[:3]))
    check("S6 partial taxonomy -> None", "3 roles", None, parsed)

    # S7: replacement char in a role name is normalised to en-dash.
    glitched = roles_ascii[:13] + ["Writing � review & editing"]
    parsed = fetcher._parse_roles_from_html(_html_with("p", glitched)) or []
    check("S7 \\ufffd normalised in fetched role",
          "Writing \\ufffd review & editing",
          True, "Writing – review & editing" in parsed)

    # S8: keyword-collision — valid single-word role kept, boilerplate dropped.
    check("S8 'Resources' kept", "Resources", True,
          fetcher._looks_like_role("Resources"))
    check("S8 'About' dropped", "About", False,
          fetcher._looks_like_role("About"))


# ---------------------------------------------------------------------------
# Area 2 & 3: File ingestion + section identification (Scenarios 9-25)
# ---------------------------------------------------------------------------

def test_reader():
    # S9: valid standard document yields authors + a section.
    authors, section = reader.read_manuscript(os.path.join(SAMPLES, "merged_01.docx"))
    check("S9 valid docx -> authors found", "merged_01.docx", True, len(authors) > 0)
    check("S9 valid docx -> section found", "merged_01.docx", True, len(section) > 0)

    # S10: corrupted DOCX -> ([], []) without raising.
    check("S10 corrupted docx -> empty", "merged_12.docx", ([], []),
          reader.read_manuscript(os.path.join(SAMPLES, "merged_12.docx")))

    # S13: non-existent file -> ([], []).
    check("S13 missing file -> empty", "NOPE.docx", ([], []),
          reader.read_manuscript(os.path.join(SAMPLES, "NOPE.docx")))

    # S14: blank document -> no section paragraphs.
    _, section = reader.read_manuscript(_blank_docx())
    check("S14 blank docx -> no section", "blank.docx", [], section)

    # S17-S20: heading recognition variants.
    check("S17 'Author Contributions' heading", "Author Contributions",
          True, reader._is_contributions_heading("Author Contributions"))
    check("S18 '...Statement' variant", "Author Contributions Statement",
          True, reader._is_contributions_heading("Author Contributions Statement"))
    check("S19 'CRediT Taxonomy' header", "CRediT Taxonomy",
          True, reader._is_contributions_heading("CRediT Taxonomy"))
    check("S20 casing 'AUTHOR CONTRIBUTIONS'", "AUTHOR CONTRIBUTIONS",
          True, reader._is_contributions_heading("AUTHOR CONTRIBUTIONS"))

    # S21: file with no section -> empty section list.
    _, section = reader.read_manuscript(os.path.join(SAMPLES, "merged_03.docx"))
    check("S21 no section -> empty", "merged_03.docx", [], section)

    # S22: heading present but body empty -> empty section list.
    _, section = reader.read_manuscript(os.path.join(SAMPLES, "merged_08.docx"))
    check("S22 empty section -> empty", "merged_08.docx", [], section)

    # S23: citation false-positive is NOT treated as a heading.
    citation = "[10] Smith A. Author Contributions: a practical guide. J Pub Ethics. 2019;4:1-9."
    check("S23 citation rejected as heading", citation,
          False, reader._is_contributions_heading(citation))

    # S25: boundary detection — real heading stops, prose with 'funding' does not.
    check("S25 'Acknowledgements' is next-section", "Acknowledgements",
          True, reader._is_next_section_heading("Acknowledgements"))
    prose = ("The funding body had no role in the design of the study, and a "
             "colleague wrote the original protocol before the project began.")
    check("S25 prose containing 'funding' is NOT next-section", prose,
          False, reader._is_next_section_heading(prose))


# ---------------------------------------------------------------------------
# Area 4: Author extraction & name normalisation (Scenarios 26-34)
# ---------------------------------------------------------------------------

def test_authors():
    check("S27 dotted initials preserved", "J.D.", "J.D.",
          reader.clean_author_name("J.D."))
    check("S29 honorific stripped", "Dr. Jane Doe", "Jane Doe",
          reader.clean_author_name("Dr. Jane Doe"))
    check("S29 'Prof.' stripped", "Prof. Anil Bose", "Anil Bose",
          reader.clean_author_name("Prof. Anil Bose"))
    check("S30 superscript affiliation stripped", "Anil Bose2†", "Anil Bose",
          reader.clean_author_name("Anil Bose2†"))
    check("S31 \\ufffd stripped from name", "Anil Bose2�", "Anil Bose",
          reader.clean_author_name("Anil Bose2�"))
    check("S34 trailing dot kept on initials", "A.B.", "A.B.",
          reader.clean_author_name("A.B."))


# ---------------------------------------------------------------------------
# Area 5: Role extraction, mapping & normalisation (Scenarios 35-43)
# ---------------------------------------------------------------------------

def test_mapper():
    extract = mapper.build_mapper(CANONICAL_ROLES)
    OD = "Writing – original draft"
    RE = "Writing – review & editing"

    check("S35 exact match", "Conceptualization",
          ["Conceptualization"], extract("Conceptualization"))
    check("S36 all-caps", "FORMAL ANALYSIS",
          ["Formal analysis"], extract("FORMAL ANALYSIS"))
    check("S36 lowercase", "data curation",
          ["Data curation"], extract("data curation"))
    check("S37 ASCII hyphen", "Writing - Original Draft", [OD],
          extract("Writing - Original Draft"))
    check("S37 en-dash", OD, [OD], extract(OD))
    check("S38 \\ufffd dash", "Writing � original draft", [OD],
          extract("Writing � original draft"))
    check("S39 qualifier '- lead' stripped", "Conceptualization – lead",
          ["Conceptualization"], extract("Conceptualization – lead"))
    check("S39 qualifier '- supporting' stripped",
          "Writing – original draft – supporting", [OD],
          extract("Writing – original draft – supporting"))
    check("S40 parenthetical roles",
          "Developed the analysis code (Software) and checked results (Validation)",
          ["Software", "Validation"],
          extract("Developed the analysis code (Software) and checked results (Validation)"))
    check("S41 'designed the study'", "designed the study",
          ["Conceptualization"], extract("designed the study"))
    check("S41 'acquired funding'", "acquired funding",
          ["Funding acquisition"], extract("acquired funding"))
    check("S41 'provided resources'", "provided resources",
          ["Resources"], extract("provided resources"))
    check("S41 'supervised the project'", "supervised the project",
          ["Supervision"], extract("supervised the project"))
    check("S42 'wrote and edited the paper' -> both", "wrote and edited the paper",
          [OD, RE], extract("wrote and edited the paper"))
    # Regression guard for the multi-role conjunction bug fixed in _map_text:
    check("S42 conjunction 'acquired funding and administered the project'",
          "acquired funding and administered the project",
          ["Funding acquisition", "Project administration"],
          extract("acquired funding and administered the project"))
    check("S42 conjunction 'designed the study and analyzed the data'",
          "designed the study and analyzed the data",
          ["Conceptualization", "Formal analysis"],
          extract("designed the study and analyzed the data"))
    check("S43 unmappable -> INVALID", "made the tea",
          ["INVALID"], extract("made the tea"))


# ---------------------------------------------------------------------------
# Area 4/5: Parser behaviour (Scenarios 32, 33, structured)
# ---------------------------------------------------------------------------

def test_parser():
    # S33: compound authors "AB and CD wrote and edited the paper."
    pairs = parser.parse_contributions(
        ["AB and CD wrote and edited the paper. EF supervised the project."],
        ["Anil Bose", "Carla Diaz", "Erin Fox"],
    )
    authors = {a for a, _ in pairs}
    check("S33 compound authors resolved", "AB and CD ...",
          {"Anil Bose", "Carla Diaz", "Erin Fox"}, authors)

    # S32: "All authors contributed equally to ..." assigns roles to every author.
    pairs = parser.parse_contributions(
        ["All authors contributed equally to conceptualization, investigation "
         "and writing of the original draft."],
        ["Jane Doe", "Anil Bose"],
    )
    authors = {a for a, _ in pairs}
    check("S32 all-authors-equal covers every author", "All authors ...",
          {"Jane Doe", "Anil Bose"}, authors)

    # Structured colon format splits roles.
    pairs = parser.parse_contributions(
        ["Jane Doe: Conceptualization, Methodology."], ["Jane Doe"])
    roles = {r for _, r in pairs}
    check("Structured 'Name: a, b' splits roles", "Jane Doe: Conceptualization, Methodology.",
          {"Conceptualization", "Methodology"}, roles)


# ---------------------------------------------------------------------------
# Area 6: CSV generation, dedup & pipeline (Scenarios 44-50)
# ---------------------------------------------------------------------------

# Corrected expected per-file row counts (after the conjunction fix).
EXPECTED_COUNTS = {
    "merged_01.docx": 4, "merged_02.docx": 5, "merged_03.docx": 0,
    "merged_04.docx": 5, "merged_05.docx": 4, "merged_06.docx": 3,
    "merged_07.docx": 5, "merged_08.docx": 0, "merged_09.docx": 0,
    "merged_10.docx": 4, "merged_11.docx": 0, "merged_12.docx": 0,
    "merged_13.docx": 0, "merged_14.docx": 2, "merged_15.docx": 3,
    "merged_16.docx": 3, "merged_17.docx": 6, "merged_18.docx": 2,
    "merged_19.docx": 4, "merged_20.docx": 7,
}


def test_pipeline():
    # Patch the live fetch so the pipeline runs offline against the fixture roles.
    out = os.path.join(tempfile.gettempdir(), "pipeline_test.csv")
    with mock.patch("main.fetch_credit_roles", return_value=list(CANONICAL_ROLES)):
        rows = main.process_folder(SAMPLES, out)

    # Per-file row counts (Scenarios 21/22/24/32/33/40/44 integration).
    counts = {}
    for r in rows:
        counts[r["file"]] = counts.get(r["file"], 0) + 1
    for fname, expected in EXPECTED_COUNTS.items():
        check(f"Pipeline row count {fname}", fname, expected, counts.get(fname, 0))

    # S44: author with 3 roles yields 3 rows (merged_02 Raj Patel).
    raj = sum(1 for r in rows if r["file"] == "merged_02.docx" and r["author"] == "Raj Patel")
    check("S44 one row per author-role (Raj Patel x3)", "merged_02 Raj Patel", 3, raj)

    # S45: no duplicate (file, author, role) triples anywhere.
    triples = [(r["file"], r["author"], r["credit_role"]) for r in rows]
    check("S45 no duplicate rows", "all rows", len(triples), len(set(triples)))

    # S49: every credit_role is a canonical role or INVALID.
    allowed = set(CANONICAL_ROLES) | {"INVALID"}
    bad = [r["credit_role"] for r in rows if r["credit_role"] not in allowed]
    check("S49 roles are canonical or INVALID", "all rows", [], bad)

    # S46: commas in a field are RFC-4180 quoted.
    qpath = os.path.join(tempfile.gettempdir(), "quote_test.csv")
    with open(qpath, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["file", "author", "credit_role"])
        w.writeheader()
        w.writerow({"file": "x.docx", "author": "Doe, Jane", "credit_role": "Conceptualization"})
    line = open(qpath, encoding="utf-8").read().splitlines()[1]
    check("S46 comma field is quoted", "author='Doe, Jane'",
          'x.docx,"Doe, Jane",Conceptualization', line)

    # S47: output directory is created automatically.
    nested = os.path.join(tempfile.gettempdir(), "credit_nested_%d" % os.getpid(),
                          "deep", "out.csv")
    with mock.patch("main.fetch_credit_roles", return_value=list(CANONICAL_ROLES)):
        main.process_folder(SAMPLES, nested)
    check("S47 output dir auto-created", nested, True, os.path.exists(nested))

    # S50: files are processed in sorted order.
    order = [r["file"] for r in rows]
    check("S50 deterministic sorted order", "row file order",
          order, sorted(order))


# ---------------------------------------------------------------------------
# Runner
# ---------------------------------------------------------------------------

def main_run() -> int:
    for fn in (test_fetcher, test_reader, test_authors, test_mapper,
               test_parser, test_pipeline):
        try:
            fn()
        except Exception as e:  # a crash in a group is itself a failure row
            RESULTS.append({
                "checks": f"{fn.__name__} crashed",
                "input": "-", "expected": "no exception",
                "actual": f"{type(e).__name__}: {e}", "result": "FAIL",
            })

    with open(REPORT_PATH, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["checks", "input", "expected", "actual", "result"])
        w.writeheader()
        w.writerows(RESULTS)

    passed = sum(1 for r in RESULTS if r["result"] == "PASS")
    failed = len(RESULTS) - passed
    print(f"\n{'='*60}")
    print(f"TOTAL: {len(RESULTS)}   PASS: {passed}   FAIL: {failed}")
    print(f"Report written to: {REPORT_PATH}")
    if failed:
        print("\nFAILURES:")
        for r in RESULTS:
            if r["result"] == "FAIL":
                print(f"  - {r['checks']}: expected [{r['expected']}] got [{r['actual']}]")
    print(f"{'='*60}")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main_run())
