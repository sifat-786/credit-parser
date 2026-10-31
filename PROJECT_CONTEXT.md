# Project Context & Progress Log

**Project:** CRediT Contributor Roles Parser (QE Hackathon)  
**Author:** SHAIK SIFAT AHMED (CS23B1035)  
**Location:** `C:\TNQ\credit-parser`  
**Created:** 2026-10-31  

---

## 1. What Has Been Done So Far

### Step 1: Requirements Exploration & File Inspection
- Unpacked `manuscripts.zip` (20 Word `.docx` documents) into inspection directory.
- Ran diagnostic scripts across all 20 documents to inspect:
  - Document headings and author byline positioning.
  - Presence, location, and varied syntax of contributorship statements.
  - Boundary headers (e.g., `Acknowledgements`, `Funding`, `References`).
- Discovered specific tricky cases:
  - `merged_03.docx`, `merged_09.docx`, `merged_11.docx`: No contributions section at all.
  - `merged_08.docx`: Section heading present, but no content paragraphs (empty).
  - `merged_12.docx`: Corrupted / zero-byte package causing `PackageNotFoundError`.
  - `merged_13.docx`: False-positive trap — the phrase "Author Contributions" only appears inside a reference citation (`[10] Smith A. Author Contributions...`).
  - `merged_15.docx`: Academic honorifics (`Dr.`, `Prof.`) and superscript affiliation markers (`Jane Doe1*`, `Anil Bose2†`).
  - `merged_16.docx`: Role degree qualifiers (`Conceptualization – lead`, `Methodology – equal`).
  - `merged_17.docx`: Collective author declaration (`All authors contributed equally...`).
  - `merged_18.docx`: Explicitly invalid contribution (`G.H. made the tea`).
  - `merged_19.docx`: Multi-line prose statement with words like "funding" inside sentences.
  - `merged_20.docx`: Mixed parenthetical roles (`Developed the analysis code (Software)`).

### Step 2: Formulation & Submission of Scenario List
- Formulated an exhaustive testing taxonomy across 5 core areas.
- Formatted and generated the official submission Word document:
  - **Filename:** `scenario_list_CS23B1035.docx` (in `C:\TNQ\credit-parser\`)
  - **Content:** Candidate identification, project overview, 40+ scenario table with clear expected behavior, and key architectural decisions.
  - Successfully generated before the 30-minute hackathon deadline.

### Step 3: Module Implementation

#### 1. Dynamic Web Fetcher (`src/fetcher.py`)
- Purpose: Dynamically fetch the 14 official CRediT roles from `https://credit.niso.org/` at runtime (hardcoding strictly prohibited).
- Implementation:
  - HTTP GET request with a 15-second timeout.
  - Primary DOM parser extracts `<p>` tags located in the roles listing section.
  - Secondary fallback heuristics check `<h2>`, `<h3>`, `<li>`, and `<dt>` tags.
  - Filtering mechanism removes site navigation boilerplate while preserving valid single-word roles (such as "Resources").
  - Verification: Successfully extracts all 14 official taxonomy terms live.

#### 2. Document Ingestion & Section Reader (`src/reader.py`)
- Purpose: Safely read `.docx` files, identify author bylines, and extract the exact contribution section.
- Implementation:
  - Wrapped `python-docx` Document loader in exception guards to handle corrupted files (like `merged_12.docx`) gracefully without crashing.
  - Author byline parser: Extracts author names from early document paragraphs while stripping affiliation markers, numbers, and symbols.
  - Contributions section extractor:
    - Recognizes variants: `Author Contributions`, `Author Contributions Statement`, `CRediT Taxonomy`.
    - Citation false-positive defense: Rejects reference citations matching regex `^\[\d+\]` or year/volume formats.
    - Boundary detector: Stops extraction when encountering sections like `Acknowledgements`, `Funding`, `References`. Distinguishes standalone headers from prose containing the word "funding" (preventing early truncation in `merged_19.docx`).
  - Name cleaner: Strips academic titles (`Dr.`, `Prof.`) while preserving trailing punctuation in dotted initials (`J.D.`).

#### 3. Contribution Syntactic Parser (`src/parser.py`)
- Purpose: Turn extracted raw section text into `(author, role_fragment)` tuples.
- Implementation:
  - Dispatches to specialized sub-parsers:
    1. Parenthetical parser: Extracts author and roles matching `Name (Role1 – sub; Role2)`.
    2. Structured parser: Splits on author blocks `Name: Role1, Role2.` and decomposes delimiters (commas and semicolons).
    3. Multi-author / Conjunction parser: Handles compound subjects like `AB and CD wrote and edited the paper.`
    4. Collective parser: Handles `All authors contributed equally...` by mapping roles across all authors from the document byline.
    5. Prose parser: Extracts authors and action clauses from natural language sentences.

#### 4. Role Normalization & Fuzzy Mapper (`src/mapper.py`)
- Purpose: Map diverse descriptions to the live-fetched official 14 CRediT roles or mark them as `INVALID`.
- Implementation:
  - Implements a runtime factory `build_mapper(credit_roles)` bound to the live scraped role list.
  - Three-tier resolution pipeline:
    1. Direct match: Exact or case-insensitive string match with dash normalization (`\ufffd` and hyphens).
    2. Semantic keyword matching: Maps natural phrases (e.g., `designed the study` -> `Conceptualization`, `performed the statistical analysis` -> `Formal analysis`).
    3. Fuzzy matching: Uses RapidFuzz `token_sort_ratio` with an 80% threshold for typos and variations.
    4. Sub-role qualifier stripping: Strips qualifiers like `– lead`, `– equal`, `– supporting`.
    5. Failure output: Descriptions that do not match (e.g., `made the tea`) return `INVALID`.

#### 5. Pipeline CLI Orchestrator (`main.py`)
- Purpose: Command-line tool accepting `input_folder` and `output_csv`.
- Implementation:
  - Invokes `fetcher.py` at launch. If internet fails, switches to safe mode where all emitted roles are `INVALID`.
  - Loops deterministically through all `.docx` files in alphabetical order.
  - Emits rows conforming to schema: `file,author,credit_role`.
  - Ensures deduplication of `(file, author, credit_role)` tuples.

### Step 4: Verification & Bug Fixing
- Executed `python main.py sample_manuscripts output.csv`.
- Bug identified and resolved #1: Dotted author initials lost trailing period (`J.D.` became `J.D`). Fixed regex in `reader.py` and `parser.py`.
- Bug identified and resolved #2: `merged_19.docx` yielded 0 rows because the sentence "M.P. acquired funding..." triggered the `Funding` next-section detector. Fixed by checking paragraph length and standalone heading criteria.
- Current pipeline result: Exactly 51 rows written to `output.csv` with zero unhandled exceptions.

---

## 2. Current File Structure in `C:\TNQ\credit-parser`

```
C:\TNQ\credit-parser/
├── main.py                         # Complete & operational
├── output.csv                      # Generated output (51 rows, valid)
├── scenario_list_CS23B1035.docx    # Word document submission artifact
├── HACKATHON_HANDOVER.md           # Master handover document
├── PROJECT_CONTEXT.md              # Context & log (this file)
├── src/
│   ├── __init__.py                 # Package initializer
│   ├── fetcher.py                  # Dynamic web scraper (tested & working)
│   ├── reader.py                   # DOCX ingestion & boundary detection
│   ├── parser.py                   # Syntactic parsing (structured & prose)
│   └── mapper.py                   # Role mapping & RapidFuzz engine
├── tests/
│   └── __init__.py                 # Tests package
└── sample_manuscripts/             # 20 provided test files (merged_01.docx - merged_20.docx)
```

---

## 3. What Remains to be Completed

1. **Automated Test Suite (`tests/test_credit.py`):**
   - Unit tests for `fetcher`, `reader`, `parser`, and `mapper`.
   - Integration tests verifying outputs for all 20 sample manuscripts.
2. **Test Report Generator (`test_report.csv`):**
   - Runner script to execute the test suite and output `test_report.csv` with columns: `checks,input,expected,actual,result`.
3. **Execution Shell Script (`run.sh`):**
   - Top-level bash script accepting `input_folder` and `output.csv`.
   - Must run `pip install` for `requests`, `beautifulsoup4`, `python-docx`, `rapidfuzz` on Ubuntu.
   - Must use LF line endings.
4. **Documentation (`README.md`):**
   - Clear markdown documentation explaining installation, how to run tests, and project structure.
5. **Chat Transcript Export:**
   - Save or export Claude chat transcript into repository before final commit.

