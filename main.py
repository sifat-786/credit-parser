"""
main.py
-------
Entry point for the CRediT role parser.

Usage:
    python main.py <input_folder> <output_csv>

Reads every .docx file in input_folder, extracts author contribution roles,
maps them to the official CRediT taxonomy (fetched live from credit.niso.org),
and writes a CSV with columns: file, author, credit_role.
"""

import csv
import os
import sys

from src.fetcher import fetch_credit_roles
from src.reader import read_manuscript
from src.parser import parse_contributions
from src.mapper import build_mapper, INVALID


def process_folder(input_folder: str, output_csv: str) -> list[dict]:
    """
    Main processing pipeline.

    Returns the list of row dicts written to the CSV (for testability).
    """
    # ── Step 1: fetch CRediT roles live ─────────────────────────────────────
    credit_roles = fetch_credit_roles()
    fetch_failed = credit_roles is None

    if fetch_failed:
        print("[main] WARNING: Live fetch failed — all credit_role values will be INVALID.")
        credit_roles = []

    extract_all_roles = build_mapper(credit_roles)

    # ── Step 2: find .docx files ─────────────────────────────────────────────
    if not os.path.isdir(input_folder):
        print(f"[main] ERROR: Input folder not found: {input_folder}")
        sys.exit(1)

    docx_files = sorted(
        f for f in os.listdir(input_folder)
        if f.lower().endswith(".docx")
    )

    if not docx_files:
        print(f"[main] No .docx files found in {input_folder}")

    # ── Step 3: process each file ─────────────────────────────────────────────
    all_rows = []

    for filename in docx_files:
        filepath = os.path.join(input_folder, filename)
        file_rows = _process_file(
            filepath, filename, extract_all_roles, fetch_failed
        )
        all_rows.extend(file_rows)

    # ── Step 4: write CSV ────────────────────────────────────────────────────
    os.makedirs(os.path.dirname(os.path.abspath(output_csv)), exist_ok=True)
    with open(output_csv, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["file", "author", "credit_role"])
        writer.writeheader()
        writer.writerows(all_rows)

    print(f"[main] Done — wrote {len(all_rows)} rows to {output_csv}")
    return all_rows


def _process_file(
    filepath: str,
    filename: str,
    extract_all_roles,
    fetch_failed: bool,
) -> list[dict]:
    """
    Process one .docx file and return its output rows.
    Never raises — returns [] on any error.
    """
    try:
        authors, section_paras = read_manuscript(filepath)

        # No contributions section → zero rows (not INVALID)
        if not section_paras:
            print(f"[main] {filename}: no contributions section found.")
            return []

        author_role_pairs = parse_contributions(section_paras, authors)

        if not author_role_pairs:
            print(f"[main] {filename}: contributions section found but no pairs parsed.")
            return []

        seen = set()
        rows = []

        for author, raw_role in author_role_pairs:
            if fetch_failed:
                roles = [INVALID]
            else:
                roles = extract_all_roles(raw_role)

            for credit_role in roles:
                key = (filename, author, credit_role)
                if key not in seen:
                    seen.add(key)
                    rows.append({
                        "file": filename,
                        "author": author,
                        "credit_role": credit_role,
                    })

        print(f"[main] {filename}: {len(rows)} rows.")
        return rows

    except Exception as e:
        print(f"[main] ERROR processing {filename}: {e}")
        return []


# ── CLI entry point ──────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python main.py <input_folder> <output_csv>")
        sys.exit(1)

    process_folder(sys.argv[1], sys.argv[2])

