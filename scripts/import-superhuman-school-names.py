#!/usr/bin/env python3
import json
import re
import sys
import unicodedata
from pathlib import Path

import pdfplumber


ROOT = Path(__file__).resolve().parent.parent
CATALOG_PATH = ROOT / "web" / "school-rankings.json"
RANKING_SOURCE_PATH = ROOT / "data" / "7sage-rankings-2026.json"
OUTPUT_PATH = ROOT / "data" / "superhuman-school-names-2026.json"

MANUAL_MATCHES = {
    "University of California—Berkeley": "Berkeley",
    "University of California—Irvine": "UC - Irvine",
    "University of Notre Dame": "Notre Dame",
    "Washington and Lee University": "Washington & Lee",
    "Arizona State University": "Arizona State",
    "University of California—Davis": "UC - Davis",
    "University of Florida (Levin)": "U Florida",
    "George Mason University": "George Mason",
    "University of Colorado—Boulder": "Colorado - Boulder",
    "Yeshiva University (Cardozo)": "Cardozo",
    "University of California (Hastings)": "UC Law San Francisco",
    "Pennsylvania State - Dickinson Law": "Penn State",
    "Seton Hall University": "Seton Hall",
    "Loyola Marymount University—Los Angeles": "Loyola Marymount - LA",
    "Rutgers University (merged)": "Rutgers",
    "Texas A&M University": "Texas A&M",
    "St. John's University": "St. John's",
    "University of New Hampshire": "New Hampshire",
    "University of Arkansas, Fayetteville": "U Arkansas - Fayetteville",
    "University of South Carolina": "South Carolina",
    "Indiana University - Indianapolis": "IU McKinney",
    "Wayne State University": "Wayne State",
    "Albany Law School Of Union University": "Albany",
    "Catholic University Of America": "Catholic University",
    "Cleveland State University": "Cleveland",
    "Santa Clara University": "Santa Clara",
    "Texas Tech University": "Texas Tech",
    "University of South Dakota": "South Dakota",
    "University of Arkansas, Little Rock": "U Arkansas - Little Rock",
    "Widener- commonwealth": "Widener - Pennsylvania",
    "Atlanta's John Marshall Law School": "John Marshall",
    "University of Detroit Mercy": "Detroit Mercy",
    "Faulkner University": "Jones",
    "Florida A&M University": "Florida A&M",
    "University of Illinois Chicago School of Law (UIC Law)": "Illinois - Chicago",
    "University of the Pacific (Mcgeorge)": "Pacific (Mcgeorge)",
    "Northern Kentucky University": "Northern Kentucky",
    "Nova Southeastern University": "Nova Southeastern",
    "Oklahoma City University": "Oklahoma City",
    "Roger Williams University": "Roger Williams",
    "South Texas College Of Law—Houston": "South Texas",
    "Southern Illinois University—Carbondale": "SIU - Carbondale",
    "St. Mary's University": "St. Mary's",
    "St. Thomas University (Florida)": "St. Thomas",
    "Western Michigan University (Cooley)": "Western Michigan",
    "Western New England University": "Western New England",
    "Western State College Of Law": "Western State",
    "Widener University—Delaware": "Widener-Delaware",
    "Inter American University Of Puerto Rico": "Inter American",
    "Pontifical Catholic University Of P.R.": "Pontifical Catholic",
    "University of Puerto Rico": "Puerto Rico",
    "Ohio State University": "Ohio State",
    "Wake Forest University": "Wake Forest",
    "Michigan State University": "Michigan State",
    "University of Buffalo — SUNY": "SUNY Buffalo",
    "Barry University": "Barry Law",
    "Texas Southern University": "Texas Southern",
    "Touro College": "Touro",
}


def normalize(value):
    text = unicodedata.normalize("NFD", str(value or ""))
    text = "".join(char for char in text if unicodedata.category(char) != "Mn")
    text = text.replace("&", " and ")
    return re.sub(r"[^a-z0-9]+", " ", text.lower()).strip()


def extract_names(pdf_path):
    names = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables() or []:
                for row in table:
                    if len(row) < 4 or not row[1]:
                        continue
                    name = re.sub(r"\s+", " ", str(row[1])).strip()
                    if name == "School":
                        continue
                    if name and name not in names:
                        names.append(name)
    return names


def main():
    if len(sys.argv) != 2:
        raise SystemExit("Usage: import-superhuman-school-names.py <source.pdf>")
    pdf_path = Path(sys.argv[1]).expanduser().resolve()
    catalog = json.loads(CATALOG_PATH.read_text())
    source_rows = json.loads(RANKING_SOURCE_PATH.read_text())
    source_names = [re.sub(r"\s+", " ", row["name"]).strip() for row in source_rows]
    lookup = {}
    for index, school in enumerate(catalog["schools"]):
        for value in [school["name"], *school.get("aliases", [])]:
            lookup.setdefault(normalize(value), index)

    entries = []
    unmatched = []
    used_indexes = set()
    manual_matches = {normalize(name): source_name for name, source_name in MANUAL_MATCHES.items()}
    for canonical_name in extract_names(pdf_path):
        source_name = manual_matches.get(normalize(canonical_name))
        index = source_names.index(source_name) if source_name in source_names else lookup.get(normalize(canonical_name))
        if index is None:
            unmatched.append(canonical_name)
            continue
        if index in used_indexes:
            raise SystemExit(f"Multiple PDF names matched {source_names[index]}")
        used_indexes.add(index)
        entries.append({"sourceName": source_names[index], "canonicalName": canonical_name})

    allowed_unmatched = {"Golden Gate University"}
    unexpected = [name for name in unmatched if name not in allowed_unmatched]
    if unexpected:
        raise SystemExit(f"Unmatched PDF names: {unexpected}")

    output = {
        "source": "Consultant Hub · Law Schools - Superhuman Docs.pdf",
        "sourceCreated": "2026-08-06",
        "note": "Canonical display names only. Numeric ranks remain sourced from the ranking catalog.",
        "entries": entries,
        "unrankedOrUnmatched": unmatched,
    }
    OUTPUT_PATH.write_text(json.dumps(output, indent=2, ensure_ascii=False) + "\n")
    print(f"Wrote {len(entries)} canonical names; retained {len(source_names) - len(entries)} ranking-source names.")
    if unmatched:
        print(f"Not added without a numeric ranking match: {', '.join(unmatched)}")


if __name__ == "__main__":
    main()
