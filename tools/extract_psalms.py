#!/usr/bin/env python3
"""
Extract BCP Psalms from Word documents and write to YAML.

One-time extraction tool.  The resulting YAML file at
bulletin/data/bcp_texts/psalms.yaml becomes the source of truth.

Usage:
    python3 tools/extract_psalms.py

Reads:
    Word Document Versions of Psalms/Psalms 1 - 75 from BCP.doc
    Word Document Versions of Psalms/Psalms 76 - 150 from BCP.doc

Writes:
    bulletin/data/bcp_texts/psalms.yaml
"""

import re
import subprocess
import sys
from pathlib import Path

import yaml

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DOC_DIR = PROJECT_ROOT / "source_documents" / "Word Document Versions of Psalms"
OUTPUT_PATH = PROJECT_ROOT / "bulletin" / "data" / "bcp_texts" / "psalms.yaml"

DOC_FILES = [
    DOC_DIR / "Psalms 1 - 75 from BCP.doc",
    DOC_DIR / "Psalms 76 - 150 from BCP.doc",
]

# ---------------------------------------------------------------------------
# Psalm 119 section names (Hebrew alphabet)
# ---------------------------------------------------------------------------
PSALM_119_SECTIONS = {
    "Aleph", "Beth", "Gimel", "Daleth", "He", "Waw", "Zayin", "Heth",
    "Teth", "Yodh", "Kaph", "Lamedh", "Mem", "Nun", "Samekh", "Ayin",
    "Pe", "Sadhe", "Qoph", "Resh", "Shin", "Taw",
}

# ---------------------------------------------------------------------------
# Regex patterns
# ---------------------------------------------------------------------------

# "23   Dominus regit me" — psalm header (number + 2+ literal spaces + Latin)
# Must use [ ] not \s so "1 \t..." (space+tab = verse start) doesn't false-match.
# Allow optional leading space (e.g., " 6   Domine, ne in furore").
RE_PSALM_HEADER = re.compile(r"^ ?(\d{1,3})[ ]{2,}(.+)$")

# "18" alone on a line — multi-part psalm number
RE_PSALM_NUMBER_ONLY = re.compile(r"^(\d{1,3})$")

# "Part I   Diligam te, Domine." — part header within multi-part psalm
RE_PART_HEADER = re.compile(r"^Part\s+([IVX]+)[ ]{2,}(.+)$")

# "Psalm 18: Part II   Et retribuet mihi" — continuation part header
RE_PSALM_PART_HEADER = re.compile(
    r"^Psalm\s+\d+:\s*Part\s+([IVX]+)[ ]{2,}(.+)$"
)

# "Aleph   Beati immaculati" — Psalm 119 section headers
RE_SECTION_HEADER = re.compile(r"^([A-Z][a-z]+)[ ]{2,}(.+)$")

# "1\tThe LORD is my shepherd; *" — verse start (number + optional space + tab)
# Some verses use "1 \t" (with a space before the tab) instead of "1\t"
RE_VERSE_START = re.compile(r"^(\d{1,3}) ?\t(.*)$")

# " \tI shall not be in want." — second-half line (space + tab)
RE_SECOND_HALF_SPACE_TAB = re.compile(r"^ \t(.*)$")

# "\tand guides me along..." — second-half line (tab only, no space)
RE_SECOND_HALF_TAB_ONLY = re.compile(r"^\t(.*)$")

# Day headers: "Fifth Day: Morning Prayer", "Twenty-fourth Day: Evening Prayer"
RE_DAY_HEADER = re.compile(
    r"^(?:First|Second|Third|Fourth|Fifth|Sixth|Seventh|Eighth|Ninth|"
    r"Tenth|Eleventh|Twelfth|Thirteenth|Fourteenth|Fifteenth|Sixteenth|"
    r"Seventeenth|Eighteenth|Nineteenth|Twentieth|Twenty).+:\s+"
    r"(?:Morning|Evening)\s+Prayer$"
)

# Book divisions: "Book One", "Book Two", etc.
RE_BOOK_HEADER = re.compile(r"^Book\s+(One|Two|Three|Four|Five)$")

# Page artifacts at end of file
RE_PAGE_ARTIFACT = re.compile(r"PAGE\s+\d+|^Psalms\s+PAGE")


# ---------------------------------------------------------------------------
# Text conversion
# ---------------------------------------------------------------------------

def convert_doc_to_text(doc_path: Path) -> str:
    """Convert a .doc file to plain text via macOS textutil."""
    result = subprocess.run(
        ["textutil", "-convert", "txt", "-stdout", str(doc_path)],
        capture_output=True, text=True, check=True,
    )
    return result.stdout


def normalize_text(text: str) -> str:
    """Normalize smart characters to plain equivalents."""
    text = text.replace("\u2018", "\u2018")   # keep left single quote
    text = text.replace("\u2019", "\u2019")   # keep right single quote
    text = text.replace("\u201c", "\u201c")   # keep left double quote
    text = text.replace("\u201d", "\u201d")   # keep right double quote
    text = text.replace("\u2011", "-")         # non-breaking hyphen → hyphen
    return text


# ---------------------------------------------------------------------------
# Parsing
# ---------------------------------------------------------------------------

def _roman_to_int(roman: str) -> int:
    """Convert a Roman numeral string to int."""
    vals = {"I": 1, "V": 5, "X": 10}
    total = 0
    prev = 0
    for ch in reversed(roman.upper()):
        v = vals.get(ch, 0)
        if v < prev:
            total -= v
        else:
            total += v
        prev = v
    return total


class PsalmParser:
    """State-machine parser for textutil output of BCP Psalm .doc files."""

    def __init__(self):
        self.psalms: dict[int, dict] = {}

        # Current psalm state
        self.psalm_num: int | None = None
        self.psalm_latin: str = ""
        self.psalm_parts: dict[int, dict] = {}
        self.psalm_sections: dict[str, dict] = {}  # Psalm 119 only
        self.verses: dict[int, dict] = {}

        # Current verse state
        self.verse_num: int | None = None
        self.first_half_lines: list[str] = []
        self.second_half_lines: list[str] = []
        self.seen_asterisk: bool = False

        # Are we inside psalm content (vs. preamble text)?
        self.in_psalms: bool = False

    # -- Flush helpers --

    def _flush_verse(self):
        """Assemble accumulated lines into a verse dict and store it."""
        if self.verse_num is None:
            return

        first_half = " ".join(self.first_half_lines).strip()
        # Strip trailing asterisk
        if first_half.endswith(" *"):
            first_half = first_half[:-2].rstrip()
        elif first_half.endswith("*"):
            first_half = first_half[:-1].rstrip()

        cleaned_second = [ln.strip() for ln in self.second_half_lines if ln.strip()]

        self.verses[self.verse_num] = {
            "first_half": first_half,
            "second_half": cleaned_second,
        }

        self.verse_num = None
        self.first_half_lines = []
        self.second_half_lines = []
        self.seen_asterisk = False

    def _flush_psalm(self):
        """Store the current psalm and reset state."""
        self._flush_verse()
        if self.psalm_num is not None and self.verses:
            entry: dict = {
                "latin": self.psalm_latin,
                "verses": dict(sorted(self.verses.items())),
            }
            if self.psalm_parts:
                entry["parts"] = dict(sorted(self.psalm_parts.items()))
            if self.psalm_sections:
                entry["sections"] = self.psalm_sections

            self.psalms[self.psalm_num] = entry

        self.psalm_num = None
        self.psalm_latin = ""
        self.psalm_parts = {}
        self.psalm_sections = {}
        self.verses = {}

    # -- Line classification --

    def _is_skip_line(self, line: str) -> bool:
        """Return True if this line is non-psalm content to skip."""
        if not line or line == ".":
            return True
        if RE_DAY_HEADER.match(line):
            return True
        if RE_BOOK_HEADER.match(line):
            return True
        if RE_PAGE_ARTIFACT.search(line):
            return True
        if line.startswith("The Psalter") or line.startswith("Concerning"):
            return True
        if "HYPERLINK" in line or "@" in line:
            return True
        return False

    # -- Main parse method --

    def parse(self, text: str):
        """Parse the full text of one .doc file."""
        text = normalize_text(text)
        lines = text.split("\n")

        for raw_line in lines:
            # Strip form feeds
            line = raw_line.replace("\x0c", "").rstrip()

            if self._is_skip_line(line):
                continue

            # --- Psalm header: "23   Dominus regit me" ---
            m = RE_PSALM_HEADER.match(line)
            if m:
                num = int(m.group(1))
                latin = m.group(2).strip()

                # Only valid psalm numbers (1-150).  The regex requires
                # 2+ spaces between number and text, which distinguishes
                # headers from verse lines (which use tabs).
                if 1 <= num <= 150:
                    self._flush_psalm()
                    self.psalm_num = num
                    self.psalm_latin = latin
                    self.in_psalms = True
                    continue

            # --- Psalm number only: "18" (multi-part) ---
            m = RE_PSALM_NUMBER_ONLY.match(line)
            if m and self.in_psalms:
                num = int(m.group(1))
                if num <= 150:
                    self._flush_psalm()
                    self.psalm_num = num
                    continue

            # --- Part header: "Part I   Diligam te" ---
            m = RE_PART_HEADER.match(line)
            if m and self.psalm_num is not None:
                part_num = _roman_to_int(m.group(1))
                latin = m.group(2).strip()
                self.psalm_parts[part_num] = {"latin": latin}
                if part_num == 1:
                    self.psalm_latin = latin
                continue

            # --- Psalm Part header: "Psalm 18: Part II   ..." ---
            m = RE_PSALM_PART_HEADER.match(line)
            if m and self.psalm_num is not None:
                part_num = _roman_to_int(m.group(1))
                latin = m.group(2).strip()
                self.psalm_parts[part_num] = {"latin": latin}
                continue

            # --- Psalm 119 section header: "Aleph   Beati immaculati" ---
            if self.psalm_num == 119:
                m = RE_SECTION_HEADER.match(line)
                if m and m.group(1) in PSALM_119_SECTIONS:
                    self._flush_verse()
                    section_name = m.group(1)
                    latin = m.group(2).strip()
                    # Use Aleph's latin as the psalm-level latin
                    if section_name == "Aleph":
                        self.psalm_latin = latin
                    # Record what verse number this section starts at
                    # (will be updated when we see the first verse)
                    self.psalm_sections[section_name] = {"latin": latin}
                    continue

            if not self.in_psalms:
                continue

            # --- Verse start: "1\tThe LORD is my shepherd; *" ---
            m = RE_VERSE_START.match(line)
            if m:
                self._flush_verse()
                self.verse_num = int(m.group(1))
                verse_text = m.group(2)

                # Update Psalm 119 section start_verse if applicable
                if self.psalm_num == 119 and self.psalm_sections:
                    for sec_name in reversed(list(self.psalm_sections)):
                        sec = self.psalm_sections[sec_name]
                        if "start_verse" not in sec:
                            sec["start_verse"] = self.verse_num
                        break

                if " *" in verse_text or verse_text.endswith("*"):
                    self.seen_asterisk = True
                    # Text before asterisk is first half
                    before_star = verse_text.rsplit(" *", 1)[0] if " *" in verse_text else verse_text.rstrip("*").rstrip()
                    self.first_half_lines.append(before_star + " *")
                else:
                    self.first_half_lines.append(verse_text)
                continue

            # --- Second-half line: " \t..." or "\t..." (no verse number) ---
            m = RE_SECOND_HALF_SPACE_TAB.match(line)
            if not m:
                m = RE_SECOND_HALF_TAB_ONLY.match(line)
            if m and self.verse_num is not None:
                text_content = m.group(1)
                if self.seen_asterisk:
                    self.second_half_lines.append(text_content)
                else:
                    # Tab-indented but before asterisk — unusual but possible
                    # Treat as continuation of first half
                    if " *" in text_content or text_content.endswith("*"):
                        self.seen_asterisk = True
                    self.first_half_lines.append(text_content)
                continue

            # --- Continuation line (no indent): part of previous line ---
            if self.verse_num is not None and line and not line[0].isspace():
                if self.seen_asterisk:
                    # Continuation of second half
                    self.second_half_lines.append(line)
                else:
                    # Continuation of first half
                    if " *" in line or line.endswith("*"):
                        self.seen_asterisk = True
                    self.first_half_lines.append(line)
                continue

        # Flush final psalm
        self._flush_psalm()


def extract_psalms() -> dict[int, dict]:
    """Extract all psalms from both .doc files."""
    parser = PsalmParser()

    for doc_path in DOC_FILES:
        if not doc_path.exists():
            print(f"Error: {doc_path} not found", file=sys.stderr)
            sys.exit(1)
        print(f"  Converting {doc_path.name}...")
        text = convert_doc_to_text(doc_path)
        print(f"  Parsing {doc_path.name} ({len(text)} chars)...")
        parser.parse(text)

    return dict(sorted(parser.psalms.items()))


# ---------------------------------------------------------------------------
# YAML output
# ---------------------------------------------------------------------------

def write_yaml(psalms: dict[int, dict], output_path: Path):
    """Write the psalms dict to a YAML file."""
    output_path.parent.mkdir(parents=True, exist_ok=True)

    # Custom representer to keep verse dicts readable
    class PsalmDumper(yaml.SafeDumper):
        pass

    # Force multi-line strings when they contain newlines
    def str_representer(dumper, data):
        if "\n" in data:
            return dumper.represent_scalar("tag:yaml.org,2002:str", data, style="|")
        return dumper.represent_scalar("tag:yaml.org,2002:str", data)

    PsalmDumper.add_representer(str, str_representer)

    with open(output_path, "w", encoding="utf-8") as f:
        f.write("# BCP Psalter - All 150 Psalms\n")
        f.write("# Parsed from 'Psalms 1-75 from BCP.doc' and 'Psalms 76-150 from BCP.doc'\n")
        f.write("# Verse numbers and asterisk positions are encoded in the structure\n")
        f.write("# but not printed in the bulletin.\n\n")
        yaml.dump(psalms, f, Dumper=PsalmDumper, default_flow_style=False,
                  allow_unicode=True, sort_keys=False, width=120)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    print("Extracting BCP Psalms...")
    psalms = extract_psalms()
    print(f"  Extracted {len(psalms)} psalms")

    # Quick stats
    total_verses = sum(len(p["verses"]) for p in psalms.values())
    print(f"  Total verses: {total_verses}")

    # Spot checks
    for pnum in [1, 23, 51, 119, 150]:
        if pnum in psalms:
            p = psalms[pnum]
            nv = len(p["verses"])
            print(f"  Psalm {pnum}: {nv} verses, latin={p['latin']!r}")

    write_yaml(psalms, OUTPUT_PATH)
    print(f"\nWritten to {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
