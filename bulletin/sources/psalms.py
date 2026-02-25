"""
BCP Psalm lookup by reference.

Loads the pre-parsed YAML psalter and retrieves specific verses
by liturgical reference strings like "Psalm 72:1-7,10-14".
"""

import re
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path
from typing import Optional

import yaml

_PSALMS_PATH = Path(__file__).parent.parent / "data" / "bcp_texts" / "psalms.yaml"


@lru_cache(maxsize=1)
def _load_psalms() -> dict:
    """Load and cache the psalms YAML file."""
    with open(_PSALMS_PATH, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class PsalmVerse:
    """A single psalm verse with its half-verse components."""
    number: int
    first_half: str
    second_half: list[str]


@dataclass
class PsalmSelection:
    """A selection of psalm verses for bulletin rendering."""
    psalm_number: int
    latin: str
    verses: list[PsalmVerse]

    def to_lines(self) -> list[str]:
        """Convert to a list of formatted lines for the bulletin.

        Each verse becomes a single string with the first half on the
        first line and each second-half line tab-indented on subsequent
        lines, joined by newlines.  No verse numbers or asterisks.

        The rendering function (_add_psalm in word_of_god.py) splits
        on the newlines and uses run.add_break() to keep each verse
        in a single "Psalm"-styled paragraph with hanging indent.
        """
        lines = []
        for verse in self.verses:
            parts = []
            if verse.first_half:
                parts.append(verse.first_half)
            for sh_line in verse.second_half:
                parts.append(f"\t{sh_line}")
            if parts:
                lines.append("\n".join(parts))
        return lines


# ---------------------------------------------------------------------------
# Reference parsing
# ---------------------------------------------------------------------------

def parse_psalm_reference(reference: str) -> tuple[int, list[tuple[int, Optional[str]]]]:
    """Parse a psalm reference string into psalm number and verse specs.

    Args:
        reference: e.g., "Psalm 72:1-7,10-14" or "Psalm 147:1-12, 21c"
                   or "Psalm 51:1-13" or "Psalm 116:1,10-17"

    Returns:
        (psalm_number, verse_specs) where verse_specs is a list of
        (verse_num, suffix) tuples.  suffix is None, "a", "b", or "c".
        Ranges are expanded: "1-3" → [(1,None), (2,None), (3,None)].
        An empty list means "all verses" (entire psalm).
    """
    # Strip any trailing rubric like " responsively" or " (read in unison)"
    ref_clean = re.sub(
        r"\s*[\(]?(?:responsively|in unison|read .*|antiphonally)[\)]?.*$",
        "", reference, flags=re.IGNORECASE,
    )

    match = re.match(r"Psalm\s+(\d+)(?::(.+))?", ref_clean.strip())
    if not match:
        raise ValueError(f"Cannot parse psalm reference: {reference!r}")

    psalm_num = int(match.group(1))
    verse_part = match.group(2)

    if not verse_part:
        return psalm_num, []  # entire psalm

    # Parse comma-separated segments: "1-7,10-14" or "1,10-17" or "1-12, 21c"
    verse_specs: list[tuple[int, Optional[str]]] = []
    segments = re.split(r",\s*", verse_part.strip())

    for segment in segments:
        segment = segment.strip()

        # Range: "1-7" or "10-14" or "1a-3"
        range_match = re.match(r"(\d+)([a-c])?-(\d+)([a-c])?", segment)
        if range_match:
            start = int(range_match.group(1))
            start_suffix = range_match.group(2)
            end = int(range_match.group(3))
            end_suffix = range_match.group(4)

            for v in range(start, end + 1):
                suffix = None
                if v == start and start_suffix:
                    suffix = start_suffix
                elif v == end and end_suffix:
                    suffix = end_suffix
                verse_specs.append((v, suffix))
            continue

        # Single verse: "1" or "21c"
        single_match = re.match(r"(\d+)([a-c])?", segment)
        if single_match:
            verse_specs.append((
                int(single_match.group(1)),
                single_match.group(2),
            ))

    return psalm_num, verse_specs


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def get_psalm(reference: str) -> PsalmSelection:
    """Look up psalm verses by reference string.

    Args:
        reference: e.g., "Psalm 72:1-7,10-14" or "Psalm 23"

    Returns:
        PsalmSelection with the requested verses.

    Raises:
        ValueError if psalm not found.
    """
    psalms = _load_psalms()
    psalm_num, verse_specs = parse_psalm_reference(reference)

    psalm_data = psalms.get(psalm_num)
    if psalm_data is None:
        raise ValueError(f"Psalm {psalm_num} not found in psalter data")

    all_verses = psalm_data.get("verses", {})
    latin = psalm_data.get("latin", "")

    if not verse_specs:
        # Return all verses in order
        selected = [
            PsalmVerse(
                number=vnum,
                first_half=v["first_half"],
                second_half=v["second_half"],
            )
            for vnum, v in sorted(all_verses.items())
        ]
    else:
        selected = []
        for vnum, suffix in verse_specs:
            v = all_verses.get(vnum)
            if v is None:
                continue  # Skip missing verses silently

            first_half = v["first_half"]
            second_half = list(v["second_half"])

            # a/b/c suffixes select partial verses:
            #   "a" = first half only
            #   "b" = second half only (first line)
            #   "c" = second half remaining (or all of it if only one line)
            if suffix == "a":
                selected.append(PsalmVerse(vnum, first_half, []))
            elif suffix == "b":
                selected.append(PsalmVerse(vnum, "", second_half[:1]))
            elif suffix == "c":
                selected.append(PsalmVerse(
                    vnum, "", second_half[1:] or second_half
                ))
            else:
                selected.append(PsalmVerse(vnum, first_half, second_half))

    return PsalmSelection(
        psalm_number=psalm_num,
        latin=latin,
        verses=selected,
    )
