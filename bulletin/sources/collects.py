"""
Collect of the Day lookup from BCP data.

Loads the pre-extracted collects YAML and matches a liturgical day
title to its collect text.  Matching is fuzzy: it tries an exact match
first, then falls back to case-insensitive substring matching against
the YAML keys.

For Ordinary Time (the season after Pentecost), the BCP assigns each
Sunday a "Proper" number keyed to a fixed calendar date — e.g.
Proper 8 is "Week of the Sunday closest to June 29."  When a title
like "Third Sunday after Pentecost" doesn't directly match a YAML key,
we compute the Proper number from the target date.
"""

from datetime import date
from functools import lru_cache
from pathlib import Path

import yaml

_DATA_PATH = Path(__file__).parent.parent / "data" / "bcp_texts" / "collects.yaml"

# BCP pp. 228-236: each Proper is assigned to the week of the Sunday
# closest to a fixed date.  The (month, day) anchor for each Proper:
_PROPER_ANCHORS: dict[int, tuple[int, int]] = {
    5:  (6, 1),    # June 1
    6:  (6, 8),    # June 8
    7:  (6, 15),   # June 15
    8:  (6, 22),   # June 22
    9:  (6, 29),   # June 29
    10: (7, 6),    # July 6
    11: (7, 13),   # July 13
    12: (7, 20),   # July 20
    13: (7, 27),   # July 27
    14: (8, 3),    # August 3
    15: (8, 10),   # August 10
    16: (8, 17),   # August 17
    17: (8, 24),   # August 24
    18: (8, 31),   # August 31
    19: (9, 7),    # September 7
    20: (9, 14),   # September 14
    21: (9, 21),   # September 21
    22: (9, 28),   # September 28
    23: (10, 5),   # October 5
    24: (10, 12),  # October 12
    25: (10, 19),  # October 19
    26: (10, 26),  # October 26
    27: (11, 2),   # November 2
    28: (11, 9),   # November 9
}


@lru_cache(maxsize=1)
def _load_collects() -> dict[str, str]:
    """Load the collects YAML (cached after first call)."""
    with open(_DATA_PATH, encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def proper_from_date(target: date) -> int | None:
    """Return the BCP Proper number for *target*, or None if outside range.

    Each Proper's anchor date is the fixed calendar date listed in the
    BCP.  The Proper used on any given Sunday is the one whose anchor
    is closest to that Sunday.
    """
    year = target.year
    best_proper = None
    best_distance = None
    for proper_num, (month, day) in _PROPER_ANCHORS.items():
        anchor = date(year, month, day)
        distance = abs((target - anchor).days)
        if best_distance is None or distance < best_distance:
            best_distance = distance
            best_proper = proper_num
    return best_proper


def get_collect(liturgical_title: str, target_date: date | None = None) -> str | None:
    """Look up the Collect of the Day for a liturgical title.

    Args:
        liturgical_title: The title from the schedule, e.g.
            "Second Sunday in Lent", "Proper 15", "Christmas Day".
        target_date: The date of the service.  Used to compute the
            Proper number when the title doesn't match directly
            (e.g. "Third Sunday after Pentecost").

    Returns:
        The collect text, or None if no match is found.
    """
    collects = _load_collects()
    title = liturgical_title.strip()

    # 1. Exact match
    if title in collects:
        return collects[title]

    # 2. Case-insensitive exact match
    title_lower = title.lower()
    for key, val in collects.items():
        if key.lower() == title_lower:
            return val

    # 3. Fuzzy: title is a substring of a key, or vice versa
    #    e.g. "Second Sunday in Lent" matches "Second Sunday in Lent"
    #    or "Pentecost" matches "Day of Pentecost / Whitsunday"
    for key, val in collects.items():
        if title_lower in key.lower() or key.lower() in title_lower:
            return val

    # 4. Date-based Proper lookup for Ordinary Time
    #    Titles like "Third Sunday after Pentecost" or any unmatched
    #    Ordinary Time Sunday can be resolved via the calendar date.
    if target_date is not None:
        proper_num = proper_from_date(target_date)
        if proper_num is not None:
            proper_key = f"Proper {proper_num}"
            if proper_key in collects:
                return collects[proper_key]

    return None
