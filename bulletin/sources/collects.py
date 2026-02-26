"""
Collect of the Day lookup from BCP data.

Loads the pre-extracted collects YAML and matches a liturgical day
title to its collect text.  Matching is fuzzy: it tries an exact match
first, then falls back to case-insensitive substring matching against
the YAML keys.
"""

from functools import lru_cache
from pathlib import Path

import yaml

_DATA_PATH = Path(__file__).parent.parent / "data" / "bcp_texts" / "collects.yaml"


@lru_cache(maxsize=1)
def _load_collects() -> dict[str, str]:
    """Load the collects YAML (cached after first call)."""
    with open(_DATA_PATH, encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def get_collect(liturgical_title: str) -> str | None:
    """Look up the Collect of the Day for a liturgical title.

    Args:
        liturgical_title: The title from the schedule, e.g.
            "Second Sunday in Lent", "Proper 15", "Christmas Day".

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

    # 4. Try matching just the core words (ignore "The", ordinals, etc.)
    #    e.g. "Third Sunday after Pentecost" should match "Proper 8"
    #    This is harder — Propers don't map to ordinal Sundays without a
    #    calendar calculation.  For now, return None and let the caller
    #    handle the fallback.

    return None
