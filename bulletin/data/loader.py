"""
Loads YAML data files for BCP texts, proper prefaces, blessings, POP forms,
reading introductions, psalm reader instructions, hymnal index, and staff.
"""

import re
from pathlib import Path
from functools import lru_cache

import yaml


_DATA_DIR = Path(__file__).parent


@lru_cache(maxsize=None)
def _load_yaml(relative_path: str) -> dict:
    """Load and cache a YAML file relative to the data directory."""
    full_path = _DATA_DIR / relative_path
    with open(full_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def load_common_prayers() -> dict:
    return _load_yaml("bcp_texts/common_prayers.yaml")


def load_eucharistic_prayers() -> dict:
    return _load_yaml("bcp_texts/eucharistic_prayers.yaml")


def load_proper_prefaces() -> dict:
    return _load_yaml("bcp_texts/proper_prefaces.yaml")


def load_blessings() -> dict:
    return _load_yaml("bcp_texts/blessings.yaml")


def load_pop_forms() -> dict:
    return _load_yaml("prayers/pop_forms.yaml")


def load_placeholders() -> dict:
    return _load_yaml("placeholders.yaml")


def load_reading_introductions() -> dict:
    return _load_yaml("bcp_texts/reading_introductions.yaml")


def load_psalm_reader_instructions() -> dict:
    return _load_yaml("bcp_texts/psalm_reader_instructions.yaml")


def load_maundy_thursday() -> dict:
    """Load Maundy Thursday proper texts."""
    return _load_yaml("bcp_texts/maundy_thursday.yaml")


def load_good_friday() -> dict:
    """Load Good Friday proper texts."""
    return _load_yaml("bcp_texts/good_friday.yaml")


def load_great_litany() -> dict:
    """Load the Great Litany (BCP pp. 148-153)."""
    return _load_yaml("bcp_texts/great_litany.yaml")


def load_palm_sunday() -> dict:
    """Load Palm Sunday proper texts (Liturgy of the Palms, etc.)."""
    return _load_yaml("bcp_texts/palm_sunday.yaml")


def load_passion_gospel(year: str) -> dict:
    """Load the Passion Gospel for a given liturgical year (A, B, or C).

    Returns:
        Dict with 'reference', 'liturgical_year', and 'lines' keys.
        Each line has 'part' (speaker name) and 'text'.
    """
    return _load_yaml(f"bcp_texts/passion_gospel_{year.lower()}.yaml")


def load_hymnal_first_lines() -> dict:
    """Load the Hymnal 1982 index of first lines.

    Returns:
        Dict with int keys (hymn numbers 1-720) mapping to canonical
        first-line strings, plus a 'cross_references' key with alternate
        first lines.
    """
    return _load_yaml("hymnal_1982_first_lines.yaml")


def get_canonical_hymn_title(hymn_number: int) -> str | None:
    """Look up the canonical first line for a Hymnal 1982 hymn number.

    Args:
        hymn_number: Integer hymn number (1-720).

    Returns:
        The canonical first line, or None if the number isn't found.
    """
    data = load_hymnal_first_lines()
    return data.get(hymn_number)


def extract_book_name(reference: str) -> str:
    """Extract the book name from a scripture reference string.

    Examples:
        "Exodus 17:1-7"           → "Exodus"
        "1 Corinthians 10:1-13"   → "1 Corinthians"
        "Song of Solomon 2:1-7"   → "Song of Solomon"
    """
    m = re.match(r"^(.+?)\s+\d+", reference.strip())
    return m.group(1) if m else reference.strip()


def get_proper_preface_text(preface_key: str, option_key: str = None) -> str:
    """Look up the actual proper preface text by key and optional sub-key.

    Args:
        preface_key: Top-level key (e.g., "advent", "easter", "lords_day", "lent")
        option_key: Sub-key for multi-option prefaces (e.g., "of_god_the_father", "option_1")

    Returns:
        The preface text string.
    """
    prefaces = load_proper_prefaces()
    entry = prefaces.get(preface_key, {})

    if option_key:
        # Multi-option prefaces (lords_day, lent)
        sub = entry.get(option_key, {})
        if isinstance(sub, dict):
            return sub.get("text", "")
        return str(sub)
    else:
        # Single-option prefaces (advent, easter, etc.)
        if isinstance(entry, dict):
            return entry.get("text", "")
        return str(entry)


def get_preface_option_labels(preface_key: str) -> list[tuple[str, str]]:
    """Get labels for multi-option prefaces, for user prompting.

    Returns:
        List of (option_key, label) tuples.
    """
    prefaces = load_proper_prefaces()
    entry = prefaces.get(preface_key, {})

    if preface_key == "lords_day":
        return [
            (k, v.get("label", k))
            for k, v in entry.items()
            if isinstance(v, dict) and "text" in v
        ]
    elif preface_key == "lent":
        return [
            (k, v.get("label", k))
            for k, v in entry.items()
            if isinstance(v, dict) and "text" in v
        ]
    return []
