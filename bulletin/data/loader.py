"""
Loads YAML data files for BCP texts, proper prefaces, blessings, POP forms, and staff.
"""

import os
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


def load_staff() -> dict:
    return _load_yaml("staff.yaml")


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
