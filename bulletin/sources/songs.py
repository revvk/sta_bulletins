"""
Song lyrics lookup from the extracted YAML data files.

Songs are identified in the Google Sheet's Service Music by either:
  - Hymnal number: "#93 Angels From the Realms of Glory"
  - Title only: "Everlasting God"
  - Special identifiers: "Gloria", "Doxology"

This module loads the YAML files and provides lookup by various identifiers.
"""

import re
from pathlib import Path
from typing import Optional

import yaml


DATA_DIR = Path(__file__).parent.parent / "data" / "hymns"


def _load_songs(service: str = "9am") -> list[dict]:
    """Load songs from YAML file for the given service."""
    filepath = DATA_DIR / f"songs_{service}.yaml"
    if not filepath.exists():
        return []
    with open(filepath, encoding="utf-8") as f:
        return yaml.safe_load(f) or []


# Cache loaded songs
_song_cache: dict[str, list[dict]] = {}


def _get_songs(service: str = "9am") -> list[dict]:
    """Get songs with caching."""
    if service not in _song_cache:
        _song_cache[service] = _load_songs(service)
    return _song_cache[service]


def lookup_song(identifier: str, service: str = "9am") -> Optional[dict]:
    """Look up a song by various identifier formats.

    The identifier may be:
      - A hymnal number with title: "#93 Angels From the Realms of Glory"
      - Just a hymnal number: "#93"
      - A song title: "Everlasting God"
      - A partial title match

    Returns a song dict with keys: title, hymnal_number, hymnal_name,
    tune_name, sections. Returns None if not found.
    """
    songs = _get_songs(service)

    # Try to extract hymnal number from identifier
    num_match = re.match(r'#(\d+)', identifier.strip())
    if num_match:
        hymnal_num = num_match.group(1)
        for song in songs:
            if song.get("hymnal_number") == hymnal_num:
                return song

    # Try exact title match (case-insensitive)
    id_lower = identifier.strip().lower()
    # Strip hymnal ref from identifier for title matching
    title_part = re.sub(r'^#\d+\s*', '', identifier).strip()

    for song in songs:
        if song["title"].lower() == id_lower:
            return song
        if title_part and song["title"].lower() == title_part.lower():
            return song

    # Try starts-with match
    for song in songs:
        if song["title"].lower().startswith(id_lower):
            return song
        if title_part and song["title"].lower().startswith(title_part.lower()):
            return song

    # Try substring match
    for song in songs:
        if id_lower in song["title"].lower():
            return song

    # Also try the 11am songs if not found in the primary service
    if service != "11am":
        result = lookup_song(identifier, service="11am")
        if result:
            return result

    return None


def clear_cache():
    """Clear the song cache (useful for testing)."""
    _song_cache.clear()
