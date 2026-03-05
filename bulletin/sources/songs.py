"""
Song lyrics lookup from the extracted YAML data files.

Songs are identified in the Google Sheet's Service Music by either:
  - Hymnal number: "#93 Angels From the Realms of Glory"
  - Title only: "Everlasting God"
  - Special identifiers: "Gloria", "Doxology"

This module loads the unified songs.yaml file and provides lookup by various
identifiers, filtering by service when applicable.

Songs in the YAML may have a `services` field:
  - services: "9am"  -> only for 9am service
  - services: "11am" -> only for 11am service
  - no services field -> available for both services
"""

import re
from pathlib import Path
from typing import Optional

import yaml


DATA_DIR = Path(__file__).parent.parent / "data" / "hymns"
SONGS_FILE = DATA_DIR / "songs.yaml"

# Cache: all songs loaded once
_all_songs: Optional[list[dict]] = None


def _load_all_songs() -> list[dict]:
    """Load all songs from the unified YAML file."""
    global _all_songs
    if _all_songs is None:
        if not SONGS_FILE.exists():
            _all_songs = []
        else:
            with open(SONGS_FILE, encoding="utf-8") as f:
                _all_songs = yaml.safe_load(f) or []
    return _all_songs


def _get_songs(service: str = "9am") -> list[dict]:
    """Get songs available for the given service.

    A song is available for a service if:
      - It has no 'services' field (available for both), or
      - Its 'services' field matches the requested service
    """
    # Normalize: "9 am" -> "9am", "11 am" -> "11am"
    service = service.replace(" ", "")
    all_songs = _load_all_songs()
    return [
        s for s in all_songs
        if "services" not in s or s["services"] == service
    ]


def lookup_song(identifier: str, service: str = "9am",
                _in_fallback: bool = False) -> Optional[dict]:
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

    # Try identifier/alias match (e.g., "Kyrie" → "Lord have mercy upon us")
    for song in songs:
        for alias in song.get("identifiers", []):
            if alias.lower() == id_lower:
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

    # Cross-service fallback: if not found in the primary service,
    # try the other service's songs (lyrics are shared when needed)
    if not _in_fallback:
        svc = service.replace(" ", "")  # normalize "11 am" → "11am"
        fallback = "9am" if svc == "11am" else "11am"
        result = lookup_song(identifier, service=fallback, _in_fallback=True)
        if result:
            return result

    return None


def clear_cache():
    """Clear the song cache (useful for testing)."""
    global _all_songs
    _all_songs = None
