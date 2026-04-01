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
HS_SONGS_FILE = DATA_DIR / "hidden_springs_songs.yaml"

# Cache: all songs loaded once
_all_songs: Optional[list[dict]] = None
_hs_songs: Optional[list[dict]] = None


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


def _normalize_service(service: str) -> str:
    """Normalize service identifiers for song filtering.

    Maps all service times to "9am" or "11am" for song lookup.
    Special weekday services (7 pm, etc.) use the 11am music pool.
    Sunrise services use the 9am pool (full lyrics, same song catalog).
    """
    svc = service.replace(" ", "").lower()
    if svc in ("9am", "sunrise"):
        return "9am"
    # 11am, 7pm, and any other special service → 11am pool
    return "11am"


def _get_songs(service: str = "9am") -> list[dict]:
    """Get songs available for the given service.

    A song is available for a service if:
      - It has no 'services' field (available for both), or
      - Its 'services' field matches the requested service
    """
    service = _normalize_service(service)
    all_songs = _load_all_songs()
    return [
        s for s in all_songs
        if "services" not in s or s["services"] == service
    ]


def _clean_identifier(identifier: str) -> str:
    """Strip parenthetical notes and hymnal refs for title matching.

    Transforms identifiers like:
      "Savior, like a shepherd lead us (H708+Bradbury)" → "Savior, like a shepherd lead us"
      "King of Love (no bridge)"                        → "King of Love"
      "Come, thou fount of every blessing H686"         → "Come, thou fount of every blessing"
      "Holy, holy, holy Lord S129 (Powell)"             → "Holy, holy, holy Lord"
    """
    clean = re.sub(r'\s*\([^)]*\)', '', identifier)   # strip (...)
    clean = re.sub(r'\s+[HS]\d+\b', '', clean)        # strip trailing H### / S###
    return clean.strip()


def lookup_song(identifier: str, service: str = "9am",
                _in_fallback: bool = False) -> Optional[dict]:
    """Look up a song by various identifier formats.

    The identifier may be:
      - A hymnal number with title: "#93 Angels From the Realms of Glory"
      - Just a hymnal number: "#93"
      - A song title: "Everlasting God"
      - A title with hymnal ref: "Come, thou fount H686"
      - A title with notes: "King of Love (no bridge)"
      - A partial title match

    Returns a song dict with keys: title, hymnal_number, hymnal_name,
    tune_name, sections. Returns None if not found.
    """
    songs = _get_songs(service)

    # Try to extract hymnal number from identifier
    # Format 1: "#93 Angels From the Realms" (11am sheet format)
    num_match = re.match(r'#(\d+)', identifier.strip())
    if num_match:
        hymnal_num = num_match.group(1)
        for song in songs:
            if song.get("hymnal_number") == hymnal_num:
                return song

    # Format 2: "Song Title H400" or "Song Title S129 (Powell)" (9am sheet format)
    hymn_ref = re.search(r'\b([HS])(\d+)\b', identifier)
    if hymn_ref:
        prefix = hymn_ref.group(1)
        number = hymn_ref.group(2)
        hymnal_num = f"S{number}" if prefix == "S" else number
        for song in songs:
            if song.get("hymnal_number") == hymnal_num:
                return song

    # Clean identifier for title matching: strip parenthetical notes
    # and hymnal references (H###, S###)
    clean = _clean_identifier(identifier)

    # Try exact title match (case-insensitive)
    id_lower = clean.lower()
    # Also strip leading #NNN from the cleaned identifier
    title_part = re.sub(r'^#\d+\s*', '', clean).strip()

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

    # Try punctuation-stripped match (handles "Bless the Lord my Soul"
    # matching "Bless the Lord, my soul")
    def _strip_punct(s):
        return re.sub(r'[,;:!?\'".\-]', '', s).lower()

    id_stripped = _strip_punct(clean)
    for song in songs:
        if _strip_punct(song["title"]) == id_stripped:
            return song
        if id_stripped in _strip_punct(song["title"]):
            return song

    # Cross-service fallback: if not found in the primary service,
    # try the other service's songs (lyrics are shared when needed)
    if not _in_fallback:
        svc = _normalize_service(service)
        fallback = "9am" if svc == "11am" else "11am"
        result = lookup_song(identifier, service=fallback, _in_fallback=True)
        if result:
            return result

    return None


def clear_cache():
    """Clear the song cache (useful for testing)."""
    global _all_songs, _hs_songs
    _all_songs = None
    _hs_songs = None


# ---------------------------------------------------------------------------
# Hidden Springs music library
# ---------------------------------------------------------------------------

def _load_hs_songs() -> list[dict]:
    """Load all songs from hidden_springs_songs.yaml (cached)."""
    global _hs_songs
    if _hs_songs is None:
        if not HS_SONGS_FILE.exists():
            _hs_songs = []
        else:
            with open(HS_SONGS_FILE, encoding="utf-8") as f:
                _hs_songs = yaml.safe_load(f) or []
    return _hs_songs


def parse_hs_music_field(raw: str) -> tuple[str, Optional[int], Optional[str]]:
    """Parse a Hidden Springs planner music field.

    Extracts optional verse count and instrument suffixes:
      "How Great Thou Art [4v]"  → ("How Great Thou Art", 4, None)
      "Song Title [PIANO]"      → ("Song Title", None, "piano")
      "Song Title"              → ("Song Title", None, None)

    Returns: (title, verse_count, instrument)
    """
    if not raw:
        return ("", None, None)

    title = raw.strip()
    verse_count = None
    instrument = None

    # Extract [Nv] suffix
    vm = re.search(r"\[(\d+)v\]", title, re.IGNORECASE)
    if vm:
        verse_count = int(vm.group(1))
        title = re.sub(r"\s*\[\d+v\]", "", title, flags=re.IGNORECASE).strip()

    # Extract [ORGAN] or [PIANO] suffix
    im = re.search(r"\[(ORGAN|PIANO)\]", title, re.IGNORECASE)
    if im:
        instrument = im.group(1).lower()
        title = re.sub(r"\s*-?\s*\[(ORGAN|PIANO)\]", "", title,
                        flags=re.IGNORECASE).strip()

    return (title, verse_count, instrument)


def _match_in_catalog(identifier: str, catalog: list[dict]) -> Optional[dict]:
    """Search a song catalog using the same matching strategies as lookup_song.

    Tries: hymnal number, exact title, identifiers/aliases, starts-with, substring.
    """
    clean = _clean_identifier(identifier)
    id_lower = clean.lower()
    title_part = re.sub(r'^#\d+\s*', '', clean).strip()

    # Try hymnal number from identifier
    num_match = re.match(r'#?(\d+)', identifier.strip())
    if num_match:
        hymnal_num = num_match.group(1)
        for song in catalog:
            if song.get("hymnal_number") == hymnal_num:
                return song

    # Try S-prefix: "S280", "S130"
    s_match = re.match(r'S(\d+)', identifier.strip(), re.IGNORECASE)
    if s_match:
        s_num = f"S{s_match.group(1)}"
        for song in catalog:
            if song.get("hymnal_number") == s_num:
                return song

    # Hymnal ref embedded in title: "Song Title H400" or "Song Title S129"
    hymn_ref = re.search(r'\b([HS])(\d+)\b', identifier)
    if hymn_ref:
        prefix = hymn_ref.group(1)
        number = hymn_ref.group(2)
        hymnal_num = f"S{number}" if prefix == "S" else number
        for song in catalog:
            if song.get("hymnal_number") == hymnal_num:
                return song

    # Exact title match (case-insensitive)
    for song in catalog:
        if song["title"].lower() == id_lower:
            return song
        if title_part and song["title"].lower() == title_part.lower():
            return song

    # Identifier/alias match
    for song in catalog:
        for alias in song.get("identifiers", []):
            if alias.lower() == id_lower:
                return song

    # Starts-with match
    for song in catalog:
        if song["title"].lower().startswith(id_lower):
            return song
        if id_lower.startswith(song["title"].lower()):
            return song

    # Substring match
    for song in catalog:
        if id_lower in song["title"].lower():
            return song

    return None


def hs_lookup_song(identifier: str) -> Optional[dict]:
    """Look up a song for Hidden Springs bulletins.

    Searches hidden_springs_songs.yaml first (for AAC file info),
    then falls back to songs.yaml for lyrics only.

    The identifier is parsed for optional [Nv] and [INSTRUMENT] suffixes.
    Returned dict includes _hs_verse_count and _hs_instrument metadata
    when specified.
    """
    title, verse_count, instrument = parse_hs_music_field(identifier)
    if not title:
        return None

    # Search HS catalog first
    hs_catalog = _load_hs_songs()
    song = _match_in_catalog(title, hs_catalog)

    if song:
        # Return a copy with HS metadata attached
        result = dict(song)
        if verse_count is not None:
            result["_hs_verse_count"] = verse_count
        if instrument is not None:
            result["_hs_instrument"] = instrument
        return result

    # Fallback: search songs.yaml for lyrics only (no AAC tracking)
    fallback = lookup_song(title, service="9am")
    if fallback:
        result = dict(fallback)
        result["_hs_fallback"] = True  # flag: no AAC data available
        if verse_count is not None:
            result["_hs_verse_count"] = verse_count
        return result

    return None


def resolve_aac_file(song_data: dict,
                     verse_count: Optional[int] = None,
                     instrument: Optional[str] = None) -> Optional[str]:
    """Resolve the AAC filename for a Hidden Springs song.

    Args:
        song_data: Song dict from hs_lookup_song() (must have 'aac_files' key)
        verse_count: Requested verse count (from [Nv] suffix)
        instrument: Requested instrument (from [PIANO]/[ORGAN] suffix)

    Returns:
        Filename relative to sorted_hidden_springs_aac_files/, or None.
    """
    aac_files = song_data.get("aac_files")
    if not aac_files:
        return None

    # If verse count requested, find matching variant
    if verse_count is not None:
        for af in aac_files:
            if af.get("verses") == verse_count:
                if instrument and af.get("instrument") != instrument:
                    continue
                return af["filename"]
        # No exact verse match — warn and fall through to default

    # If instrument requested, find matching variant
    if instrument is not None:
        for af in aac_files:
            if af.get("instrument") == instrument:
                return af["filename"]

    # Default: prefer the entry without verses/instrument tags
    for af in aac_files:
        if "verses" not in af and "instrument" not in af:
            return af["filename"]

    # If all entries have verse counts, pick the one with the most verses
    versioned = [af for af in aac_files if "verses" in af]
    if versioned:
        return max(versioned, key=lambda af: af["verses"])["filename"]

    # Last resort: first file
    return aac_files[0]["filename"]


def slice_song_verses(song_data: dict, max_verses: int) -> dict:
    """Return a copy of song_data with sections trimmed to max_verses.

    Keeps all chorus/refrain sections; only counts and limits verse sections.
    """
    sections = song_data.get("sections", [])
    if not sections:
        return song_data

    trimmed = []
    verse_count = 0
    for section in sections:
        if section.get("type") == "verse":
            verse_count += 1
            if verse_count > max_verses:
                continue
        trimmed.append(section)

    result = dict(song_data)
    result["sections"] = trimmed
    return result
