"""
Parse 11am service music from the shared Service Music Google Sheet row.

The 11am service uses hymnal references in the format '#NNN Title' (e.g.,
'#473 Lift High the Cross'). This module converts a ServiceMusicRow into
a list of MusicSlot objects compatible with the builder's _lookup_slot()
interface.
"""

import re
from typing import Optional

from bulletin.sources.google_sheet import ServiceMusicRow
from bulletin.sources.music_9am import MusicSlot


# Mapping from ServiceMusicRow field names to service part labels used
# by the builder's _lookup_slot().
_FIELD_TO_SLOT = {
    "processional": "Processional",
    "song_of_praise": "Song of Praise",
    "sequence": "Sequence",
    "anthem": "Anthem",
    "sanctus": "Sanctus",
    "recessional": "Recessional",
}


def get_11am_music_slots(music_row: ServiceMusicRow) -> list[MusicSlot]:
    """Convert a ServiceMusicRow into a list of MusicSlot objects.

    Handles:
      - Simple fields (processional, song_of_praise, etc.) → one slot each
      - Communion field → split on ';' into Communion 1, Communion 2, etc.
    """
    slots: list[MusicSlot] = []

    # Simple single-song fields
    for field_name, slot_label in _FIELD_TO_SLOT.items():
        raw = getattr(music_row, field_name, "").strip()
        if raw:
            slots.append(MusicSlot(
                service_part=slot_label,
                song_title=raw,
            ))

    # Communion songs — may have multiple separated by ';'
    communion_raw = (music_row.communion or "").strip()
    if communion_raw:
        parts = [p.strip() for p in communion_raw.split(";") if p.strip()]
        for i, part in enumerate(parts, 1):
            slots.append(MusicSlot(
                service_part=f"Communion {i}",
                song_title=part,
            ))

    return slots


def parse_11am_identifier(raw_title: str) -> dict:
    """Parse a raw song identifier from the service music sheet.

    Handles formats like:
      '#473 Lift High the Cross'  → hymnal_number='473', title='Lift High the Cross'
      '#S129 Powell'              → hymnal_number='S129', title='', setting='Powell'
      'Bless the Lord, my soul'   → no hymnal_number, title='Bless the Lord, my soul'

    Returns a dict with:
      title, hymnal_number, hymnal_name, setting
    """
    result = {
        "raw": raw_title,
        "title": raw_title.strip(),
        "hymnal_number": None,
        "hymnal_name": None,
        "setting": None,
    }

    if not raw_title or not raw_title.strip():
        return result

    raw = raw_title.strip()

    # Extract parenthetical info like (Powell), (Schubert)
    paren_match = re.search(r'\(([^)]+)\)', raw)
    if paren_match:
        result["setting"] = paren_match.group(1).strip()
        raw = raw[:paren_match.start()] + raw[paren_match.end():]
        raw = raw.strip()

    # Try #S### or #H### format (explicit prefix)
    sh_match = re.match(r'^#([SH])(\d+)\s*(.*)', raw, re.IGNORECASE)
    if sh_match:
        prefix = sh_match.group(1).upper()
        number = sh_match.group(2)
        rest = sh_match.group(3).strip()
        if prefix == "S":
            result["hymnal_number"] = f"S{number}"
        else:
            result["hymnal_number"] = number
        result["hymnal_name"] = "Hymnal 1982"
        result["title"] = rest if rest else ""
        return result

    # Try #NNN format (no prefix)
    num_match = re.match(r'^#(\d+)\s*(.*)', raw)
    if num_match:
        number = num_match.group(1)
        rest = num_match.group(2).strip()
        result["hymnal_number"] = number
        result["hymnal_name"] = "Hymnal 1982"
        result["title"] = rest if rest else ""
        return result

    # No hymnal number — it's just a title
    result["title"] = raw.strip()
    return result
