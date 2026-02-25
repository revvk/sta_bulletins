"""
Fetch 9am service music from the separate music planning spreadsheet.

This sheet uses a "table of tables" layout: a 3x3 grid of sub-tables,
each representing one week's service music. Each sub-table has 4 columns:
  - Service Part (e.g., "Processional:", "Song of Praise:")
  - Song title (with optional H### hymnal refs, S### service music refs)
  - Key (musical key)
  - Lead (musician name)

The 3 horizontal sub-tables are separated by an empty column (col E, J).
The 3 vertical rows of sub-tables are separated by blank rows.
"""

import csv
import io
import re
from datetime import date, datetime
from dataclasses import dataclass, field
from typing import Optional

import requests

# 9am music planning spreadsheet
MUSIC_9AM_SPREADSHEET_ID = "119FYdOXYhjDYS42tYlDm_OTEIegz-ZQXN-3PwoRYwd0"
MUSIC_9AM_GID = 1254884308


@dataclass
class MusicSlot:
    """One music slot in the service."""
    service_part: str       # e.g., "Processional", "Song of Praise", "Communion 1"
    song_title: str         # Raw title from sheet (may include H###, S###, verse info)
    key: str = ""           # Musical key (e.g., "F", "G", "Dm")
    lead: str = ""          # Lead musician


@dataclass
class ServiceMusic9am:
    """All music for one 9am service."""
    date: date
    liturgical_label: str   # e.g., "Lent 2A", "Easter Day A"
    slots: list[MusicSlot] = field(default_factory=list)

    def get_slot(self, part_name: str) -> Optional[MusicSlot]:
        """Look up a music slot by service part name (case-insensitive, partial match)."""
        part_lower = part_name.lower().strip()
        for slot in self.slots:
            if slot.service_part.lower().strip() == part_lower:
                return slot
            if part_lower in slot.service_part.lower():
                return slot
        return None

    def get_slots(self, part_prefix: str) -> list[MusicSlot]:
        """Get all slots matching a prefix (e.g., 'Communion' for Communion 1/2/3)."""
        prefix_lower = part_prefix.lower().strip()
        return [s for s in self.slots if s.service_part.lower().startswith(prefix_lower)]


def fetch_9am_music(target_date: date) -> Optional[ServiceMusic9am]:
    """Fetch 9am music for a specific date from the planning spreadsheet.

    Returns None if the date is not found in the current 9-week window.
    """
    url = (
        f"https://docs.google.com/spreadsheets/d/{MUSIC_9AM_SPREADSHEET_ID}"
        f"/export?format=csv&gid={MUSIC_9AM_GID}"
    )
    response = requests.get(url, timeout=30)
    response.raise_for_status()

    all_rows = list(csv.reader(io.StringIO(response.text)))
    sub_tables = _find_sub_tables(all_rows)

    for st in sub_tables:
        if st.date == target_date:
            return st

    return None


def _find_sub_tables(all_rows: list[list[str]]) -> list[ServiceMusic9am]:
    """Parse the 3x3 grid of sub-tables from the CSV data.

    We scan for rows containing "Service Planner:" to locate each sub-table's
    header, then extract the date and music data.
    """
    results = []

    # The 3 horizontal column offsets for sub-tables
    col_offsets = [0, 5, 10]

    for row_idx, row in enumerate(all_rows):
        for col_offset in col_offsets:
            if col_offset >= len(row):
                continue
            cell = row[col_offset].strip()
            if cell.startswith("Service Planner:"):
                st = _parse_sub_table(all_rows, row_idx, col_offset)
                if st:
                    results.append(st)

    return results


def _parse_sub_table(all_rows: list[list[str]], header_row: int,
                     col_offset: int) -> Optional[ServiceMusic9am]:
    """Parse one sub-table starting at the given position.

    Layout:
      Row 0 (header_row):   "Service Planner: This Week" | "" | "Date:" | "2026-02-15"
      Row 1 (header_row+1): "Service Part" | "Song (9 am) - Lent 1A" | "Key" | "Lead"
      Row 2+ (data):        "Processional:" | "Build My Life" | "G" | "Steph"
    """
    # Extract date from header row
    date_col = col_offset + 3
    if date_col >= len(all_rows[header_row]):
        return None

    date_str = all_rows[header_row][date_col].strip()
    parsed_date = _parse_date(date_str)
    if not parsed_date:
        return None

    # Extract liturgical label from the second header row
    liturgical_label = ""
    if header_row + 1 < len(all_rows):
        song_col = col_offset + 1
        if song_col < len(all_rows[header_row + 1]):
            header_text = all_rows[header_row + 1][song_col].strip()
            # Format: "Song (9 am) - Lent 1A" or "Song (9 am) - Easter Day A (CC Celebration Sunday)"
            match = re.search(r'-\s*(.+)', header_text)
            if match:
                liturgical_label = match.group(1).strip()

    # Extract music data rows
    slots = []
    for i in range(header_row + 2, min(header_row + 20, len(all_rows))):
        row = all_rows[i]
        if col_offset >= len(row):
            break

        part = row[col_offset].strip() if col_offset < len(row) else ""
        song = row[col_offset + 1].strip() if col_offset + 1 < len(row) else ""
        key = row[col_offset + 2].strip() if col_offset + 2 < len(row) else ""
        lead = row[col_offset + 3].strip() if col_offset + 3 < len(row) else ""

        # Stop at empty part column (end of sub-table data)
        if not part:
            break

        # Skip header-like rows
        if part.lower().startswith("service part"):
            continue

        # Clean up the service part name (remove trailing colon)
        part_clean = part.rstrip(":").strip()

        if song:  # Only include slots that have a song assigned
            slots.append(MusicSlot(
                service_part=part_clean,
                song_title=song,
                key=key,
                lead=lead,
            ))

    return ServiceMusic9am(
        date=parsed_date,
        liturgical_label=liturgical_label,
        slots=slots,
    )


def _parse_date(date_str: str) -> Optional[date]:
    """Parse date strings in various formats."""
    if not date_str:
        return None
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y"):
        try:
            return datetime.strptime(date_str.strip(), fmt).date()
        except ValueError:
            continue
    return None


# ---------------------------------------------------------------------------
# Song identifier parsing
# ---------------------------------------------------------------------------

def parse_song_identifier(raw_title: str) -> dict:
    """Parse a raw song title from the 9am sheet into structured data.

    Examples:
        "Build My Life" -> {"title": "Build My Life", "hymnal_number": None}
        "All Creatures of Our God and King H400 (V1,3-4)" ->
            {"title": "All Creatures of Our God and King", "hymnal_number": "400",
             "hymnal_name": "Hymnal 1982", "verses": "V1,3-4"}
        "S129 (Powell)" ->
            {"title": "Holy, holy, holy Lord", "hymnal_number": "S129",
             "hymnal_name": "Hymnal 1982", "setting": "Powell"}
    """
    result = {
        "raw": raw_title,
        "title": raw_title,
        "hymnal_number": None,
        "hymnal_name": None,
        "verses": None,
        "setting": None,
    }

    # Extract verse indications like (V1,3-4) or (v1,3)
    verse_match = re.search(r'\(V[\d,\-]+\)', raw_title, re.IGNORECASE)
    if verse_match:
        result["verses"] = verse_match.group(0).strip("()")
        raw_title = raw_title[:verse_match.start()] + raw_title[verse_match.end():]

    # Extract parenthetical info like (Powell), (Schubert), (Hallé)
    paren_match = re.search(r'\(([^)]+)\)', raw_title)
    if paren_match:
        result["setting"] = paren_match.group(1).strip()
        raw_title = raw_title[:paren_match.start()] + raw_title[paren_match.end():]

    # Extract hymnal number H### or S###
    hymn_match = re.search(r'\b([HS])(\d+)\b', raw_title)
    if hymn_match:
        prefix = hymn_match.group(1)
        number = hymn_match.group(2)
        if prefix == "H":
            result["hymnal_number"] = number
        else:
            result["hymnal_number"] = f"S{number}"
        result["hymnal_name"] = "Hymnal 1982"
        # Remove the hymnal ref from the title
        raw_title = raw_title[:hymn_match.start()] + raw_title[hymn_match.end():]

    result["title"] = raw_title.strip().rstrip("-").strip()

    return result
