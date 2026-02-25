"""
Fetch liturgical schedule, clergy rota, and service music data from the
St. Andrew's Google Sheet.

The sheet is publicly viewable, so we use CSV export (no API key needed).
Each worksheet is exported by its GID and parsed into a list of dicts.
"""

import csv
import io
from datetime import date, datetime
from dataclasses import dataclass, field
from typing import Optional

import requests

from bulletin.config import SPREADSHEET_ID, SHEET_GIDS


def _fetch_sheet_csv(gid: int) -> list[dict]:
    """Download a worksheet as CSV and return a list of row dicts.

    The Google Sheets often have title/description rows above the real headers.
    We detect the header row by looking for a row containing 'Date' as a value.
    """
    url = (
        f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}"
        f"/export?format=csv&gid={gid}"
    )
    response = requests.get(url, timeout=30)
    response.raise_for_status()

    # Parse all rows as lists first, then find the header row
    all_rows = list(csv.reader(io.StringIO(response.text)))

    # Find the header row: the first row containing "Date" as a cell value
    header_idx = None
    for i, row in enumerate(all_rows):
        normalized = [c.strip().lower().replace("\n", " ") for c in row]
        if "date" in normalized:
            header_idx = i
            break

    if header_idx is None:
        raise ValueError(f"Could not find header row with 'Date' column in sheet gid={gid}")

    # Normalize header names: collapse newlines, strip whitespace
    headers = [c.strip().replace("\n", " ") for c in all_rows[header_idx]]

    # Build dicts for each data row after the header
    results = []
    for row in all_rows[header_idx + 1:]:
        if len(row) < len(headers):
            row.extend([""] * (len(headers) - len(row)))
        row_dict = {headers[j]: row[j] for j in range(len(headers))}
        results.append(row_dict)

    return results


def _parse_date(date_str: str) -> Optional[date]:
    """Parse a date string like '1/4/2026' or '2/18/2026' into a date object."""
    if not date_str or date_str.strip() == "":
        return None
    date_str = date_str.strip()
    for fmt in ("%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d"):
        try:
            return datetime.strptime(date_str, fmt).date()
        except ValueError:
            continue
    return None


# ---------------------------------------------------------------------------
# Data classes for the three key sheets
# ---------------------------------------------------------------------------

@dataclass
class LiturgicalScheduleRow:
    """One row from the Liturgical Schedule sheet."""
    service_type: str
    date: Optional[date]
    title: str                  # e.g., "Third Sunday in Lent"
    proper: str                 # e.g., "23" or "-"
    color: str                  # e.g., "Green", "Violet", "White", "Red"
    eucharistic_prayer: str     # "A", "B", or "C"
    preface: str                # e.g., "Incarnation", "Lent (1)", "Lord's Day"
    reading: str                # e.g., "1 Corinthians 10:1-13"
    psalm: str                  # e.g., "Psalm 63:1-8 responsively"
    gospel: str                 # e.g., "Luke 13:1-9"
    pop_form: str               # Prayers of the People form: "I", "II", etc.
    special_blessing: str       # e.g., "Solemn Prayer - Lent 1 (BOS)"
    closing_prayer: str         # "Almighty" or "Eternal God"
    dismissal: str              # "1", "2", "3", or "4"
    notes: str                  # Free-text notes

    # Hidden Springs alternate readings (may differ from main campus)
    hs_reading: str = ""
    hs_psalm: str = ""
    hs_gospel: str = ""


@dataclass
class ClergyRotaRow:
    """One row from the Clergy Rota sheet."""
    service_type: str
    date: Optional[date]
    title: str

    # 9 am service assignments (the primary bulletin target)
    celebrant_9am: str = ""
    deacon_word_9am: str = ""
    deacon_table_9am: str = ""
    preacher_9am: str = ""
    assisting_9am: str = ""
    subdeacon_9am: str = ""

    # 8 am service
    celebrant_8am: str = ""
    preacher_8am: str = ""
    deacon_8am: str = ""

    # 11 am service
    celebrant_11am: str = ""
    deacon_word_11am: str = ""
    deacon_table_11am: str = ""
    preacher_11am: str = ""
    assisting_11am: str = ""
    subdeacon_11am: str = ""


@dataclass
class ServiceMusicRow:
    """One row from the Service Music sheet."""
    service_type: str
    date: Optional[date]
    title: str
    reading: str = ""
    psalm: str = ""
    gospel: str = ""
    notes: str = ""

    prelude: str = ""
    processional: str = ""
    song_of_praise: str = ""
    sequence: str = ""
    anthem: str = ""        # offering/anthem
    sanctus: str = ""
    communion: str = ""     # may have multiple songs separated by delimiter
    recessional: str = ""   # closing hymn
    postlude: str = ""


# ---------------------------------------------------------------------------
# Column name mapping
# ---------------------------------------------------------------------------
# Google Sheets column headers may vary slightly. These mappings handle the
# known header names. We normalize by stripping whitespace and lowercasing.

def _normalize_key(key: str) -> str:
    return key.strip().lower().replace("\n", " ")


def _get(row: dict, *possible_keys, default: str = "") -> str:
    """Get a value from a row dict, trying multiple possible column names."""
    normalized = {_normalize_key(k): v for k, v in row.items()}
    for key in possible_keys:
        val = normalized.get(_normalize_key(key))
        if val is not None:
            return val.strip()
    return default


# ---------------------------------------------------------------------------
# Sheet fetching functions
# ---------------------------------------------------------------------------

def fetch_liturgical_schedule() -> list[LiturgicalScheduleRow]:
    """Fetch and parse the Liturgical Schedule sheet."""
    rows = _fetch_sheet_csv(SHEET_GIDS["liturgical_schedule"])
    results = []
    for row in rows:
        dt = _parse_date(_get(row, "date"))
        results.append(LiturgicalScheduleRow(
            service_type=_get(row, "service type"),
            date=dt,
            title=_get(row, "sunday/commemoration title", "title"),
            proper=_get(row, "proper"),
            color=_get(row, "color"),
            eucharistic_prayer=_get(row, "eucharistic prayer"),
            preface=_get(row, "preface"),
            reading=_get(row, "reading"),
            psalm=_get(row, "psalm"),
            gospel=_get(row, "gospel"),
            pop_form=_get(row, "pop"),
            special_blessing=_get(row, "special blessing"),
            closing_prayer=_get(row, "closing prayer"),
            dismissal=_get(row, "dismissal"),
            notes=_get(row, "notes"),
            hs_reading=_get(row, "hidden springs reading"),
            hs_psalm=_get(row, "hidden springs psalm"),
            hs_gospel=_get(row, "hidden springs gospel"),
        ))
    return results


def fetch_clergy_rota() -> list[ClergyRotaRow]:
    """Fetch and parse the Clergy Rota sheet."""
    rows = _fetch_sheet_csv(SHEET_GIDS["clergy_rota"])
    results = []
    for row in rows:
        dt = _parse_date(_get(row, "date"))
        results.append(ClergyRotaRow(
            service_type=_get(row, "service type"),
            date=dt,
            title=_get(row, "sunday/commemoration title", "title"),
            # 8 am
            celebrant_8am=_get(row, "8:00 am celebrant", "8 am celebrant",
                               "celebrant (8:00)"),
            preacher_8am=_get(row, "8:00 am preacher", "8 am preacher",
                              "preacher (8:00)"),
            deacon_8am=_get(row, "8:00 am deacon", "8 am deacon",
                            "deacon (8:00)"),
            # 9 am
            celebrant_9am=_get(row, "9:00 am celebrant", "9 am celebrant",
                               "celebrant (9:00)"),
            deacon_word_9am=_get(row, "9:00 am deacon of the word",
                                  "deacon of the word (9:00)",
                                  "9 am deacon of the word"),
            deacon_table_9am=_get(row, "9:00 am deacon of the table",
                                   "deacon of the table (9:00)",
                                   "9 am deacon of the table"),
            preacher_9am=_get(row, "9:00 am preacher", "9 am preacher",
                              "preacher (9:00)", "main campus preacher"),
            assisting_9am=_get(row, "9:00 am assisting priest",
                                "assisting priest (9:00)",
                                "9 am assisting priest"),
            subdeacon_9am=_get(row, "9:00 am subdeacon", "subdeacon (9:00)",
                               "9 am subdeacon"),
            # 11 am
            celebrant_11am=_get(row, "11:00 am celebrant", "11 am celebrant",
                                "celebrant (11:00)"),
            deacon_word_11am=_get(row, "11:00 am deacon of the word",
                                   "deacon of the word (11:00)",
                                   "11 am deacon of the word"),
            deacon_table_11am=_get(row, "11:00 am deacon of the table",
                                    "deacon of the table (11:00)",
                                    "11 am deacon of the table"),
            preacher_11am=_get(row, "11:00 am preacher", "11 am preacher",
                               "preacher (11:00)"),
            assisting_11am=_get(row, "11:00 am assisting priest",
                                 "assisting priest (11:00)",
                                 "11 am assisting priest"),
            subdeacon_11am=_get(row, "11:00 am subdeacon",
                                 "subdeacon (11:00)", "11 am subdeacon"),
        ))
    return results


def fetch_service_music() -> list[ServiceMusicRow]:
    """Fetch and parse the Service Music sheet."""
    rows = _fetch_sheet_csv(SHEET_GIDS["service_music"])
    results = []
    for row in rows:
        dt = _parse_date(_get(row, "date"))
        results.append(ServiceMusicRow(
            service_type=_get(row, "service type"),
            date=dt,
            title=_get(row, "sunday/commemoration title", "title"),
            reading=_get(row, "reading"),
            psalm=_get(row, "psalm"),
            gospel=_get(row, "gospel"),
            notes=_get(row, "notes"),
            prelude=_get(row, "prelude"),
            processional=_get(row, "processional"),
            song_of_praise=_get(row, "song of praise"),
            sequence=_get(row, "sequence"),
            anthem=_get(row, "anthem"),
            sanctus=_get(row, "sanctus"),
            communion=_get(row, "communion"),
            recessional=_get(row, "recessional"),
            postlude=_get(row, "postlude"),
        ))
    return results


# ---------------------------------------------------------------------------
# Lookup functions
# ---------------------------------------------------------------------------

@dataclass
class BulletinData:
    """All data from the Google Sheet needed for one bulletin."""
    schedule: LiturgicalScheduleRow
    clergy: Optional[ClergyRotaRow]
    music: Optional[ServiceMusicRow]


def get_bulletin_data(target_date: date) -> BulletinData:
    """Fetch all sheet data and look up the row for the target date.

    Raises ValueError if the target date is not found in the Liturgical Schedule.
    """
    schedule_rows = fetch_liturgical_schedule()
    clergy_rows = fetch_clergy_rota()
    music_rows = fetch_service_music()

    # Find the matching row in liturgical schedule
    schedule = None
    for row in schedule_rows:
        if row.date == target_date and row.service_type.lower().strip() == "sunday":
            schedule = row
            break

    if schedule is None:
        # Also try non-Sunday service types (feasts, etc.)
        for row in schedule_rows:
            if row.date == target_date:
                schedule = row
                break

    if schedule is None:
        available_dates = sorted(set(
            r.date.isoformat() for r in schedule_rows if r.date
        ))
        raise ValueError(
            f"Date {target_date.isoformat()} not found in the Liturgical Schedule. "
            f"Available dates range from {available_dates[0]} to {available_dates[-1]}."
        )

    # Find matching clergy rota (may not exist)
    clergy = None
    for row in clergy_rows:
        if row.date == target_date:
            clergy = row
            break

    # Find matching music
    music = None
    for row in music_rows:
        if row.date == target_date:
            music = row
            break

    return BulletinData(schedule=schedule, clergy=clergy, music=music)
