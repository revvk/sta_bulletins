"""
Parish Cycle of Prayers lookup.

The cycle assigns 3 ministries per Sunday in alphabetical order,
rotating through 48 ministries over ~18 weeks (16 regular + 2 special).
The cycle repeats year after year.

Source: Google Sheet with Date/Ministry columns from a reference year (2022).
"""

import csv
import io
from datetime import date, datetime, timedelta
from typing import Optional

import requests

PARISH_PRAYERS_SPREADSHEET_ID = "1GzhkbQIKxmrOpnmp4w3QWHZX_DIu-IlTj5WuP6eVJYE"
PARISH_PRAYERS_GID = 0


def fetch_parish_cycle() -> list[tuple[str, list[str]]]:
    """Fetch the parish cycle of prayers from the Google Sheet.

    Returns a list of (date_label, [ministry1, ministry2, ministry3]) tuples,
    representing each week in order.
    """
    url = (
        f"https://docs.google.com/spreadsheets/d/{PARISH_PRAYERS_SPREADSHEET_ID}"
        f"/export?format=csv&gid={PARISH_PRAYERS_GID}"
    )
    response = requests.get(url, timeout=30)
    response.raise_for_status()

    rows = list(csv.reader(io.StringIO(response.text)))

    # Skip header row
    header_idx = None
    for i, row in enumerate(rows):
        if any("date" in c.lower() for c in row):
            header_idx = i
            break

    if header_idx is None:
        header_idx = 0

    weeks = []
    current_date_label = ""
    current_ministries = []

    for row in rows[header_idx + 1:]:
        if len(row) < 2:
            continue

        date_cell = row[0].strip()
        ministry_cell = row[1].strip()

        if date_cell:
            # New week -- flush previous
            if current_ministries:
                weeks.append((current_date_label, current_ministries))
            current_date_label = date_cell
            current_ministries = []

        if ministry_cell:
            current_ministries.append(ministry_cell)

    # Flush last week
    if current_ministries:
        weeks.append((current_date_label, current_ministries))

    return weeks


def get_ministries_for_date(target_date: date) -> list[str]:
    """Get the ministries for the Parish Cycle of Prayers for a given Sunday.

    The cycle is 18 weeks long (16 regular + 2 special). We calculate
    which week of the cycle a given Sunday falls in by counting weeks
    from a known anchor date, modulo the cycle length.
    """
    cycle = fetch_parish_cycle()

    if not cycle:
        return ["[Parish cycle data not available]"]

    # Filter to regular weeks (exclude special prayers weeks)
    regular_weeks = []
    special_weeks = {}

    for date_label, ministries in cycle:
        # Check if this is a special week
        if len(ministries) == 1 and "special" in ministries[0].lower():
            special_weeks[ministries[0]] = ministries
        else:
            regular_weeks.append(ministries)

    # Total regular cycle length
    cycle_length = len(regular_weeks)

    if cycle_length == 0:
        return ["[No regular ministry weeks found]"]

    # Parse the first date from the cycle to establish an anchor
    anchor_date = _parse_cycle_date(cycle[0][0])
    if not anchor_date:
        # Fallback: use a known start (August 28, 2022 from the data)
        anchor_date = date(2022, 8, 28)

    # Calculate week offset
    days_diff = (target_date - anchor_date).days
    week_offset = days_diff // 7

    # Map to cycle position (modulo regular cycle length)
    cycle_pos = week_offset % cycle_length

    return regular_weeks[cycle_pos]


def _parse_cycle_date(date_str: str) -> Optional[date]:
    """Parse dates like 'August 28, 2022' from the parish cycle sheet."""
    if not date_str:
        return None
    for fmt in ("%B %d, %Y", "%m/%d/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(date_str.strip(), fmt).date()
        except ValueError:
            continue
    return None


def format_ministries(ministries: list[str]) -> str:
    """Format a list of ministries for insertion into the POP text.

    e.g., ["Altar Guild", "Belize Mission Team", "Bible Builders"]
    -> "Altar Guild, Belize Mission Team, and Bible Builders"
    """
    if not ministries:
        return ""
    if len(ministries) == 1:
        return ministries[0]
    if len(ministries) == 2:
        return f"{ministries[0]} and {ministries[1]}"
    return ", ".join(ministries[:-1]) + ", and " + ministries[-1]
