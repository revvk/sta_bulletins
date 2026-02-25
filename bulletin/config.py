"""
Church-specific constants and configuration for St. Andrew's Episcopal Church.
"""

# Google Sheet
SPREADSHEET_ID = "1419MFr7ctOypj0fJsXLNj-I2LoIJptvNGmO4sy2HCB8"

# Sheet GIDs (for CSV export)
SHEET_GIDS = {
    "liturgical_schedule": 1113515164,
    "clergy_rota": 1436942576,
    "service_music": 1988853205,
    "hidden_springs": 2077691139,
}

# 9am Music Planning Sheet (separate spreadsheet, "table of tables" layout)
MUSIC_9AM_SPREADSHEET_ID = "119FYdOXYhjDYS42tYlDm_OTEIegz-ZQXN-3PwoRYwd0"
MUSIC_9AM_GID = 1254884308

# Parish Cycle of Prayers (ministry rotation)
PARISH_PRAYERS_SPREADSHEET_ID = "1GzhkbQIKxmrOpnmp4w3QWHZX_DIu-IlTj5WuP6eVJYE"
PARISH_PRAYERS_GID = 0

# Lectionary URLs
LECTIONARY_BASE_URL = "https://lectionarypage.net"
OREMUS_BASE_URL = "https://bible.oremus.org"
OREMUS_PARAMS = {
    "vnum": "yes",
    "version": "NRSVAE",
    "fnote": "no",
    "heading": "no",
}

# Service times that generate separate bulletins
SERVICE_TIMES = ["8 am", "9 am", "11 am"]

# Church info
CHURCH_NAME = "St. Andrew's Episcopal Church"
CHURCH_CITY = "McKinney, Texas"
GIVING_URL = "standrewsmckinney.org/give"
CONNECT_URL = "standrewsmckinney.org/connect"

# Page dimensions (half-letter / statement size)
PAGE_WIDTH_INCHES = 7.0
PAGE_HEIGHT_INCHES = 8.5
MARGIN_INCHES = 0.5

# Font names used in the bulletin
FONT_BODY = "Adobe Garamond Pro"
FONT_HEADING = "Gill Sans Nova SemiBold"
FONT_HEADING2 = "Gill Sans Nova Medium"
FONT_LYRICS = "Gill Sans Light"
FONT_HEADER_FOOTER = "Gill Sans Nova Light"

# Cross symbol used throughout the liturgy
CROSS_SYMBOL = "\u2720"  # Maltese cross ✠

# Lectionary years cycle: year % 3 -> A=1, B=2, C=0
# The church year begins on Advent 1, but for calendar mapping
# the lectionary year letter is based on the calendar year of
# the Sunday in question (with Advent starting the NEXT year's cycle).
LECTIONARY_YEAR_MAP = {
    1: "A",
    2: "B",
    0: "C",
}


def get_lectionary_year(calendar_year: int) -> str:
    """Return the lectionary year letter (A, B, or C) for a given calendar year.

    Note: Advent Sundays in late November/December actually begin the NEXT
    year's lectionary cycle. That adjustment is handled in the calendar module,
    not here. This function gives the base year letter.
    """
    return LECTIONARY_YEAR_MAP[calendar_year % 3]
