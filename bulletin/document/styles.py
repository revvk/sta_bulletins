"""
Defines all paragraph and character styles used in the bulletin.

These styles replicate the formatting found in the existing St. Andrew's
bulletins. The document uses half-letter (7" x 8.5") pages with 0.5" margins.

Two font families are used:
  - Adobe Garamond Pro: body text, prayers, scripture, rubrics
  - Gill Sans Nova:     headings (Bold for H1, Medium for H2)
  - Gill Sans Light:    lyrics
  - Gill Sans Nova Light: headers/footers

Character styles:
  - "People": bold — applied to runs where the people respond.
    This replaces the old "Body - People" paragraph style.
"""

from docx import Document
from docx.shared import Pt, Inches, Emu, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

from bulletin.config import (
    FONT_BODY,
    FONT_BODY_BOLD,
    FONT_HEADING,
    FONT_HEADING2,
    FONT_LYRICS,
    FONT_HEADER_FOOTER,
    PAGE_WIDTH_INCHES,
    PAGE_HEIGHT_INCHES,
    MARGIN_INCHES,
)


# Tab stop positions used in Body and Body - Dialogue styles.
# Left tab at 1.5" aligns Celebrant/People label text.
# Right tab at 6" right-aligns hymnal references on song header lines.
_TAB_LEFT_INCHES = 1.5
_TAB_RIGHT_INCHES = PAGE_WIDTH_INCHES - 2 * MARGIN_INCHES  # 6.0"


def create_document() -> Document:
    """Create a new Document with all bulletin styles and page setup."""
    doc = Document()
    configure_document(doc)
    return doc


def configure_document(doc: Document):
    """Set up page layout and register all bulletin styles on *doc*.

    Use this when the Document already exists (e.g. opened from a
    template) rather than being freshly created.
    """
    _setup_page(doc)
    _create_styles(doc)


def _setup_page(doc: Document):
    """Configure the first section's page size and margins."""
    section = doc.sections[0]
    section.page_width = Inches(PAGE_WIDTH_INCHES)
    section.page_height = Inches(PAGE_HEIGHT_INCHES)
    section.left_margin = Inches(MARGIN_INCHES)
    section.right_margin = Inches(MARGIN_INCHES)
    section.top_margin = Inches(MARGIN_INCHES)
    section.bottom_margin = Inches(MARGIN_INCHES)
    section.header_distance = Inches(MARGIN_INCHES)
    section.footer_distance = Inches(0.25)
    section.different_first_page_header_footer = True


# ---------------------------------------------------------------------------
# Style definitions
# ---------------------------------------------------------------------------
# Each entry: (style_name, font_name, size_pt, bold, italic, alignment,
#               left_indent_in, first_indent_in, space_before_pt,
#               space_after_pt, line_spacing)

_STYLE_DEFS = [
    # Major section dividers (centered): "The Word of God", "The Holy Communion"
    # Gill Sans Nova Bold, 14 pt, centered
    ("Heading", FONT_HEADING, 14, True, False,
     WD_ALIGN_PARAGRAPH.CENTER, 0, 0, 0, 0, 0.9),

    # Section headers: "Processional", "Sermon", "The Peace", etc.
    # Gill Sans Nova Medium, 13 pt, left-aligned
    ("Heading 2", FONT_HEADING2, 13, False, False,
     WD_ALIGN_PARAGRAPH.LEFT, 0, 0, 0, 0, 1.0),

    # Main body text (prayers, collect text, general content)
    # Tab stops: left at 1.5", right at 6"
    ("Body", FONT_BODY, 12, False, False,
     WD_ALIGN_PARAGRAPH.LEFT, 0.32, 0, 0, 0, 1.0),

    # Celebrant/People dialogue lines with hanging indent:
    #   "Celebrant  [tab]  The Lord be with you."
    #   "People     [tab]  And also with you."   (with People char style)
    # Hanging indent ensures wrapped text aligns at the 1.5" tab stop.
    ("Body - Dialogue", FONT_BODY, 12, False, False,
     WD_ALIGN_PARAGRAPH.LEFT, 1.5, -1.18, 0, 0, 1.0),

    # Rubrics within the service: "After the reading, the Reader will say"
    ("Body - Rubric", FONT_BODY, 10, False, True,
     WD_ALIGN_PARAGRAPH.LEFT, 0.32, 0, 2, 0, 1.0),

    # Introductory rubrics: "Please stand.", "Be seated.", "Remain standing."
    ("Body - Introductory Rubric", FONT_BODY, 10, False, True,
     WD_ALIGN_PARAGRAPH.LEFT, 0.004, 0, 0, 0, 1.0),

    # Song lyrics (verse text; choruses use italic runs)
    ("Body - Lyrics", FONT_LYRICS, 12, False, False,
     WD_ALIGN_PARAGRAPH.LEFT, 0.5, -0.18, 0, 0, 1.0),

    # Right-column lyrics in two-column layouts — larger hanging indent
    # because the narrower column needs more wrap room
    ("Body - Lyrics Right", FONT_LYRICS, 12, False, False,
     WD_ALIGN_PARAGRAPH.LEFT, 0.5, -0.38, 0, 0, 1.0),

    # Scripture reading prose text
    ("Reading/Gospel Text", FONT_BODY, 11, False, False,
     WD_ALIGN_PARAGRAPH.LEFT, 0.32, 0, 0, 0, 1.0),

    # Scripture reading poetry (prophetic/poetic quotations)
    # Indent 0 (base poetry): left 0.705" to visually set off from prose (0.32")
    ("Reading (Poetry)", FONT_BODY, 11, False, False,
     WD_ALIGN_PARAGRAPH.LEFT, 0.705, -0.135, 0, 0, 1.0),

    # Poetry indent level 1: additional 0.25" indent
    ("Reading (Poetry Indent 1)", FONT_BODY, 11, False, False,
     WD_ALIGN_PARAGRAPH.LEFT, 0.955, -0.135, 0, 0, 1.0),

    # Poetry indent level 2: additional 0.50" indent
    ("Reading (Poetry Indent 2)", FONT_BODY, 11, False, False,
     WD_ALIGN_PARAGRAPH.LEFT, 1.205, -0.135, 0, 0, 1.0),

    # Psalm text (non-bold base; bold applied per-verse for responsive readings)
    ("Psalm", FONT_BODY, 12, False, False,
     WD_ALIGN_PARAGRAPH.LEFT, 0.455, -0.13, 3, 0, 1.0),

    # People's recitation of long texts: Confession, Lord's Prayer
    ("Body - People Recitation", FONT_BODY, 12, False, False,
     WD_ALIGN_PARAGRAPH.LEFT, 0.32, 0, 6, 0, 1.0),

    # Nicene Creed (slightly different indent for the three articles)
    ("Body - People Recitation (Creed)", FONT_BODY, 12, False, False,
     WD_ALIGN_PARAGRAPH.LEFT, 0.455, -0.18, 6, 0, 1.0),

    # Cover page welcome text
    ("Cover Note", FONT_BODY, 11, False, False,
     WD_ALIGN_PARAGRAPH.JUSTIFY, 0, 0, 0, 0, 1.0),

    # Small vertical spacer between elements
    ("Spacer - Small", FONT_BODY, 10, False, False,
     WD_ALIGN_PARAGRAPH.LEFT, 0, 0, 0, 0, 1.0),

    # Staff listing - name (bold, centered)
    ("Staff - Name", FONT_BODY, 9, True, False,
     WD_ALIGN_PARAGRAPH.CENTER, 0, 0, 5, 0, 1.0),

    # Staff listing - title (italic, centered)
    ("Staff - Title", FONT_BODY, 9, False, True,
     WD_ALIGN_PARAGRAPH.CENTER, 0, 0, 0, 0, 1.0),

    # Staff listing - email (centered)
    ("Staff - Email", FONT_BODY, 9, False, False,
     WD_ALIGN_PARAGRAPH.CENTER, 0, 0, 0, 0, 1.0),

    # Header & Footer text
    ("Header & Footer", FONT_HEADER_FOOTER, 9, False, False,
     WD_ALIGN_PARAGRAPH.LEFT, 0, 0, 0, 0, 1.0),
]

# Passion Gospel part styles — created dynamically in _create_passion_gospel_styles.
# Each part gets its own style (e.g., "Passion Gospel - Narrator") with a hanging
# indent so the speaker label sits on the left and spoken text wraps cleanly.
# All style definitions are identical; separate styles allow generating
# individual reading sheets for each part later.
_PASSION_GOSPEL_PARTS = [
    # Matthew (Year A)
    "Narrator", "Jesus", "Pilate", "Wife", "Soldier",
    "Servant 1", "Servant 2", "Passerby 1", "Passerby 2",
    "Chief Priest", "Scribe", "Elder", "Bystander 1", "Bystander 2",
    "Centurion",
    # John (Good Friday) — additional parts
    "Peter", "Guard", "Police", "Priest", "Slave",
    # Luke (Year C) — additional parts
    "Disciples", "Maid", "Officer",
    # John (Good Friday) — additional parts
    "Elder 1", "Elder 2", "Slave",
]

# Styles that get the left + right tab stops.
_STYLES_WITH_TABS = {"Body", "Body - Dialogue"}

# ---------------------------------------------------------------------------
# Reading sheet style definitions — same style names, larger sizes, red rubrics
# ---------------------------------------------------------------------------
# Extends the 11-field format with an optional 12th field: color (RGBColor).
# When omitted or None, defaults to black.

_RS_RED = RGBColor(0xED, 0x22, 0x0B)

_RS_STYLE_DEFS = [
    # Section marker headings (red, centered)
    ("Heading", FONT_HEADING, 16, True, False,
     WD_ALIGN_PARAGRAPH.CENTER, 0, 0, 0, 0, 1.0, _RS_RED),

    # Sub-headings (red)
    ("Heading 2", FONT_HEADING2, 16, False, False,
     WD_ALIGN_PARAGRAPH.LEFT, 0, 0, 0, 0, 1.0, _RS_RED),

    # Body text — readings, prayers
    ("Body", FONT_BODY, 16, False, False,
     WD_ALIGN_PARAGRAPH.LEFT, 0.32, 0, 0, 0, 1.0),

    # Celebrant/People dialogue (larger for readability)
    ("Body - Dialogue", FONT_BODY, 16, False, False,
     WD_ALIGN_PARAGRAPH.LEFT, 1.5, -1.18, 0, 0, 1.0),

    # Rubrics — red italic instructions for the reader
    ("Body - Rubric", FONT_BODY, 14, False, True,
     WD_ALIGN_PARAGRAPH.LEFT, 0.32, 0, 2, 0, 1.0, _RS_RED),

    # Introductory rubrics (red italic)
    ("Body - Introductory Rubric", FONT_BODY, 14, False, True,
     WD_ALIGN_PARAGRAPH.LEFT, 0.004, 0, 0, 0, 1.0, _RS_RED),

    # Scripture reading prose text (large for reading aloud)
    ("Reading/Gospel Text", FONT_BODY, 16, False, False,
     WD_ALIGN_PARAGRAPH.LEFT, 0.32, 0, 0, 0, 1.0),

    # Scripture reading poetry
    ("Reading (Poetry)", FONT_BODY, 16, False, False,
     WD_ALIGN_PARAGRAPH.LEFT, 0.705, -0.135, 0, 0, 1.0),

    # Poetry indent level 1
    ("Reading (Poetry Indent 1)", FONT_BODY, 16, False, False,
     WD_ALIGN_PARAGRAPH.LEFT, 0.955, -0.135, 0, 0, 1.0),

    # Poetry indent level 2
    ("Reading (Poetry Indent 2)", FONT_BODY, 16, False, False,
     WD_ALIGN_PARAGRAPH.LEFT, 1.205, -0.135, 0, 0, 1.0),

    # Psalm text
    ("Psalm", FONT_BODY, 16, False, False,
     WD_ALIGN_PARAGRAPH.LEFT, 0.455, -0.13, 3, 0, 1.0),

    # People's recitation
    ("Body - People Recitation", FONT_BODY, 16, False, False,
     WD_ALIGN_PARAGRAPH.LEFT, 0.32, 0, 6, 0, 1.0),

    # Spacer
    ("Spacer - Small", FONT_BODY, 12, False, False,
     WD_ALIGN_PARAGRAPH.LEFT, 0, 0, 0, 0, 1.0),
]


def _create_styles(doc: Document, style_defs=None):
    """Register paragraph and character styles in the document.

    Args:
        doc: The Document to configure.
        style_defs: List of style definition tuples. Each tuple has 11 or 12
            fields — the 12th (optional) is an RGBColor for font color.
            Defaults to _STYLE_DEFS (bulletin styles).
    """
    if style_defs is None:
        style_defs = _STYLE_DEFS

    # --- Paragraph styles ---
    for entry in style_defs:
        name, font_name, size_pt, bold, italic, alignment, \
            left_in, first_in, sp_before_pt, sp_after_pt, \
            line_spacing = entry[:11]
        color = entry[11] if len(entry) > 11 else None

        # If the style already exists (e.g., "Heading"), modify it;
        # otherwise create a new one.
        try:
            style = doc.styles[name]
        except KeyError:
            style = doc.styles.add_style(name, 1)  # 1 = WD_STYLE_TYPE.PARAGRAPH

        # Font — use the Bold weight font name for bold styles so Word
        # picks the real bold weight instead of applying synthetic faux-bold.
        if bold and font_name == FONT_BODY:
            style.font.name = FONT_BODY_BOLD
        else:
            style.font.name = font_name
        style.font.size = Pt(size_pt)
        style.font.bold = bold
        style.font.italic = italic
        style.font.color.rgb = color if color else RGBColor(0, 0, 0)

        # Built-in heading styles carry theme-font attributes (asciiTheme,
        # hAnsiTheme, …) that override the explicit font name.  Strip them
        # so our Gill Sans Nova settings actually take effect.
        _strip_theme_fonts(style)

        # Built-in heading styles are "linked" (paragraph + character).
        # Remove the link so they behave as pure paragraph styles.
        link_elem = style.element.find(qn("w:link"))
        if link_elem is not None:
            style.element.remove(link_elem)

        # Paragraph format
        pf = style.paragraph_format
        pf.alignment = alignment
        pf.left_indent = Inches(left_in) if left_in else Inches(0)
        pf.first_line_indent = Inches(first_in) if first_in else Inches(0)
        pf.space_before = Pt(sp_before_pt) if sp_before_pt else Pt(0)
        pf.space_after = Pt(sp_after_pt) if sp_after_pt else Pt(0)

        if line_spacing is not None:
            if isinstance(line_spacing, float) and line_spacing < 3:
                # Proportional (e.g., 1.0, 1.1, 0.9)
                pf.line_spacing = line_spacing
            else:
                pf.line_spacing = Pt(line_spacing)

        # Keep with next for headings (prevents orphaned section headers)
        if name.startswith("Heading"):
            pf.keep_with_next = True

        # Tab stops for body-text styles
        if name in _STYLES_WITH_TABS:
            _add_tab_stops(pf)

    # --- Passion Gospel part styles ---
    if style_defs is _STYLE_DEFS:
        _create_passion_gospel_styles(doc)

    # --- Character styles ---
    _create_character_styles(doc)


def _strip_theme_fonts(style):
    """Remove asciiTheme / hAnsiTheme / etc. from a style's <w:rFonts>.

    Word's built-in heading styles (Heading 1, Heading 2, …) carry
    theme-font attributes that take priority over the explicit ascii/hAnsi
    font names set by python-docx.  Stripping the theme attributes lets
    the explicit font name win.
    """
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    rPr = style.element.find(f"{{{W}}}rPr")
    if rPr is None:
        return
    rFonts = rPr.find(f"{{{W}}}rFonts")
    if rFonts is None:
        return
    for attr in list(rFonts.attrib):
        if "Theme" in attr:
            del rFonts.attrib[attr]


def _add_tab_stops(pf):
    """Add left (1.5") and right (6") tab stops to a paragraph format."""
    tab_stops = pf.tab_stops
    tab_stops.add_tab_stop(Inches(_TAB_LEFT_INCHES), WD_TAB_ALIGNMENT.LEFT)
    tab_stops.add_tab_stop(Inches(_TAB_RIGHT_INCHES), WD_TAB_ALIGNMENT.RIGHT)


def _create_passion_gospel_styles(doc: Document):
    """Create a paragraph style for each Passion Gospel speaking part.

    Each style has a hanging indent: the speaker name sits at the left margin
    and spoken text wraps at a tab stop (1.18" indent with -0.86" first-line).
    The styles are identical in formatting but named separately
    (e.g., "Passion Gospel - Narrator") so individual reading sheets can
    be generated later by filtering on style name.
    """
    for part_name in _PASSION_GOSPEL_PARTS:
        style_name = f"Passion Gospel - {part_name}"
        try:
            style = doc.styles[style_name]
        except KeyError:
            style = doc.styles.add_style(style_name, 1)  # PARAGRAPH

        style.font.name = FONT_BODY
        style.font.size = Pt(11)
        style.font.bold = False
        style.font.italic = False
        style.font.color.rgb = RGBColor(0, 0, 0)

        pf = style.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        # Hanging indent: left at 1.18", first line pulls back to 0.32"
        pf.left_indent = Inches(1.18)
        pf.first_line_indent = Inches(-0.86)
        pf.space_before = Pt(0)
        pf.space_after = Pt(7)
        pf.line_spacing = 1.0

        # Tab stop at the hanging indent position for clean text alignment
        tab_stops = pf.tab_stops
        tab_stops.add_tab_stop(Inches(1.18), WD_TAB_ALIGNMENT.LEFT)


def _create_character_styles(doc: Document):
    """Create character styles (applied to runs, not paragraphs)."""
    # "People" character style — bold text for congregational responses.
    # We explicitly set the font name so that Word uses the proper bold
    # weight ("Adobe Garamond Pro Bold") rather than applying a synthetic
    # faux-bold to the Regular weight.
    try:
        people = doc.styles["People"]
    except KeyError:
        people = doc.styles.add_style("People", 2)  # 2 = WD_STYLE_TYPE.CHARACTER
    people.font.bold = True
    people.font.name = FONT_BODY_BOLD
    people.font.color.rgb = RGBColor(0, 0, 0)


# ---------------------------------------------------------------------------
# Reading sheet document configuration
# ---------------------------------------------------------------------------

def configure_reading_sheet_document(doc: Document):
    """Set up page layout and styles for a reading sheet.

    Uses the same style names as the bulletin but with larger fonts
    (16pt body, 14pt rubrics) and red rubric/heading colors.
    Mirror margins are enabled for booklet printing.
    """
    _setup_reading_sheet_page(doc)
    _create_styles(doc, style_defs=_RS_STYLE_DEFS)


def _setup_reading_sheet_page(doc: Document):
    """Configure page size and mirror margins for booklet printing.

    Mirror margins swap left/right on even pages:
      Odd pages:  left (inside) = 0.75", right (outside) = 0.5"
      Even pages: left (outside) = 0.5", right (inside) = 0.75"
    """
    section = doc.sections[0]
    section.page_width = Inches(PAGE_WIDTH_INCHES)
    section.page_height = Inches(PAGE_HEIGHT_INCHES)
    # With mirror margins, "left" = inside (binding edge), "right" = outside
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.5)
    section.top_margin = Inches(MARGIN_INCHES)
    section.bottom_margin = Inches(MARGIN_INCHES)

    # Enable mirror margins via Word XML settings
    W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    settings_elem = doc.settings.element
    mirror = parse_xml(f'<w:mirrorMargins {nsdecls("w")}/>')
    settings_elem.append(mirror)
