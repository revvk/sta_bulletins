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
    ("Reading (Poetry)", FONT_BODY, 11, False, False,
     WD_ALIGN_PARAGRAPH.LEFT, 0.455, -0.135, 0, 0, 1.0),

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

# Styles that get the left + right tab stops.
_STYLES_WITH_TABS = {"Body", "Body - Dialogue"}


def _create_styles(doc: Document):
    """Register all custom paragraph and character styles in the document."""

    # --- Paragraph styles ---
    for (name, font_name, size_pt, bold, italic, alignment,
         left_in, first_in, sp_before_pt, sp_after_pt,
         line_spacing) in _STYLE_DEFS:

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
        style.font.color.rgb = RGBColor(0, 0, 0)  # Explicit black; overrides theme

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
