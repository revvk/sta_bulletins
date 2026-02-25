"""
Defines all paragraph and character styles used in the bulletin.

These styles replicate the formatting found in the existing St. Andrew's
bulletins. The document uses half-letter (7" x 8.5") pages with 0.5" margins.

Two font families are used:
  - Adobe Garamond Pro: body text, prayers, scripture, rubrics
  - Gill Sans variants: headings, lyrics, headers/footers
"""

from docx import Document
from docx.shared import Pt, Inches, Emu, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn

from bulletin.config import (
    FONT_BODY,
    FONT_HEADING,
    FONT_HEADING2,
    FONT_LYRICS,
    FONT_HEADER_FOOTER,
    PAGE_WIDTH_INCHES,
    PAGE_HEIGHT_INCHES,
    MARGIN_INCHES,
)


def create_document() -> Document:
    """Create a new Document with all bulletin styles and page setup."""
    doc = Document()
    _setup_page(doc)
    _create_styles(doc)
    return doc


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
    ("Heading", FONT_HEADING, 14, False, False,
     WD_ALIGN_PARAGRAPH.CENTER, 0, 0, 0, 0, 0.9),

    # Section headers: "Processional", "Sermon", "The Peace", etc.
    ("Heading 2", FONT_HEADING2, 13, False, False,
     WD_ALIGN_PARAGRAPH.LEFT, 0, 0, 0, 5, None),

    # Main body text (prayers, collect text, general content)
    ("Body", FONT_BODY, 12, False, False,
     WD_ALIGN_PARAGRAPH.LEFT, 0.32, 0, 3, 0, 1.0),

    # Celebrant dialogue lines: "Celebrant    The Lord be with you."
    ("Body - Celebrant", FONT_BODY, 12, False, False,
     WD_ALIGN_PARAGRAPH.LEFT, 1.04, -0.72, 3, 0, 1.0),

    # People's responses (bold): "People    And also with you."
    ("Body - People", FONT_BODY, 12, True, False,
     WD_ALIGN_PARAGRAPH.LEFT, 1.04, -0.72, 1, 0, 1.0),

    # Rubrics within the service: "After the reading, the Reader will say"
    ("Body - Rubric", FONT_BODY, 10, False, True,
     WD_ALIGN_PARAGRAPH.LEFT, 0.32, 0, 2, 0, 1.0),

    # Introductory rubrics: "Please stand.", "Be seated.", "Remain standing."
    ("Body - Introductory Rubric", FONT_BODY, 10, False, True,
     WD_ALIGN_PARAGRAPH.LEFT, 0.004, 0, 0, 0, 1.0),

    # Song lyrics (verse text; choruses use italic runs)
    ("Body - Lyrics", FONT_LYRICS, 12, False, False,
     WD_ALIGN_PARAGRAPH.LEFT, 0.5, -0.18, 0, 0, 1.0),

    # Scripture reading prose text
    ("Reading/Gospel Text", FONT_BODY, 11, False, False,
     WD_ALIGN_PARAGRAPH.LEFT, 0.32, 0, 0, 0, 1.0),

    # Scripture reading poetry (prophetic/poetic quotations)
    ("Reading (Poetry)", FONT_BODY, 11, False, False,
     WD_ALIGN_PARAGRAPH.LEFT, 0.455, -0.135, 0, 0, 1.0),

    # Psalm text (bold, all congregation reads)
    ("Psalm", FONT_BODY, 12, True, False,
     WD_ALIGN_PARAGRAPH.LEFT, 0.455, -0.13, 3, 0, 1.0),

    # People's recitation of long texts: Confession, Lord's Prayer
    ("Body - People Recitation", FONT_BODY, 12, True, False,
     WD_ALIGN_PARAGRAPH.LEFT, 0.32, 0, 6, 0, 1.1),

    # Nicene Creed (slightly different indent for the three articles)
    ("Body - People Recitation (Creed)", FONT_BODY, 12, True, False,
     WD_ALIGN_PARAGRAPH.LEFT, 0.455, -0.18, 6, 0, 1.1),

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
     WD_ALIGN_PARAGRAPH.CENTER, 0, 0, 0, 8, 1.0),

    # Header & Footer text
    ("Header & Footer", FONT_HEADER_FOOTER, 9, False, False,
     WD_ALIGN_PARAGRAPH.LEFT, 0, 0, 0, 0, 1.0),
]


def _create_styles(doc: Document):
    """Register all custom paragraph styles in the document."""
    for (name, font_name, size_pt, bold, italic, alignment,
         left_in, first_in, sp_before_pt, sp_after_pt,
         line_spacing) in _STYLE_DEFS:

        # If the style already exists (e.g., "Heading"), modify it;
        # otherwise create a new one.
        try:
            style = doc.styles[name]
        except KeyError:
            style = doc.styles.add_style(name, 1)  # 1 = WD_STYLE_TYPE.PARAGRAPH

        # Font
        style.font.name = font_name
        style.font.size = Pt(size_pt)
        style.font.bold = bold
        style.font.italic = italic
        style.font.color.rgb = RGBColor(0, 0, 0)  # Explicit black; overrides theme

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
