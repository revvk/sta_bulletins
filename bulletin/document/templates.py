"""
Template handling for pre-designed .docx cover pages.

Opens a Word template, replaces placeholder strings (e.g. {{DATE}}),
and returns a Document ready for additional content to be appended.

Placeholder replacement handles the common case where Word splits a
placeholder across multiple XML runs (due to spell-check or grammar
marks).  The algorithm concatenates all run text in a paragraph, does
the replacement on the joined string, puts the result into the first
run, and clears the rest.
"""

from copy import deepcopy
from pathlib import Path

from docx import Document
from docx.oxml.ns import nsdecls, qn
from docx.oxml import parse_xml

# Resolve templates directory relative to the project root.
# bulletin/document/templates.py  ->  ../../templates/
_TEMPLATES_DIR = Path(__file__).resolve().parent.parent.parent / "templates"

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def load_front_cover(
    date_str: str,
    service_time: str,
    liturgical_title: str,
    subtitle: str = " ",
) -> Document:
    """Open the front-cover template and fill in dynamic fields.

    Args:
        date_str:          e.g. "March 1, 2026"
        service_time:      e.g. "9 am"
        liturgical_title:  e.g. "Second Sunday in Lent"
        subtitle:          Optional subtitle line; defaults to a single
                           space to preserve the template's spacing.

    Returns:
        A python-docx Document with placeholders replaced.
    """
    template_path = _TEMPLATES_DIR / "front_cover.docx"
    if not template_path.exists():
        raise FileNotFoundError(f"Front cover template not found: {template_path}")

    doc = Document(str(template_path))

    replacements = {
        "{{DATE}}": date_str,
        "{{SERVICE_TIME}}": service_time,
        "{{TITLE OF DAY}}": liturgical_title,
        "{{SUBTITLE}}": subtitle,
    }

    _replace_all_placeholders(doc, replacements)
    return doc


def append_back_cover(doc: Document):
    """Append the back-cover template as a new page section at the end.

    Inserts a section break (next-page) after the existing content, then
    copies all content and page setup from the back-cover template into
    the new section.
    """
    template_path = _TEMPLATES_DIR / "back_cover.docx"
    if not template_path.exists():
        raise FileNotFoundError(f"Back cover template not found: {template_path}")

    back_doc = Document(str(template_path))
    body = doc.element.body
    back_body = back_doc.element.body

    # Move the main document's trailing <w:sectPr> (which defines the
    # current last section's page setup) into the last paragraph's <w:pPr>.
    # This "closes" the current section and makes room for the back cover's
    # section to become the new trailing section.
    main_sect_pr = body.find(qn("w:sectPr"))
    if main_sect_pr is not None:
        body.remove(main_sect_pr)

        # Set the break type to "nextPage" so the back cover starts on a
        # fresh page.
        sect_type = main_sect_pr.find(qn("w:type"))
        if sect_type is not None:
            sect_type.set(qn("w:val"), "nextPage")
        else:
            main_sect_pr.insert(
                0, parse_xml(f'<w:type {nsdecls("w")} w:val="nextPage"/>')
            )

        # Find the last paragraph and attach the sectPr to its pPr.
        paragraphs = list(body.iterchildren(qn("w:p")))
        if paragraphs:
            last_p = paragraphs[-1]
            pPr = last_p.find(qn("w:pPr"))
            if pPr is None:
                pPr = parse_xml(f'<w:pPr {nsdecls("w")}/>')
                last_p.insert(0, pPr)
            pPr.append(main_sect_pr)

    # Append all elements from the back-cover template (paragraphs,
    # tables, and the trailing sectPr which defines page setup).
    for element in list(back_body):
        body.append(deepcopy(element))


# ------------------------------------------------------------------
# Internals
# ------------------------------------------------------------------

def _replace_all_placeholders(doc: Document, replacements: dict[str, str]):
    """Walk every <w:p> in the document (body + text boxes) and replace."""
    body = doc.element.body
    for para in body.iter(qn("w:p")):
        _replace_in_paragraph(para, replacements)


def _replace_in_paragraph(para_elem, replacements: dict[str, str]):
    """Replace placeholders in one paragraph, handling cross-run splits.

    Strategy: concatenate all run texts, do the replacement on the
    joined string, deposit the result into the first <w:t> element,
    and blank out the remaining <w:t> elements.  This preserves the
    first run's character formatting (font, colour, size, etc.).
    """
    runs = para_elem.findall(f"{{{_W_NS}}}r")
    if not runs:
        return

    # Collect text from each run's <w:t>
    t_elements = []
    for run in runs:
        t = run.find(f"{{{_W_NS}}}t")
        t_elements.append(t)

    full_text = "".join((t.text or "") if t is not None else "" for t in t_elements)

    # Check whether any placeholder appears in this paragraph
    changed = False
    for placeholder, replacement in replacements.items():
        if placeholder in full_text:
            full_text = full_text.replace(placeholder, replacement)
            changed = True

    if not changed:
        return

    # Deposit the replaced text into the first <w:t>, blank the rest.
    first_set = False
    for t in t_elements:
        if t is None:
            continue
        if not first_set:
            t.text = full_text
            t.set(qn("xml:space"), "preserve")
            first_set = True
        else:
            t.text = ""
