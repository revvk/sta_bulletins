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

from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

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
        "{{TIME}}": service_time,
        "{{TITLE OF DAY}}": liturgical_title,
        "{{SUBTITLE}}": subtitle,
    }

    _replace_all_placeholders(doc, replacements)
    return doc


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
