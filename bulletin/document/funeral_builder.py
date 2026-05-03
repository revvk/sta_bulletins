"""
Top-level orchestrator for funeral / memorial bulletin generation.

Parallels ``BulletinBuilder`` (the Sunday/special-service builder) but
on the YAML-driven side: it consumes a per-service ``FuneralData``
dataclass + the existing scripture cache + the existing song-lookup
pipeline, and produces a Document.

Usage
=====

    from bulletin.document.funeral_builder import FuneralBuilder
    from bulletin.sources.funeral_data import load_service

    fd = load_service("2026-01-31-cox")
    builder = FuneralBuilder(fd, scripture_readings, song_lookup_fn)
    doc = builder.build()
    doc.save("output/2026-01-31 - Burial of the Dead - Annette Cox.docx")

The CLI driver (``generate.py --funeral <slug>``) wraps this with a
small helper that fetches the readings and supplies the song lookup.
"""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

from bulletin.document.styles import configure_document
from bulletin.document.templates import (
    _replace_all_placeholders,
    _pin_floating_shapes_to_first_paragraph,
    append_back_cover,
    setup_footers,
)
from bulletin.document.sections.burial import add_burial_service
from bulletin.sources.funeral_data import FuneralData


_TEMPLATES_DIR = Path(__file__).resolve().parent.parent.parent / "templates"
_FUNERAL_COVER_TEMPLATE = "front_cover_funeral.docx"


class FuneralBuilder:
    """Builds a funeral / memorial bulletin .docx from a ``FuneralData``.

    Args:
        fd:                  A loaded ``FuneralData``.
        scripture_readings:  Dict ``{reference: ScriptureReading}`` for
                             every reading the YAML names — the CLI
                             driver fetches these via the existing
                             scripture cache before constructing the
                             builder.
        song_lookup_fn:      ``callable(title, service) -> dict | None``
                             — same shape as the Sunday pipeline.
        report:              Optional ``RunReport`` to record warnings.
    """

    def __init__(self, fd: FuneralData, scripture_readings: dict,
                 song_lookup_fn, report=None):
        self.fd = fd
        self.scripture = scripture_readings
        self.song_lookup = song_lookup_fn
        self.report = report

    # ------------------------------------------------------------------
    # build()
    # ------------------------------------------------------------------

    def build(self) -> Document:
        """Open the cover template, substitute placeholders, render the
        liturgy, append the back cover, and return the Document."""
        doc = self._load_cover()
        configure_document(doc)

        add_burial_service(doc, self.fd, self.scripture, self.song_lookup)

        # The Sunday back_cover.docx has parish-staff content that
        # doesn't belong on a funeral bulletin (Welcome blurb, staff
        # directory, vestry list). burial.py instead adds the
        # participants page + BCP boilerplate inline at the end of
        # the body, and a parish-logo back-cover template is a TODO
        # follow-up. Skip the Sunday back-cover for now.
        setup_footers(
            doc,
            date_str=self.fd.service_date_long,
            service_time="",  # funerals don't print a service-time footer
            liturgical_title=self._footer_title(),
        )
        return doc

    # ------------------------------------------------------------------
    # Cover template handling
    # ------------------------------------------------------------------

    def _load_cover(self) -> Document:
        """Load the two-page funeral cover template and substitute the
        five text placeholders (``{{DATE}}``, ``{{SUBTITLE}}``,
        ``{{NAME}}``, ``{{LIFE_DATES}}``, ``{{BIO}}``).

        The photo placeholder is an embedded image in the template that
        the user swaps out manually in the post-generation edit pass.
        Services with no bio AND no photo: the user deletes the inner
        page in the same edit pass — simpler than conditional template
        assembly.
        """
        path = _TEMPLATES_DIR / _FUNERAL_COVER_TEMPLATE
        if not path.exists():
            raise FileNotFoundError(
                f"Funeral cover template not found: {path}\n"
                f"Expected at: {path}"
            )
        doc = Document(str(path))

        bio = self.fd.deceased.get("bio") or ""
        # Multi-paragraph bios: collapse to a single string with double
        # spaces between paragraphs. The user's hand-edit pass splits
        # them back into paragraphs in Word — supporting true multi-
        # paragraph substitution requires custom XML manipulation that
        # we'll add in a follow-up if it turns out to be tedious.
        bio_collapsed = "  ".join(p.strip() for p in bio.split("\n\n") if p.strip())

        replacements = {
            "{{DATE}}":       self.fd.service_date_long,
            "{{SUBTITLE}}":   self.fd.cover_subtitle_resolved,
            "{{NAME}}":       self.fd.deceased.get("full_name", ""),
            "{{LIFE_DATES}}": self.fd.life_dates,
            "{{BIO}}":        bio_collapsed,
        }
        _replace_all_placeholders(doc, replacements)
        _pin_floating_shapes_to_first_paragraph(doc)
        return doc

    def _footer_title(self) -> str:
        """Footer left text — matches the convention in the existing
        bulletins: 'Burial of the Dead, Rite II' / 'Memorial Service'
        / etc."""
        kind = self.fd.service_kind
        rite = f", Rite {self.fd.rite}"
        if kind == "burial":
            return f"Burial of the Dead{rite}"
        if kind == "memorial":
            return f"Memorial Service{rite}"
        if kind == "committal":
            return f"Graveside Service{rite}"
        return f"Funeral Service{rite}"

    # ------------------------------------------------------------------
    # Output filename
    # ------------------------------------------------------------------

    def output_filename(self) -> str:
        """Suggested filename for the generated .docx — mirrors the
        convention used by the existing Pages bulletins, e.g.
        '2026-01-31 - Burial of the Dead - Annette Cox.docx'."""
        kind_label = {
            "burial":   "Burial of the Dead",
            "memorial": "Memorial Service",
            "committal": "Graveside Service",
        }.get(self.fd.service_kind, "Funeral Service")
        date = self.fd.service["date"]
        if hasattr(date, "isoformat"):
            date = date.isoformat()
        first = self.fd.deceased.get("preferred_name", "") or ""
        last = self.fd.deceased.get("full_name", "").split()[-1]
        name = f"{first} {last}".strip() or self.fd.deceased.get("full_name", "")
        return f"{date} - {kind_label} - {name}.docx"
