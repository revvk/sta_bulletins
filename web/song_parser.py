"""
Convert pasted plain-text or uploaded markdown into a song dict that
matches the shape of entries in ``bulletin/data/hymns/songs.yaml``.

This is a *pure function* module — it imports nothing from the rest of
the codebase, has no side effects, and produces ordinary Python dicts.
The web app calls it from a form handler; tests call it on golden
fixtures from ``source_documents/*.md``.

Two input formats are supported:

1. **Paste-text** (the simpler path most users will reach for):

   - A blank line separates sections.
   - A line beginning with ``Chorus:``, ``Refrain:``, or ``Antiphon:``
     marks the *following* lines as a ``chorus`` section (the keyword
     and any trailing colon are stripped).
   - A line beginning with ``Tag:`` marks the following lines as a
     ``tag`` section.
   - Otherwise, each block is a ``verse``.
   - Single newlines within a block become entries in ``lines:``.

2. **Markdown** (matches the format Andrew already produces by hand,
   e.g. ``source_documents/Alleluia, alleluia, alleluia, ...md``):

   - The first ``#`` heading (with optional ``**bold**``) becomes the
     ``title``.
   - An italic-only line containing ``Hymnal`` and a ``#NNN`` reference
     becomes ``hymnal_number`` + ``hymnal_name: Hymnal 1982``.
   - Other italic-only lines (Antiphon / Chorus / Refrain / Tag) start
     a new section, with type inferred:
       * ``antiphon``, ``refrain``, ``chorus`` → ``chorus``
       * ``tag`` → ``tag``
       * anything else → ``chorus`` (best-effort fallback)
   - Plain paragraphs (separated by blank lines) become ``verse``
     sections.
   - Markdown line breaks (trailing two-space, trailing ``\\``, or a
     literal ``  ↵``) are normalized to single newlines within a block.
   - Markdown escapes (``\\!``, ``\\#``, ``\\.``) are unescaped.
   - ``**bold**`` and ``*italic*`` markers in body text are stripped
     (we don't carry inline emphasis into the YAML — bulletin styles
     handle that at render time).

The output shape:

.. code-block:: python

    {
        "title": "All glory, laud, and honor",
        "hymnal_number": "154",        # only if found
        "hymnal_name": "Hymnal 1982",  # only if hymnal_number is set
        "services": "9am",              # only if explicitly provided
        "sections": [
            {"type": "verse",  "lines": [...]},
            {"type": "chorus", "lines": [...]},
            ...
        ],
    }

Field order in the dict is intentional so the YAML serializer (the web
app uses ``ruamel.yaml``) emits keys in the same order as the
hand-written entries already in ``songs.yaml``.
"""

from __future__ import annotations

import re
from typing import Optional


# Section types that exist in songs.yaml today. Everything else gets
# normalized to one of these.
_VERSE = "verse"
_CHORUS = "chorus"
_TAG = "tag"

# Keywords (case-insensitive) that introduce a non-verse section.
# The mapping value is the section type to emit.
_LABEL_TO_TYPE: dict[str, str] = {
    "chorus":   _CHORUS,
    "refrain":  _CHORUS,
    "antiphon": _CHORUS,
    "tag":      _TAG,
}


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def parse_paste(
    text: str,
    *,
    title: str,
    hymnal_number: Optional[str] = None,
    hymnal_name: Optional[str] = None,
    services: Optional[str] = None,
) -> dict:
    """Parse plain-text lyrics from a paste-into-textarea form.

    The ``title``, ``hymnal_number``, ``hymnal_name``, and ``services``
    arguments come from separate form fields, not from the lyrics text
    itself — this keeps the textarea purely about the words.
    """
    sections = _sections_from_blocks(_split_blocks(_normalize_paste(text)))
    return _assemble(
        title=title,
        hymnal_number=hymnal_number,
        hymnal_name=hymnal_name,
        services=services,
        sections=sections,
    )


def parse_markdown(
    md: str,
    *,
    services: Optional[str] = None,
) -> dict:
    """Parse a markdown-formatted lyrics file.

    The title and hymnal reference are extracted from the markdown
    itself (first ``#`` heading and the italic ``*Hymnal 1982* \\#NNN``
    line, respectively). The optional ``services`` argument lets the
    caller force a service filter that the markdown doesn't carry.
    """
    md = _normalize_markdown(md)

    title = _extract_md_title(md)
    hymnal_number, hymnal_name = _extract_md_hymnal(md)

    # Strip the lines we've already consumed (title + hymnal) so they
    # don't accidentally show up as a verse.
    body = _strip_consumed_lines(md, title=title, hymnal_number=hymnal_number)

    sections = _sections_from_blocks(_split_blocks(body), markdown_mode=True)

    return _assemble(
        title=title or "",
        hymnal_number=hymnal_number,
        hymnal_name=hymnal_name,
        services=services,
        sections=sections,
    )


# ---------------------------------------------------------------------------
# Output assembly
# ---------------------------------------------------------------------------

def _assemble(
    *,
    title: str,
    hymnal_number: Optional[str],
    hymnal_name: Optional[str],
    services: Optional[str],
    sections: list[dict],
) -> dict:
    """Build the final dict in the same key order as songs.yaml entries."""
    out: dict = {"title": title}

    if services:
        out["services"] = services
    if hymnal_number:
        out["hymnal_number"] = str(hymnal_number)
        # Default the hymnal name when only a number was provided.
        out["hymnal_name"] = hymnal_name or "Hymnal 1982"

    out["sections"] = sections
    return out


# ---------------------------------------------------------------------------
# Block splitting (shared between paste and markdown paths)
# ---------------------------------------------------------------------------

def _split_blocks(text: str) -> list[list[str]]:
    """Split normalized text into blocks of non-empty lines.

    A block is one or more consecutive non-blank lines. Blank lines
    (any whitespace-only line) act as block separators. Returns a list
    of blocks, where each block is a list of stripped lines.
    """
    blocks: list[list[str]] = []
    current: list[str] = []
    for raw in text.splitlines():
        line = raw.rstrip()
        if line.strip():
            current.append(line.lstrip())
        else:
            if current:
                blocks.append(current)
                current = []
    if current:
        blocks.append(current)
    return blocks


def _sections_from_blocks(
    blocks: list[list[str]],
    *,
    markdown_mode: bool = False,
) -> list[dict]:
    """Convert blocks of text into a list of section dicts.

    Both input formats share this logic. The differences:

    - **Paste mode** detects sections from a leading keyword line
      (``Chorus:``, ``Refrain:``, ``Tag:``) — the keyword sits on its
      own line or as the first line of the block.
    - **Markdown mode** also detects sections from italic-only
      single-line blocks (e.g. ``*Antiphon (at the end)*``), with the
      following block(s) becoming that section's ``lines``.
    """
    sections: list[dict] = []
    pending_label: Optional[str] = None  # set by a markdown italic-label block

    for block in blocks:
        # Markdown-only: a leading italic-only line ("*Antiphon (at the
        # beginning)*") is a section label. The rest of the block (if
        # any) is that section's body; if the block is just the label,
        # the *next* block becomes its body.
        if markdown_mode:
            italic = _italic_only(block[0])
            if italic is not None:
                tail = block[1:]
                if tail:
                    sections.append({
                        "type": _LABEL_TO_TYPE.get(italic, _CHORUS),
                        "lines": _clean_lines(tail, markdown_mode=True),
                    })
                    pending_label = None
                else:
                    pending_label = italic
                continue

        # Paste-mode: a leading "Chorus:" / "Refrain:" / "Tag:" line.
        first = block[0]
        m = re.match(r"^(chorus|refrain|antiphon|tag)\s*:?\s*(.*)$",
                     first, flags=re.IGNORECASE)
        if m and not _looks_like_lyric(first[m.end(1):]):
            label = m.group(1).lower()
            remainder = m.group(2).strip()
            tail = block[1:]
            lines = ([remainder] if remainder else []) + tail
            if lines:
                sections.append({
                    "type": _LABEL_TO_TYPE.get(label, _CHORUS),
                    "lines": _clean_lines(lines, markdown_mode=markdown_mode),
                })
            continue

        # Otherwise: this block is the body of either a pending labeled
        # section (markdown) or a plain verse.
        section_type: str
        if pending_label is not None:
            section_type = _LABEL_TO_TYPE.get(pending_label, _CHORUS)
            pending_label = None
        else:
            section_type = _VERSE

        sections.append({
            "type": section_type,
            "lines": _clean_lines(block, markdown_mode=markdown_mode),
        })

    return sections


# ---------------------------------------------------------------------------
# Line cleaning
# ---------------------------------------------------------------------------

def _clean_lines(lines: list[str], *, markdown_mode: bool) -> list[str]:
    """Final clean-up applied to each line of a section."""
    out = []
    for line in lines:
        text = line.strip()
        if markdown_mode:
            text = _strip_markdown_inline(text)
        # Drop any stray leading bullet that some pastes carry over from
        # rich-text sources.
        text = re.sub(r"^[•·\-\*]\s+", "", text)
        if text:
            out.append(text)
    return out


# ---------------------------------------------------------------------------
# Paste normalization
# ---------------------------------------------------------------------------

def _normalize_paste(text: str) -> str:
    """Normalize line endings and collapse extra blank lines for pastes."""
    # Normalize line endings.
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    # Collapse runs of 3+ blank lines into a single blank-line separator
    # so the user can paste loosely-spaced text.
    text = re.sub(r"\n\s*\n\s*\n+", "\n\n", text)
    return text.strip("\n")


# ---------------------------------------------------------------------------
# Markdown normalization
# ---------------------------------------------------------------------------

# Patterns we use to recognize bits of the markdown format.
_RE_MD_TITLE = re.compile(r"^\s*#\s+(.*?)\s*$", flags=re.MULTILINE)
_RE_MD_HYMNAL = re.compile(
    r"^\s*\*([^*]*?Hymnal[^*]*?)\*\s*\\?#?\s*(\d+[A-Za-z]?)\s*$",
    flags=re.IGNORECASE | re.MULTILINE,
)
_RE_MD_BOLD = re.compile(r"\*\*(.+?)\*\*")
_RE_MD_ITALIC = re.compile(r"(?<!\*)\*(?!\*)([^*\n]+?)(?<!\*)\*(?!\*)")


def _normalize_markdown(md: str) -> str:
    """Collapse markdown line breaks and escapes into plain text.

    - ``\\r\\n`` → ``\\n``
    - Trailing ``  `` (two-space) → newline (already a newline; just
      strip the trailing spaces so we don't double-count)
    - Trailing ``\\`` → newline
    - ``\\!`` → ``!``, ``\\#`` → ``#``, ``\\.`` → ``.``, etc.
    """
    md = md.replace("\r\n", "\n").replace("\r", "\n")
    # A trailing backslash is markdown's "hard line break". The newline
    # is already there, so just remove the backslash.
    md = re.sub(r"\\\n", "\n", md)
    # Remove the two-space trailing marker; the newline that follows it
    # already does the line-break job for our purposes.
    md = re.sub(r"  +\n", "\n", md)
    # Unescape common markdown punctuation escapes. Only do the
    # punctuation we actually see in lyrics — don't touch \\ or anything
    # exotic.
    md = re.sub(r"\\([!#.,;:?\"'])", r"\1", md)
    return md.strip("\n")


def _italic_only(line: str) -> Optional[str]:
    """If ``line`` is a single ``*…*`` italic span (and nothing else),
    return the inner text lowercased. Otherwise return None.

    Used for markdown section labels like ``*Antiphon (at the end)*``.
    """
    stripped = line.strip()
    m = re.fullmatch(r"\*([^*]+)\*", stripped)
    if not m:
        return None
    inner = m.group(1).strip().lower()
    # Pull out the leading word (before any parenthetical) and check
    # that it's a label we recognize.
    head = re.match(r"^([a-z]+)\b", inner)
    if not head:
        return None
    word = head.group(1)
    if word in _LABEL_TO_TYPE:
        return word
    return None


def _looks_like_lyric(rest_of_line: str) -> bool:
    """Heuristic: after a leading ``chorus:`` / ``refrain:`` / etc., is
    the rest of the *first* line itself a lyric line?

    If yes, treat the whole block as a verse (don't split off the
    keyword as a label) — this avoids false positives on lines like
    ``Chorus of angels, raise your voice…`` where "Chorus" is part of
    the lyric.
    """
    text = rest_of_line.strip()
    # If there's no colon and the rest of the line continues with words
    # that look like lyric content, it's probably not a label.
    return bool(text) and not text.startswith(":") and len(text) > 30


# ---------------------------------------------------------------------------
# Markdown title / hymnal extraction
# ---------------------------------------------------------------------------

def _extract_md_title(md: str) -> str:
    """Grab the first ``# Heading`` in the markdown, stripping bold."""
    m = _RE_MD_TITLE.search(md)
    if not m:
        return ""
    title = m.group(1).strip()
    title = _strip_markdown_inline(title)
    return title


def _extract_md_hymnal(md: str) -> tuple[Optional[str], Optional[str]]:
    """Find the hymnal reference line and split it into (number, name).

    Recognizes lines like:

    - ``*Hymnal 1982* \\#208``
    - ``*Hymnal 1982* #208``
    - ``*Hymnal 1982 #208*``
    """
    m = _RE_MD_HYMNAL.search(md)
    if not m:
        # Try the all-italic form: *Hymnal 1982 #208*
        m2 = re.search(
            r"^\s*\*([^*]*?Hymnal[^*]*?)\s*\\?#?\s*(\d+[A-Za-z]?)\s*\*\s*$",
            md, flags=re.IGNORECASE | re.MULTILINE,
        )
        if not m2:
            return None, None
        m = m2
    name_part = m.group(1).strip()
    number = m.group(2).strip()
    # Normalize the name to "Hymnal 1982" specifically; otherwise fall
    # back to whatever was written.
    name = "Hymnal 1982" if "1982" in name_part else name_part
    return number, name


def _strip_consumed_lines(md: str, *, title: str,
                          hymnal_number: Optional[str]) -> str:
    """Remove the title heading and hymnal-reference line from the body.

    We do this by line-matching rather than by index so we don't break
    when the file has unusual blank-line padding.
    """
    out_lines = []
    for line in md.splitlines():
        s = line.strip()
        # Drop the title line.
        if title and s.startswith("#") and \
                _strip_markdown_inline(s.lstrip("#").strip()) == title:
            continue
        # Drop the hymnal-reference line.
        if hymnal_number and "hymnal" in s.lower() and hymnal_number in s:
            continue
        out_lines.append(line)
    return "\n".join(out_lines)


def _strip_markdown_inline(text: str) -> str:
    """Strip ``**bold**`` and ``*italic*`` markers from a string.

    Preserves a single literal ``*`` (e.g., footnote marker) since the
    pattern requires non-empty content between markers.
    """
    text = _RE_MD_BOLD.sub(r"\1", text)
    text = _RE_MD_ITALIC.sub(r"\1", text)
    return text
