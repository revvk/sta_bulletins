"""
Fetch poetry structure from BibleGateway.com (NRSVUE).

BibleGateway properly formats poetry with indent levels, which oremus.org
does not. We use BibleGateway solely for structural information — which
verses contain poetry and at what indent level — then merge that structure
with the authoritative NRSV text from oremus.org.

HTML structure on BibleGateway:
  - Prose: <p> tags containing <span class="text Book-Ch-V"> elements
  - Poetry: <div class="poetry"> containing <p class="line"> elements
    - Indent 0: bare <span class="text ..."> (no indent wrapper)
    - Indent 1: <span class="indent-1"> wrapping the text span
    - Indent 2: <span class="indent-2"> wrapping the text span
  - Verse numbers: <sup class="versenum">5 </sup>
  - Chapter numbers: <span class="chapternum">21 </span>
"""

import re
import requests
from bs4 import BeautifulSoup, NavigableString, Tag


BIBLEGATEWAY_URL = "https://www.biblegateway.com/passage/"


def fetch_poetry_structure(reference: str) -> list[dict] | None:
    """Fetch a passage from BibleGateway and extract its prose/poetry structure.

    Returns a list of segments, each either:
      {"type": "prose", "verses": [int, ...]}
      {"type": "poetry", "lines": [{"verse": int, "indent": 0|1|2, "text": str}, ...]}

    Returns None if the passage contains no poetry.
    """
    params = {"search": reference, "version": "NRSVUE"}
    response = requests.get(BIBLEGATEWAY_URL, params=params, timeout=30)
    response.raise_for_status()

    soup = BeautifulSoup(response.text, "html.parser")
    return _parse_structure(soup)


def _get_verse_num(element: Tag) -> int | None:
    """Extract verse number from a versenum sup or chapternum span."""
    versenum = element.find("sup", class_="versenum")
    if versenum:
        num = versenum.get_text().strip()
        if num.isdigit():
            return int(num)
    chapternum = element.find("span", class_="chapternum")
    if chapternum:
        # Chapter number marks verse 1
        return 1
    return None


def _extract_verse_from_class(classes: list[str]) -> int | None:
    """Extract verse number from BibleGateway span class like 'Phil-2-5'."""
    for cls in classes:
        # Pattern: Book-Chapter-Verse (e.g., "Phil-2-5", "Matt-21-1")
        match = re.match(r'^[A-Za-z]+-\d+-(\d+)$', cls)
        if match:
            return int(match.group(1))
    return None


def _parse_structure(soup: BeautifulSoup) -> list[dict] | None:
    """Parse BibleGateway HTML into prose/poetry segments."""
    passage_div = soup.find("div", class_="passage-text")
    if not passage_div:
        return None

    # Find the actual content div
    content_div = passage_div.find("div", class_="result-text-style-normal")
    if not content_div:
        content_div = passage_div

    segments = []
    has_poetry = False

    # Process direct children of the content div to identify prose vs poetry blocks
    for child in content_div.children:
        if not isinstance(child, Tag):
            continue

        # Skip headings, footnotes sections, crossrefs
        if child.name in ('h1', 'h2', 'h3', 'h4'):
            continue
        if child.get('class') and any(c in child.get('class', [])
                                       for c in ['footnotes', 'crossrefs',
                                                  'full-chap-link', 'passage-other-trans',
                                                  'publisher-info-bottom']):
            continue

        # Poetry block
        if child.name == 'div' and 'poetry' in child.get('class', []):
            has_poetry = True
            lines = _parse_poetry_block(child)
            if lines:
                # Merge with previous poetry segment if adjacent
                if segments and segments[-1]["type"] == "poetry":
                    segments[-1]["lines"].extend(lines)
                else:
                    segments.append({"type": "poetry", "lines": lines})

        # Prose paragraph
        elif child.name == 'p':
            classes = child.get('class', [])
            # Skip non-content paragraphs
            if any(c in classes for c in ['passage-display', 'passage-display-bcv']):
                continue

            verses = _extract_prose_verses(child)
            if verses:
                # Merge with previous prose segment if adjacent
                if segments and segments[-1]["type"] == "prose":
                    segments[-1]["verses"].extend(verses)
                else:
                    segments.append({"type": "prose", "verses": verses})

    if not has_poetry:
        return None

    return segments


def _parse_poetry_block(div: Tag) -> list[dict]:
    """Parse a <div class="poetry"> block into lines with indent levels.

    Each line is: {"verse": int, "indent": 0|1|2, "text": str}
    """
    lines = []
    current_verse = None

    for p in div.find_all('p', class_='line'):
        # Walk through the content of this <p class="line">
        # Lines are separated by <br/> tags
        # Each line is either bare text (indent 0), or wrapped in indent-1/indent-2 spans

        line_parts = _split_poetry_line_element(p)
        for indent, text_elements in line_parts:
            # Extract verse number if present
            for elem in text_elements:
                if isinstance(elem, Tag):
                    v = _get_verse_num(elem)
                    if v is not None:
                        current_verse = v
                    # Check text spans for verse class
                    if 'text' in elem.get('class', []):
                        v = _extract_verse_from_class(elem.get('class', []))
                        if v is not None and current_verse is None:
                            current_verse = v

            # Get the text content
            text = ""
            for elem in text_elements:
                if isinstance(elem, Tag):
                    # Skip the indent-breaks spans (just &nbsp; padding)
                    if any(c.endswith('-breaks') for c in elem.get('class', [])):
                        continue
                    # Skip footnote/crossref markers
                    if elem.name == 'sup' and ('footnote' in elem.get('class', []) or
                                                'crossreference' in elem.get('class', [])):
                        continue
                    text += elem.get_text()
                elif isinstance(elem, NavigableString):
                    text += str(elem)

            text = text.strip()
            # Remove verse numbers from the text (we track them separately)
            text = re.sub(r'^\d+\s*', '', text)
            # Strip cross-reference markers like (A), (B), etc.
            text = re.sub(r'\([A-Z]+\)\s*$', '', text).rstrip()
            if text:
                lines.append({
                    "verse": current_verse,
                    "indent": indent,
                    "text": text,
                })

    return lines


def _split_poetry_line_element(p: Tag) -> list[tuple[int, list]]:
    """Split a <p class="line"> into individual lines separated by <br/>.

    Returns list of (indent_level, [elements]) tuples.
    """
    lines = []
    current_elements = []
    current_indent = 0

    for child in p.children:
        if isinstance(child, Tag) and child.name == 'br':
            if current_elements:
                lines.append((current_indent, current_elements))
            current_elements = []
            current_indent = 0
            continue

        # Determine indent level
        if isinstance(child, Tag):
            classes = child.get('class', [])
            if 'indent-2' in classes:
                current_indent = 2
                current_elements.append(child)
                continue
            elif 'indent-1' in classes:
                current_indent = 1
                current_elements.append(child)
                continue

        current_elements.append(child)

    if current_elements:
        lines.append((current_indent, current_elements))

    return lines


def _extract_prose_verses(p: Tag) -> list[int]:
    """Extract verse numbers from a prose <p> element."""
    verses = []
    for span in p.find_all('span', class_='text'):
        v = _extract_verse_from_class(span.get('class', []))
        if v is not None and v not in verses:
            verses.append(v)
    # Also check for chapternum (verse 1)
    if p.find('span', class_='chapternum'):
        if 1 not in verses:
            verses.insert(0, 1)
    return verses
