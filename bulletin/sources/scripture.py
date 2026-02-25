"""
Fetch scripture readings with verse numbers from bible.oremus.org.

The Google Sheet provides scripture references (e.g., "Genesis 12:1-4a").
This module fetches the full NRSV text with inline verse numbers.

Oremus HTML is somewhat malformed (nested <p> tags, inline <h2> headings),
so we take a stream-based approach: walk the entire bibletext div's
descendants and collect text, detecting verse numbers and paragraph breaks.
"""

import re
import time
from dataclasses import dataclass, field
from typing import Optional

import requests
from bs4 import BeautifulSoup, NavigableString, Comment, Tag

from bulletin.config import OREMUS_BASE_URL, OREMUS_PARAMS


@dataclass
class ScriptureReading:
    """A scripture reading with verse-numbered text."""
    reference: str              # e.g., "Genesis 12:1-4a"
    paragraphs: list[str]       # Prose paragraphs, verse nums inline
    poetry_lines: list[str]     # If poetry section exists, individual lines
    has_poetry: bool            # Whether the text contains poetic formatting

    @property
    def text(self) -> str:
        """Full text as a single string with paragraph breaks."""
        return "\n\n".join(self.paragraphs)


def fetch_reading(reference: str) -> ScriptureReading:
    """Fetch a scripture reading from the Oremus Bible Browser."""
    params = dict(OREMUS_PARAMS)
    params["passage"] = reference

    response = requests.get(OREMUS_BASE_URL, params=params, timeout=30)
    response.raise_for_status()

    soup = BeautifulSoup(response.text, "html.parser")
    return _parse_oremus_response(soup, reference)


def _parse_oremus_response(soup: BeautifulSoup, reference: str) -> ScriptureReading:
    """Parse the Oremus Bible Browser HTML.

    Strategy: walk all descendants of the bibletext div in document order.
    Collect text, inserting verse number markers. Detect paragraph breaks
    from <p> opening tags and <span class="vv"> (which indicate new paragraphs
    in the original text).
    """
    bibletext_div = soup.find("div", class_="bibletext")
    if not bibletext_div:
        raise ValueError(
            f"Could not find scripture text for '{reference}' in Oremus response"
        )

    # Walk all descendants and build a list of tokens
    tokens = []  # List of ("text", str) | ("verse", str) | ("para", None) | ("poetry_br", None)
    in_heading = False
    seen_first_verse = False

    for desc in bibletext_div.descendants:
        if isinstance(desc, Comment):
            continue

        if isinstance(desc, Tag):
            classes = desc.get("class", [])

            # Skip section headings
            if desc.name in ("h2", "h3", "h4"):
                in_heading = True
                continue

            # Verse numbers -- three CSS classes used:
            #   "cc" = chapter-start (e.g., "13" for chapter 13, verse 1)
            #          We convert this to the actual starting verse number.
            #   "vv" = paragraph-start verse (e.g., "5 " with nbsp)
            #          These also indicate a paragraph break in the original.
            #   "ww" = inline verse number (e.g., "2", "3", "4")
            if "vnumVis" in classes:
                num = desc.get_text().strip()
                if num:
                    if "cc" in classes:
                        # Chapter number -- convert to the starting verse
                        # Extract starting verse from the reference
                        start_verse = _get_start_verse(reference)
                        num = start_verse
                    if not seen_first_verse:
                        seen_first_verse = True
                    elif "vv" in classes:
                        # Paragraph-start verse = paragraph break
                        tokens.append(("para", None))
                    tokens.append(("verse", num))
                continue

            # <br/> in poetry = line break
            if desc.name == "br":
                tokens.append(("poetry_br", None))
                continue

            # <blockquote> = start of poetry section
            if desc.name == "blockquote":
                tokens.append(("para", None))
                continue

            # <p> tags can indicate paragraph breaks
            # But only if we've already seen content (not the initial wrapper)
            if desc.name == "p" and seen_first_verse:
                # Check if this <p> has a vv-class span child (those handle their own para break)
                vv_child = desc.find("span", class_="vv", recursive=False)
                cc_child = desc.find("span", class_="cc", recursive=False)
                if not vv_child and not cc_child:
                    tokens.append(("para", None))
                continue

        elif isinstance(desc, NavigableString):
            # Skip text inside headings
            if in_heading:
                parent = desc.parent
                while parent and parent != bibletext_div:
                    if parent.name in ("h2", "h3", "h4"):
                        break
                    parent = parent.parent
                if parent and parent.name in ("h2", "h3", "h4"):
                    continue
                in_heading = False

            # Skip text that is a direct child of verse number elements
            # (because .descendants yields both the tag and its text children)
            parent = desc.parent
            if parent and isinstance(parent, Tag):
                parent_classes = parent.get("class", [])
                if "vnumVis" in parent_classes:
                    continue
                # Also skip text inside thinspace spans
                if "thinspace" in parent_classes:
                    continue

            text = str(desc)
            if text.strip():
                tokens.append(("text", text))
            elif text == " ":
                tokens.append(("text", " "))

    # Debug: uncomment to see tokens
    # for t in tokens:
    #     print(t)

    # Now assemble tokens into paragraphs
    paragraphs = []
    poetry_lines = []
    has_poetry = False
    current = []

    for ttype, tval in tokens:
        if ttype == "para":
            text = "".join(current).strip()
            if text:
                paragraphs.append(text)
            current = []
        elif ttype == "verse":
            current.append(f"{tval} ")
        elif ttype == "text":
            current.append(tval)
        elif ttype == "poetry_br":
            has_poetry = True
            # Save current line and start new one
            line = "".join(current).strip()
            if line:
                poetry_lines.append(line)
            current = []

    # Flush remaining
    text = "".join(current).strip()
    if text:
        if has_poetry and poetry_lines:
            poetry_lines.append(text)
        paragraphs.append(text)

    # Clean up whitespace in paragraphs
    cleaned = []
    for p in paragraphs:
        # Collapse multiple spaces
        p = re.sub(r'  +', ' ', p).strip()
        if p:
            cleaned.append(p)

    return ScriptureReading(
        reference=reference,
        paragraphs=cleaned,
        poetry_lines=[re.sub(r'  +', ' ', l).strip() for l in poetry_lines if l.strip()],
        has_poetry=has_poetry,
    )


def _get_start_verse(reference: str) -> str:
    """Extract the starting verse number from a scripture reference.

    Examples:
        "Luke 13:1-9" -> "1"
        "Acts 2:1-21" -> "1"
        "1 Corinthians 10:1-13" -> "1"
        "Genesis 12:1-4a" -> "1"
        "Psalm 104:25-35" -> "25"
    """
    match = re.search(r':(\d+)', reference)
    if match:
        return match.group(1)
    return "1"


# ---------------------------------------------------------------------------
# Batch fetching with rate limiting
# ---------------------------------------------------------------------------

def fetch_readings(references: dict[str, str],
                   delay: float = 0.5) -> dict[str, ScriptureReading]:
    """Fetch multiple readings with a delay between requests.

    Args:
        references: Dict mapping label to reference,
                    e.g., {"reading": "Genesis 12:1-4a", "gospel": "John 3:1-17"}
        delay: Seconds to wait between requests (be nice to oremus.org)

    Returns:
        Dict mapping label to ScriptureReading.
    """
    results = {}
    for i, (label, ref) in enumerate(references.items()):
        if i > 0:
            time.sleep(delay)
        try:
            results[label] = fetch_reading(ref)
        except Exception as e:
            print(f"Warning: Could not fetch {label} ({ref}): {e}")
            results[label] = ScriptureReading(
                reference=ref,
                paragraphs=[f"[Reading text not available: {ref}]"],
                poetry_lines=[],
                has_poetry=False,
            )
    return results
