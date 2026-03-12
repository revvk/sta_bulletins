"""
Fetch scripture readings with verse numbers from bible.oremus.org.

The Google Sheet provides scripture references (e.g., "Genesis 12:1-4a").
This module fetches the full NRSV text with inline verse numbers.

Oremus HTML is somewhat malformed (nested <p> tags, inline <h2> headings),
so we take a stream-based approach: walk the entire bibletext div's
descendants and collect text, detecting verse numbers and paragraph breaks.
"""

import json
import re
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import requests
from bs4 import BeautifulSoup, NavigableString, Comment, Tag

from bulletin.config import OREMUS_BASE_URL, OREMUS_PARAMS


# ---------------------------------------------------------------------------
# Scripture cache — stores fetched readings to avoid repeated HTTP requests.
# Over a three-year lectionary cycle, this builds a complete library of all
# readings used in the bulletin.
# ---------------------------------------------------------------------------

SCRIPTURE_CACHE_FILE = Path(__file__).parent.parent / "data" / "scripture_cache.json"


def _load_cache() -> dict:
    """Load the scripture cache from disk."""
    if SCRIPTURE_CACHE_FILE.exists():
        try:
            with open(SCRIPTURE_CACHE_FILE, encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            return {}
    return {}


def _save_cache(cache: dict):
    """Write the scripture cache to disk."""
    with open(SCRIPTURE_CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2, sort_keys=True)


def _reading_to_cache(reading: "ScriptureReading") -> dict:
    """Serialize a ScriptureReading for JSON storage."""
    return {
        "paragraphs": reading.paragraphs,
        "poetry_lines": reading.poetry_lines,
        "has_poetry": reading.has_poetry,
    }


def _reading_from_cache(reference: str, data: dict) -> "ScriptureReading":
    """Reconstruct a ScriptureReading from cached JSON data."""
    return ScriptureReading(
        reference=reference,
        paragraphs=data["paragraphs"],
        poetry_lines=data.get("poetry_lines", []),
        has_poetry=data.get("has_poetry", False),
    )


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

            # <br/> = line break, but only inside blockquotes (poetry).
            # Prose <br/> tags are just visual wrapping and should be ignored.
            if desc.name == "br":
                if desc.find_parent("blockquote"):
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
            # Mark verse numbers with \x01 delimiters so the formatter
            # can reliably detect them without heuristic regex.
            current.append(f"\x01{tval}\x01 ")
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

    # Americanize British spellings (NRSV from Oremus uses British English)
    cleaned = [_americanize_text(p) for p in cleaned]
    poetry_cleaned = [
        _americanize_text(re.sub(r'  +', ' ', l).strip())
        for l in poetry_lines if l.strip()
    ]

    return ScriptureReading(
        reference=reference,
        paragraphs=cleaned,
        poetry_lines=poetry_cleaned,
        has_poetry=has_poetry,
    )


# ---------------------------------------------------------------------------
# British → American spelling conversion
# ---------------------------------------------------------------------------

# Word-level replacements: British → American (lowercase).
# Applied case-sensitively by _americanize_text().
_BRITISH_TO_AMERICAN = {
    # -our → -or
    "saviour": "savior",
    "favour": "favor",
    "favoured": "favored",
    "favourable": "favorable",
    "honour": "honor",
    "honoured": "honored",
    "honourable": "honorable",
    "honour's": "honor's",
    "neighbour": "neighbor",
    "neighbours": "neighbors",
    "neighbour's": "neighbor's",
    "colour": "color",
    "coloured": "colored",
    "colours": "colors",
    "labour": "labor",
    "laboured": "labored",
    "labours": "labors",
    "behaviour": "behavior",
    "endeavour": "endeavor",
    "endeavoured": "endeavored",
    "splendour": "splendor",
    "vapour": "vapor",
    "vapours": "vapors",
    "valour": "valor",
    "armour": "armor",
    "tumour": "tumor",
    "rigour": "rigor",
    "fervour": "fervor",
    "clamour": "clamor",
    "rancour": "rancor",
    "odour": "odor",
    # -ise → -ize
    "baptise": "baptize",
    "baptised": "baptized",
    "recognise": "recognize",
    "recognised": "recognized",
    "realise": "realize",
    "realised": "realized",
    "organise": "organize",
    "organised": "organized",
    "apologise": "apologize",
    "criticise": "criticize",
    "emphasise": "emphasize",
    "symbolise": "symbolize",
    "authorise": "authorize",
    "authorised": "authorized",
    # -ence/-ense
    "defence": "defense",
    "offence": "offense",
    "offences": "offenses",
    "licence": "license",
    # -re → -er
    "centre": "center",
    "centres": "centers",
    "metre": "meter",
    "metres": "meters",
    "theatre": "theater",
    "sombre": "somber",
    # Other common differences
    "judgement": "judgment",
    "judgements": "judgments",
    "fulfil": "fulfill",
    "fulfilled": "fulfilled",
    "fulfilment": "fulfillment",
    "enrol": "enroll",
    "enrolled": "enrolled",
    "skilful": "skillful",
    "wilful": "willful",
    "plough": "plow",
    "ploughs": "plows",
    "ploughed": "plowed",
    "ploughman": "plowman",
    "ploughmen": "plowmen",
    "amongst": "among",
    "whilst": "while",
    "towards": "toward",
    "counsellor": "counselor",
    "counsellors": "counselors",
    "traveller": "traveler",
    "travellers": "travelers",
    "marvelled": "marveled",
    "marvellous": "marvelous",
    "jewellery": "jewelry",
    "grey": "gray",
    "draught": "draft",
}

# Build a regex that matches any British word (case-insensitive, whole words).
# Sorted longest-first so longer forms match before shorter ones.
_BRIT_PATTERN = re.compile(
    r"\b(" +
    "|".join(re.escape(w) for w in sorted(_BRITISH_TO_AMERICAN, key=len, reverse=True)) +
    r")\b",
    re.IGNORECASE,
)


def _americanize_text(text: str) -> str:
    """Replace common British spellings with American equivalents.

    Preserves original capitalization: if the British word was capitalized
    (e.g., 'Saviour'), the replacement will be too ('Savior').
    """
    def _replace(match):
        word = match.group(0)
        replacement = _BRITISH_TO_AMERICAN[word.lower()]
        # Preserve capitalization
        if word[0].isupper():
            replacement = replacement[0].upper() + replacement[1:]
        if word.isupper():
            replacement = replacement.upper()
        return replacement

    return _BRIT_PATTERN.sub(_replace, text)


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
                   delay: float = 0.5,
                   force_fetch: bool = False) -> dict[str, ScriptureReading]:
    """Fetch multiple readings, using a local cache when available.

    On first fetch, readings are saved to scripture_cache.json. Subsequent
    runs load cached text instantly — no network request needed. Over a
    three-year lectionary cycle this builds a complete offline library.

    Args:
        references: Dict mapping label to reference,
                    e.g., {"reading": "Genesis 12:1-4a", "gospel": "John 3:1-17"}
        delay: Seconds to wait between requests (be nice to oremus.org)
        force_fetch: If True, bypass the cache and re-fetch from oremus.org.

    Returns:
        Dict mapping label to ScriptureReading.
    """
    cache = _load_cache()
    results = {}
    fetched_new = False

    for label, ref in references.items():
        cache_key = ref.strip()

        # Use cache if available (and not forcing a refresh)
        if not force_fetch and cache_key in cache:
            results[label] = _reading_from_cache(ref, cache[cache_key])
            print(f"    {label}: {ref} (cached)")
            continue

        # Fetch from oremus.org
        if fetched_new:
            time.sleep(delay)
        try:
            reading = fetch_reading(ref)
            results[label] = reading
            cache[cache_key] = _reading_to_cache(reading)
            fetched_new = True
        except Exception as e:
            print(f"Warning: Could not fetch {label} ({ref}): {e}")
            results[label] = ScriptureReading(
                reference=ref,
                paragraphs=[f"[Reading text not available: {ref}]"],
                poetry_lines=[],
                has_poetry=False,
            )

    # Persist any newly fetched readings
    if fetched_new:
        _save_cache(cache)

    return results
