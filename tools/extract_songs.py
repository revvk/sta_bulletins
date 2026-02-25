#!/usr/bin/env python3
"""
Extract song lyrics from the Word document and write to YAML files.

This is a one-time extraction tool. The resulting YAML files in
bulletin/data/hymns/ become the source of truth for song lyrics.

Usage:
    python3 tools/extract_songs.py

Reads:
    Bulletin Formatted Song Lyrics - 9 am.docx
    Bulletin Formatted Song Lyrics - 11 am.docx

Writes:
    bulletin/data/hymns/songs_9am.yaml
    bulletin/data/hymns/songs_11am.yaml
"""

import os
import re
import sys
from pathlib import Path

from docx import Document
from lxml import etree
import yaml

# Paths
PROJECT_ROOT = Path(__file__).parent.parent
DATA_DIR = PROJECT_ROOT / "bulletin" / "data" / "hymns"

SONG_FILES = {
    "9am": PROJECT_ROOT / "source_documents" / "Bulletin Formatted Song Lyrics - 9 am.docx",
    "11am": PROJECT_ROOT / "source_documents" / "Bulletin Formatted Song Lyrics - 11 am.docx",
}

NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
}


def extract_songs_from_docx(filepath: Path) -> list[dict]:
    """Extract all songs from a lyrics Word document.

    The document uses text boxes (from Pages conversion), so we parse the
    XML directly. Content is organized alphabetically under heading letters.

    Paragraph styles:
      - 'Heading': alphabetical section divider (A, B, C, ...)
      - 'Body': song title lines + blank separators
      - 'Body - Lyrics': lyric lines

    Songs are separated by: Body-Lyrics -> blank Body -> non-blank Body (title).
    Choruses are identified by italic run formatting.
    """
    doc = Document(str(filepath))
    body = doc.element.body

    # Extract paragraphs from text boxes
    txbx_contents = body.findall('.//wps:txbx/w:txbxContent', NS)
    if not txbx_contents:
        # Fallback: try non-text-box paragraphs
        txbx_contents = [body]

    paragraphs = []
    for txbx in txbx_contents:
        for p_elem in txbx.findall('.//w:p', NS):
            paragraphs.append(_parse_paragraph(p_elem))

    # Now walk paragraphs and extract songs
    songs = []
    current_alpha = ""
    current_song = None
    current_section_lines = []
    current_section_is_chorus = None
    in_lyrics = False

    for para in paragraphs:
        style = para["style"]
        text = para["text"]
        is_italic = para["all_italic"]

        if style == "Heading":
            # Alphabetical section marker
            if text.strip():
                current_alpha = text.strip()
            continue

        if style == "Body":
            # Could be a song title or a blank separator
            if text.strip():
                # Save any in-progress song
                if current_song is not None:
                    _flush_section(current_song, current_section_lines, current_section_is_chorus)
                    songs.append(current_song)

                # Parse the title line
                current_song = _parse_title_line(text, current_alpha)
                current_section_lines = []
                current_section_is_chorus = None
                in_lyrics = False
            # blank Body = separator between songs (or before title)
            continue

        if style == "Body - Lyrics":
            if current_song is None:
                continue

            if not text.strip():
                # Blank lyric line = section break within a song
                if current_section_lines:
                    _flush_section(current_song, current_section_lines, current_section_is_chorus)
                    current_section_lines = []
                    current_section_is_chorus = None
                continue

            # Determine if this is chorus (italic) or verse
            if current_section_is_chorus is None:
                current_section_is_chorus = is_italic
            current_section_lines.append(text)
            in_lyrics = True
            continue

        if style == "Default":
            # Special instructions (rare) -- treat as a note
            if current_song is not None and text.strip():
                if "note" not in current_song:
                    current_song["note"] = text.strip()
            continue

    # Flush the last song
    if current_song is not None:
        _flush_section(current_song, current_section_lines, current_section_is_chorus)
        songs.append(current_song)

    return songs


def _parse_paragraph(p_elem) -> dict:
    """Parse a <w:p> element into a dict with style, text, and italic info."""
    pPr = p_elem.find('w:pPr', NS)
    style = "Normal"
    if pPr is not None:
        pStyle = pPr.find('w:pStyle', NS)
        if pStyle is not None:
            style = pStyle.get(f'{{{NS["w"]}}}val', 'Normal')

    runs = p_elem.findall('w:r', NS)
    text_parts = []
    all_italic = True
    has_text = False

    for r in runs:
        # Check for tab elements within the run (w:tab)
        for child in r:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag == 'tab':
                text_parts.append('\t')

        t_elem = r.find('w:t', NS)
        if t_elem is not None and t_elem.text:
            text_parts.append(t_elem.text)
            has_text = True
            # Check if this run is italic
            rPr = r.find('w:rPr', NS)
            is_italic = False
            if rPr is not None:
                is_italic = rPr.find('w:i', NS) is not None
            if not is_italic:
                all_italic = False

    if not has_text:
        all_italic = False

    return {
        "style": style,
        "text": "".join(text_parts),
        "all_italic": all_italic,
    }


def _parse_title_line(text: str, alpha_section: str) -> dict:
    """Parse a song title line into a song dict.

    Formats:
      "Everlasting God"
      "Ah, holy Jesus, how hast thou offended\t#158 (Hymnal 1982)"
      "Amazing grace!\t\t#671 (Hymnal 1982)"
      "Wait for the Lord [refrain - verse - refrain]\t\tCommunaute de Taize"
    """
    parts = text.split("\t")
    title = parts[0].strip()

    hymnal_number = None
    hymnal_name = None
    tune_name = None

    # Check remaining parts for hymnal reference
    for part in parts[1:]:
        part = part.strip()
        if not part:
            continue

        # Match "#NNN" pattern
        num_match = re.search(r'#(\d+)', part)
        if num_match:
            hymnal_number = num_match.group(1)
            # Check for hymnal name in parentheses
            name_match = re.search(r'\((.+?)\)', part)
            if name_match:
                hymnal_name = name_match.group(1)
            continue

        # Otherwise it might be a tune name or attribution
        if not hymnal_number:
            tune_name = part

    return {
        "title": title,
        "hymnal_number": hymnal_number,
        "hymnal_name": hymnal_name,
        "tune_name": tune_name,
        "alpha_section": alpha_section,
        "sections": [],
    }


def _flush_section(song: dict, lines: list[str], is_chorus: bool | None):
    """Add accumulated lines as a verse or chorus section to the song."""
    if not lines:
        return
    section = {
        "type": "chorus" if is_chorus else "verse",
        "lines": list(lines),
    }
    song["sections"].append(section)


def songs_to_yaml(songs: list[dict]) -> str:
    """Convert song list to YAML string."""
    # Clean up for YAML output
    clean_songs = []
    for song in songs:
        clean = {
            "title": song["title"],
        }
        if song.get("hymnal_number"):
            clean["hymnal_number"] = song["hymnal_number"]
        if song.get("hymnal_name"):
            clean["hymnal_name"] = song["hymnal_name"]
        if song.get("tune_name"):
            clean["tune_name"] = song["tune_name"]
        if song.get("note"):
            clean["note"] = song["note"]
        clean["sections"] = song["sections"]
        clean_songs.append(clean)

    return yaml.dump(
        clean_songs,
        allow_unicode=True,
        default_flow_style=False,
        width=120,
        sort_keys=False,
    )


def main():
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    for label, filepath in SONG_FILES.items():
        if not filepath.exists():
            print(f"Warning: {filepath} not found, skipping")
            continue

        print(f"Extracting songs from {filepath.name}...")
        songs = extract_songs_from_docx(filepath)
        print(f"  Found {len(songs)} songs")

        output_path = DATA_DIR / f"songs_{label}.yaml"
        yaml_text = songs_to_yaml(songs)
        output_path.write_text(yaml_text, encoding="utf-8")
        print(f"  Written to {output_path}")

        # Show a sample
        if songs:
            s = songs[0]
            print(f"  Sample: '{s['title']}' - {len(s['sections'])} sections")


if __name__ == "__main__":
    main()
