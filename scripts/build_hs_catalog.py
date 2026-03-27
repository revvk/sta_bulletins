#!/usr/bin/env python3
"""One-time script to build hidden_springs_songs.yaml from AAC files + lyrics sources.

Scans AAC files, cross-references songs.yaml and the Google Sheet lyrics CSV,
and produces a complete catalog with lyrics populated where available.

Usage:
    python scripts/build_hs_catalog.py
"""

import csv
import io
import os
import re
import urllib.request
from pathlib import Path

import yaml

PROJECT_ROOT = Path(__file__).resolve().parent.parent
AAC_BASE = PROJECT_ROOT / "hidden_springs_music" / "sorted_hidden_springs_aac_files"
SONGS_YAML = PROJECT_ROOT / "bulletin" / "data" / "hymns" / "songs.yaml"
OUTPUT_YAML = PROJECT_ROOT / "bulletin" / "data" / "hymns" / "hidden_springs_songs.yaml"
LYRICS_SHEET_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1vX9llfMg0bAWZiSM10RaAsgExyKIN5YflKaaaEPoOOI/gviz/tq?tqx=out:csv"
)

# Songs that don't need printed lyrics
NO_LYRICS_TITLES = {
    "chimes", "happy birthday", "god bless america",
}
NO_LYRICS_HYMNAL = {"S280", "S130"}

CATEGORY_MAP = {
    "Hymnal 1982": "hymnal",
    "Other": "other",
    "Prelude - Postlude": "prelude_postlude",
}


def parse_aac_filename(filename: str, category: str):
    """Parse an AAC filename into structured metadata."""
    name = filename.rsplit(".", 1)[0]  # strip extension

    hymnal_num = None
    title = name

    # Extract hymnal number prefix: "362 - Title" or "S280 - Title"
    m = re.match(r"^(\d+|S\d+)\s*-\s*(.+)", name)
    if m:
        hymnal_num = m.group(1)
        title = m.group(2).strip()

    # Extract [N verses]
    verses = None
    vm = re.search(r"\[(\d+)\s*verses?\]", title)
    if vm:
        verses = int(vm.group(1))
        title = re.sub(r"\s*\[\d+\s*verses?\]", "", title).strip()

    # Extract [ORGAN], [PIANO], [PRELUDE], [POSTLUDE], [PRELUDE-POSTLUDE]
    instrument = None
    usage = None
    im = re.search(r"\[([A-Z-]+)\]", title)
    if im:
        tag = im.group(1)
        if tag in ("ORGAN", "PIANO"):
            instrument = tag.lower()
        elif tag in ("PRELUDE", "POSTLUDE", "PRELUDE-POSTLUDE"):
            usage = tag.lower()
        title = re.sub(r"\s*-?\s*\[([A-Z-]+)\]", "", title).strip()

    aac_entry = {"filename": f"{category}/{filename}"}
    if verses:
        aac_entry["verses"] = verses
    if instrument:
        aac_entry["instrument"] = instrument
    if usage:
        aac_entry["usage"] = usage

    return {
        "hymnal_num": hymnal_num,
        "title": title,
        "aac_entry": aac_entry,
        "category_key": CATEGORY_MAP[category],
    }


def scan_aac_files():
    """Scan AAC directory and group by song."""
    songs = {}  # key: (hymnal_num or None, cleaned_title) -> song dict

    for category in ["Hymnal 1982", "Other", "Prelude - Postlude"]:
        cat_path = AAC_BASE / category
        if not cat_path.is_dir():
            continue
        for f in sorted(cat_path.iterdir()):
            if f.suffix not in (".m4a", ".aac"):
                continue
            parsed = parse_aac_filename(f.name, category)
            key = (parsed["hymnal_num"], parsed["title"])

            if key not in songs:
                is_no_lyrics = (
                    parsed["category_key"] == "prelude_postlude"
                    or parsed["title"].lower() in NO_LYRICS_TITLES
                    or (parsed["hymnal_num"] or "") in NO_LYRICS_HYMNAL
                )
                songs[key] = {
                    "title": parsed["title"],
                    "hymnal_num": parsed["hymnal_num"],
                    "category": parsed["category_key"],
                    "no_lyrics": is_no_lyrics,
                    "aac_files": [],
                    "sections": [],
                }
            songs[key]["aac_files"].append(parsed["aac_entry"])

    return songs


def load_songs_yaml():
    """Load existing songs.yaml and index by hymnal number and title."""
    with open(SONGS_YAML) as f:
        songs = yaml.safe_load(f)

    by_hymnal = {}
    by_title = {}
    for s in songs:
        sections = s.get("sections", [])
        if not sections:
            continue
        hnum = s.get("hymnal_number", "")
        title = s.get("title", "")
        if hnum:
            by_hymnal[hnum] = sections
        by_title[title.lower()] = sections
        for ident in s.get("identifiers", []):
            by_title[ident.lower()] = sections

    return by_hymnal, by_title


def fetch_lyrics_sheet():
    """Fetch the Hidden Springs lyrics Google Sheet as CSV."""
    req = urllib.request.Request(LYRICS_SHEET_URL)
    with urllib.request.urlopen(req) as resp:
        data = resp.read().decode("utf-8")

    by_hymnal = {}
    by_title = {}
    reader = csv.reader(io.StringIO(data))
    next(reader)  # skip header

    for row in reader:
        if len(row) < 2 or not row[0].strip():
            continue
        title = row[0].strip()
        lyrics_raw = row[1].strip() if len(row) > 1 else ""
        hymn_num = row[3].strip() if len(row) > 3 else ""

        if len(lyrics_raw) < 20:
            continue

        sections = parse_plain_lyrics(lyrics_raw)
        if not sections:
            continue

        if hymn_num:
            by_hymnal[hymn_num] = sections
        by_title[title.lower()] = sections

    return by_hymnal, by_title


def parse_plain_lyrics(text: str) -> list[dict]:
    """Convert plain text lyrics (double-newline separated verses) to sections format."""
    # Split on double newlines (verse separators)
    blocks = re.split(r"\n\s*\n", text.strip())
    sections = []
    for block in blocks:
        lines = [line.strip() for line in block.strip().split("\n") if line.strip()]
        if lines:
            sections.append({"type": "verse", "lines": lines})
    return sections


def match_lyrics(song, yaml_by_hymnal, yaml_by_title, sheet_by_hymnal, sheet_by_title):
    """Try to find lyrics for a song from available sources. Returns sections or []."""
    hnum = song["hymnal_num"]
    title = song["title"]
    title_lower = title.lower()

    # Priority 1: songs.yaml by hymnal number
    if hnum and hnum in yaml_by_hymnal:
        return yaml_by_hymnal[hnum]

    # Priority 2: songs.yaml by title (exact)
    if title_lower in yaml_by_title:
        return yaml_by_title[title_lower]

    # Priority 2b: songs.yaml by prefix/substring
    for yt, sections in yaml_by_title.items():
        if title_lower.startswith(yt) or yt.startswith(title_lower):
            return sections

    # Priority 3: Google Sheet by hymnal number
    if hnum and hnum in sheet_by_hymnal:
        return sheet_by_hymnal[hnum]

    # Priority 4: Google Sheet by title (fuzzy)
    if title_lower in sheet_by_title:
        return sheet_by_title[title_lower]

    # Try prefix/substring matching on sheet titles
    for sheet_title, sections in sheet_by_title.items():
        # Strip parenthetical notes from sheet titles for matching
        clean_sheet = re.sub(r"\s*\(.*?\)\s*$", "", sheet_title).strip()
        if title_lower.startswith(clean_sheet) or clean_sheet.startswith(title_lower):
            return sections
        if title_lower in clean_sheet or clean_sheet in title_lower:
            return sections

    return []


def build_catalog():
    """Main: build the hidden_springs_songs.yaml catalog."""
    print("Scanning AAC files...")
    songs = scan_aac_files()
    print(f"  Found {len(songs)} unique songs/pieces")

    print("Loading songs.yaml...")
    yaml_by_hymnal, yaml_by_title = load_songs_yaml()
    print(f"  {len(yaml_by_hymnal)} hymnal entries, {len(yaml_by_title)} title entries")

    print("Fetching Google Sheet lyrics...")
    try:
        sheet_by_hymnal, sheet_by_title = fetch_lyrics_sheet()
        print(f"  {len(sheet_by_hymnal)} hymnal entries, {len(sheet_by_title)} title entries")
    except Exception as e:
        print(f"  Warning: Could not fetch sheet: {e}")
        sheet_by_hymnal, sheet_by_title = {}, {}

    # Match lyrics
    matched = 0
    no_lyrics_count = 0
    missing = []

    for key, song in songs.items():
        if song["no_lyrics"]:
            no_lyrics_count += 1
            continue

        sections = match_lyrics(
            song, yaml_by_hymnal, yaml_by_title, sheet_by_hymnal, sheet_by_title
        )
        if sections:
            song["sections"] = sections
            matched += 1
        else:
            missing.append(song["title"])

    print(f"\nResults:")
    print(f"  Lyrics matched: {matched}")
    print(f"  No lyrics needed: {no_lyrics_count}")
    print(f"  Missing lyrics: {len(missing)}")
    if missing:
        print(f"  Missing songs:")
        for t in sorted(missing):
            print(f"    - {t}")

    # Build output YAML
    output = []
    # Sort: hymnal first (by number), then other (by title), then prelude_postlude (by title)
    sorted_songs = sorted(
        songs.values(),
        key=lambda s: (
            {"hymnal": 0, "other": 1, "prelude_postlude": 2}[s["category"]],
            int(re.sub(r"[^\d]", "", s["hymnal_num"] or "9999") or "9999"),
            s["title"],
        ),
    )

    for song in sorted_songs:
        entry = {"title": song["title"]}
        if song["hymnal_num"]:
            entry["hymnal_number"] = song["hymnal_num"]
        entry["category"] = song["category"]
        if song["no_lyrics"]:
            entry["no_lyrics"] = True
        entry["aac_files"] = song["aac_files"]
        if song["sections"]:
            entry["sections"] = song["sections"]

        output.append(entry)

    # Write YAML
    with open(OUTPUT_YAML, "w") as f:
        yaml.dump(output, f, default_flow_style=False, allow_unicode=True,
                  sort_keys=False, width=120)

    print(f"\nWritten to: {OUTPUT_YAML}")
    print(f"Total entries: {len(output)}")


if __name__ == "__main__":
    build_catalog()
