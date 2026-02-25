#!/usr/bin/env python3
"""
Generate a church bulletin for St. Andrew's Episcopal Church, McKinney, TX.

Usage:
    python generate.py 2026-03-01
    python generate.py 2026-03-01 --output output/bulletin.docx
    python generate.py 2026-03-01 --no-prompt  (use defaults for all choices)
"""

import argparse
import sys
from datetime import date, datetime
from pathlib import Path

from bulletin.config import CHURCH_NAME
from bulletin.sources.google_sheet import get_bulletin_data
from bulletin.sources.music_9am import fetch_9am_music
from bulletin.sources.scripture import fetch_readings
from bulletin.sources.songs import lookup_song
from bulletin.sources.parish_prayers import get_ministries_for_date, format_ministries
from bulletin.document.builder import BulletinBuilder


def prompt_choice(question: str, options: list[str]) -> str:
    """Interactive prompt for user choices."""
    print(f"\n{question}")
    if not options:
        # Free-text input
        return input("> ").strip()

    for i, opt in enumerate(options, 1):
        print(f"  {i}. {opt}")

    while True:
        raw = input(f"Choose (1-{len(options)}) [1]: ").strip()
        if not raw:
            return options[0]
        try:
            idx = int(raw)
            if 1 <= idx <= len(options):
                return options[idx - 1]
        except ValueError:
            pass
        print(f"  Please enter a number 1-{len(options)}")


def main():
    parser = argparse.ArgumentParser(
        description=f"Generate a bulletin for {CHURCH_NAME}")
    parser.add_argument("date",
                        help="Target date in YYYY-MM-DD format")
    parser.add_argument("--output", "-o",
                        help="Output .docx file path (default: output/<date>.docx)")
    parser.add_argument("--no-prompt", action="store_true",
                        help="Use defaults for all choices (no interactive prompts)")
    args = parser.parse_args()

    # Parse the target date
    try:
        target_date = datetime.strptime(args.date, "%Y-%m-%d").date()
    except ValueError:
        print(f"Error: Invalid date format '{args.date}'. Use YYYY-MM-DD.")
        sys.exit(1)

    print(f"Generating bulletin for {target_date.strftime('%B %-d, %Y')}...")

    # Step 1: Fetch Google Sheet data
    print("  Fetching liturgical schedule from Google Sheets...")
    try:
        sheet_data = get_bulletin_data(target_date)
    except ValueError as e:
        print(f"Error: {e}")
        sys.exit(1)

    schedule = sheet_data.schedule
    print(f"  Found: {schedule.title} ({schedule.color})")
    if schedule.eucharistic_prayer:
        print(f"  Eucharistic Prayer: {schedule.eucharistic_prayer}")

    # Step 2: Fetch 9am music data
    print("  Fetching 9am music planning data...")
    try:
        music_data = fetch_9am_music(target_date)
        if music_data:
            print(f"  Found {len(music_data.slots)} music slots")
        else:
            print("  Warning: No 9am music data found for this date")
    except Exception as e:
        print(f"  Warning: Could not fetch 9am music: {e}")
        music_data = None

    # Step 3: Fetch scripture readings
    print("  Fetching scripture readings from bible.oremus.org...")
    refs_to_fetch = {}
    if schedule.reading:
        refs_to_fetch["reading"] = schedule.reading
    if schedule.gospel:
        refs_to_fetch["gospel"] = schedule.gospel

    scripture_readings = {}
    if refs_to_fetch:
        try:
            scripture_readings = fetch_readings(refs_to_fetch)
            print(f"  Fetched {len(scripture_readings)} readings")
        except Exception as e:
            print(f"  Warning: Could not fetch scriptures: {e}")

    # Step 4: Parish ministries
    print("  Looking up parish cycle of prayers...")
    try:
        ministries = get_ministries_for_date(target_date)
        parish_ministries = format_ministries(ministries)
        print(f"  Ministries: {parish_ministries}")
    except Exception as e:
        print(f"  Warning: Could not fetch ministries: {e}")
        parish_ministries = "[ministries]"

    # Step 5: Build the bulletin
    print("  Assembling bulletin...")
    prompt_fn = None if args.no_prompt else prompt_choice

    def song_lookup_fn(identifier, service):
        return lookup_song(identifier, service)

    builder = BulletinBuilder(
        target_date=target_date,
        sheet_data=sheet_data,
        music_data=music_data,
        scripture_readings=scripture_readings,
        song_lookup_fn=song_lookup_fn,
        parish_ministries=parish_ministries,
    )

    builder.resolve_all(prompt_fn=prompt_fn)

    doc = builder.build()

    # Step 6: Save
    output_dir = Path("output")
    output_dir.mkdir(exist_ok=True)

    if args.output:
        output_path = Path(args.output)
    else:
        date_str = target_date.strftime("%Y-%m-%d")
        title_slug = schedule.title
        output_path = output_dir / f"{date_str} - {title_slug} - 9 am - Bulletin.docx"

    doc.save(str(output_path))
    print(f"\nBulletin saved to: {output_path}")

    if builder._missing_songs:
        print("\nWarning: Missing song lyrics for:")
        for s in builder._missing_songs:
            print(f"  - {s}")
        print("These will show as [Song lyrics not found] in the bulletin.")


if __name__ == "__main__":
    main()
