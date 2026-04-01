#!/usr/bin/env python3
"""
Generate a church bulletin for St. Andrew's Episcopal Church, McKinney, TX.

Usage:
    python generate.py 2026-03-01                          # all services
    python generate.py 2026-03-01 --service "9 am"         # 9am only
    python generate.py 2026-03-01 --service "11 am"        # 11am only
    python generate.py 2026-03-01 --no-prompt              # use defaults
    python generate.py 2026-03-01 --reading-sheets         # bulletins + reading sheets
    python generate.py 2026-04-02                          # Maundy Thursday → 7 pm
    python generate.py 2026-04-03                          # Good Friday → 7 pm
"""

import argparse
import sys
from datetime import date, datetime
from pathlib import Path

from bulletin.config import CHURCH_NAME, SERVICE_TIMES, get_lectionary_year
from bulletin.logic.rules import get_short_liturgical_title, detect_special_service
from bulletin.sources.google_sheet import get_bulletin_data, get_hidden_springs_data
from bulletin.sources.music_9am import fetch_9am_music
from bulletin.sources.music_11am import get_11am_music_slots
from bulletin.sources.scripture import fetch_readings
from bulletin.sources.songs import lookup_song
from bulletin.sources.parish_prayers import get_ministries_for_date, format_ministries
from bulletin.document.builder import BulletinBuilder
from bulletin.document.reading_sheet import build_reading_sheet


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
    parser.add_argument("--service", "-s",
                        choices=SERVICE_TIMES + ["7 pm", "sunrise",
                                                 "hidden_springs", "all"],
                        default="all",
                        help="Which service to generate (default: all)")
    parser.add_argument("--output", "-o",
                        help="Output .docx file path (only with single service)")
    parser.add_argument("--no-prompt", action="store_true",
                        help="Use defaults for all choices (no interactive prompts)")
    parser.add_argument("--reading-sheets", action="store_true",
                        help="Also generate reading sheets for lay readers")
    parser.add_argument("--force-fetch", action="store_true",
                        help="Re-fetch scripture readings from oremus.org "
                             "(ignore cache)")
    args = parser.parse_args()

    # Parse the target date
    try:
        target_date = datetime.strptime(args.date, "%Y-%m-%d").date()
    except ValueError:
        print(f"Error: Invalid date format '{args.date}'. Use YYYY-MM-DD.")
        sys.exit(1)

    print(f"Generating bulletin for {target_date.strftime('%B %-d, %Y')}...")

    # Determine special service modes
    is_hidden_springs = args.service == "hidden_springs"
    is_sunrise = args.service == "sunrise"

    # Step 1: Fetch data from Google Sheets
    hs_data = None
    if is_hidden_springs:
        # Hidden Springs: fetch HS planner first, then build a synthetic
        # BulletinData from the HS row for compatibility with the builder
        print("  Fetching Hidden Springs planner data...")
        try:
            hs_row, hs_upcoming = get_hidden_springs_data(target_date)
            hs_data = (hs_row, hs_upcoming)
            print(f"  Found: {hs_row.title} ({hs_row.service_type})")
            print(f"  Upcoming services: {len(hs_upcoming)}")
        except ValueError as e:
            print(f"Error: {e}")
            sys.exit(1)

        # Build a synthetic BulletinData from the HS row
        from bulletin.sources.google_sheet import LiturgicalScheduleRow, BulletinData
        synth_schedule = LiturgicalScheduleRow(
            service_type=hs_row.service_type,
            date=hs_row.date,
            title=hs_row.title,
            proper=hs_row.proper,
            color=hs_row.color,
            eucharistic_prayer=hs_row.eucharistic_prayer,
            preface=hs_row.preface,
            reading=hs_row.reading,
            psalm=hs_row.psalm,
            gospel=hs_row.gospel,
            pop_form=hs_row.pop_form,
            special_blessing=hs_row.special_blessing,
            closing_prayer=hs_row.closing_prayer,
            dismissal=hs_row.dismissal,
            notes=hs_row.notes,
        )
        sheet_data = BulletinData(
            schedule=synth_schedule, clergy=None, music=None)
        schedule = synth_schedule
        services = ["hidden_springs"]
    else:
        print("  Fetching liturgical schedule from Google Sheets...")
        try:
            svc_filter = "sunrise" if is_sunrise else None
            sheet_data = get_bulletin_data(target_date,
                                           service_type_filter=svc_filter)
        except ValueError as e:
            print(f"Error: {e}")
            sys.exit(1)
        schedule = sheet_data.schedule

    print(f"  Found: {schedule.title} ({schedule.color})")
    if schedule.eucharistic_prayer:
        print(f"  Eucharistic Prayer: {schedule.eucharistic_prayer}")

    # Detect special weekday services (Maundy Thursday, Good Friday)
    special = detect_special_service(schedule.title)
    is_weekday_special = special in ("maundy_thursday", "good_friday")

    # Determine which services to generate (if not already set for HS)
    if not is_hidden_springs:
        if is_sunrise:
            # Sunrise service uses the 9am pipeline with Service Music data
            services = ["sunrise"]
        elif args.service != "all":
            services = [args.service]
        elif is_weekday_special:
            services = ["7 pm"]
            print(f"  Detected weekday special service: generating 7 pm bulletin")
        else:
            services = list(SERVICE_TIMES)

    # Step 2: Fetch scripture readings
    print("  Fetching scripture readings...")
    refs_to_fetch = {}
    if is_hidden_springs and hs_data:
        hs_row = hs_data[0]
        if hs_row.reading:
            refs_to_fetch["reading"] = hs_row.reading
        if hs_row.gospel:
            refs_to_fetch["gospel"] = hs_row.gospel
    else:
        if schedule.reading:
            refs_to_fetch["reading"] = schedule.reading
        if schedule.gospel:
            refs_to_fetch["gospel"] = schedule.gospel

    # Palm Sunday: also fetch the Palm Gospel (triumphal entry reading)
    if special == "palm_sunday":
        from bulletin.data.loader import load_palm_sunday
        palm_texts = load_palm_sunday()
        liturgy = palm_texts["liturgy_of_the_palms"]
        year_num = target_date.year
        remainder = year_num % 3
        lect_year = "A" if remainder == 1 else ("B" if remainder == 2 else "C")
        palm_gospel_ref = liturgy["palm_gospel"].get(lect_year, "Matthew 21:1-11")
        refs_to_fetch["palm_gospel"] = palm_gospel_ref

    scripture_readings = {}
    if refs_to_fetch:
        try:
            scripture_readings = fetch_readings(
                refs_to_fetch, force_fetch=args.force_fetch)
            print(f"  Fetched {len(scripture_readings)} readings")
        except Exception as e:
            print(f"  Warning: Could not fetch scriptures: {e}")

    # Step 3: Parish ministries (shared across all services)
    print("  Looking up parish cycle of prayers...")
    try:
        ministries = get_ministries_for_date(target_date)
        parish_ministries = format_ministries(ministries)
        print(f"  Ministries: {parish_ministries}")
    except Exception as e:
        print(f"  Warning: Could not fetch ministries: {e}")
        parish_ministries = "[ministries]"

    # Step 4: Fetch service-specific music data
    music_9am = None
    if "9 am" in services:
        print("  Fetching 9am music planning data...")
        try:
            music_9am = fetch_9am_music(target_date)
            if music_9am:
                print(f"  Found {len(music_9am.slots)} music slots")
            else:
                print("  Warning: No 9am music data found for this date")
        except Exception as e:
            print(f"  Warning: Could not fetch 9am music: {e}")

    music_11am_slots = None
    if "11 am" in services or "7 pm" in services or "sunrise" in services:
        # 11am and weekday services use the Service Music tab
        if sheet_data.music:
            music_11am_slots = get_11am_music_slots(sheet_data.music)
            label = "7 pm" if "7 pm" in services else "11am"
            print(f"  Found {len(music_11am_slots)} music slots for {label}")
        else:
            label = "7 pm" if "7 pm" in services else "11am"
            print(f"  Warning: No {label} music data found for this date")

    prompt_fn = None if args.no_prompt else prompt_choice

    def song_lookup_fn(identifier, service):
        return lookup_song(identifier, service)

    # Step 5: Generate each service bulletin
    output_dir = Path("output")
    output_dir.mkdir(exist_ok=True)

    # Shared liturgical resolutions: prompt once, reuse across services
    shared_resolutions = None

    for service_time in services:
        print(f"\n  === Assembling {service_time} bulletin ===")

        # Select the appropriate music data for this service
        if service_time == "hidden_springs":
            music_data = None  # HS music is in the HS planner row
        elif service_time == "sunrise":
            # Sunrise uses Service Music tab but 9am pipeline (full lyrics)
            music_data = music_11am_slots
        elif service_time == "9 am":
            music_data = music_9am
        elif service_time in ("11 am", "7 pm"):
            # Weekday services use the same Service Music tab as 11am
            music_data = music_11am_slots
        else:
            music_data = None

        builder = BulletinBuilder(
            target_date=target_date,
            sheet_data=sheet_data,
            music_data=music_data,
            scripture_readings=scripture_readings,
            song_lookup_fn=song_lookup_fn,
            parish_ministries=parish_ministries,
            service_time=service_time,
            hidden_springs_data=hs_data if service_time == "hidden_springs" else None,
        )

        builder.resolve_all(prompt_fn=prompt_fn,
                            shared_resolutions=shared_resolutions)

        # Capture resolved choices from the first service to share
        if shared_resolutions is None:
            shared_resolutions = builder.get_shared_resolutions()

        doc = builder.build()

        # Determine output path
        if args.output and len(services) == 1:
            output_path = Path(args.output)
        elif service_time == "hidden_springs":
            date_str = target_date.strftime("%Y-%m-%d")
            hs_row = hs_data[0]
            hs_title = hs_row.title or "Hidden Springs"
            svc_type = hs_row.service_type or "LOW"
            output_path = output_dir / f"{date_str} - Hidden Springs - {hs_title} ({svc_type}).docx"
        else:
            date_str = target_date.strftime("%Y-%m-%d")
            short_title = get_short_liturgical_title(schedule.title, schedule.proper)
            year_letter = get_lectionary_year(target_date.year)
            ep_letter = builder.eucharistic_prayer
            # Use original service_time for filename (not remapped)
            file_svc = service_time

            if builder.special_service == "good_friday":
                # Good Friday has no Eucharist → no EP designation
                output_path = output_dir / f"{date_str} - {short_title}{year_letter} - {file_svc} - Bulletin.docx"
            elif builder.special_service:
                # Other special weekday services (Maundy Thursday)
                output_path = output_dir / f"{date_str} - {short_title}{year_letter} - {file_svc} (HEII-{ep_letter}) - Bulletin.docx"
            else:
                output_path = output_dir / f"{date_str} - {short_title}{year_letter} - {file_svc} (HEII-{ep_letter}) - Bulletin.docx"

        # Remove old file first (macOS quarantine attribute workaround)
        if output_path.exists():
            output_path.unlink()

        doc.save(str(output_path))
        print(f"  Saved: {output_path}")

        if builder._missing_songs:
            print(f"  Warning: Missing song lyrics for:")
            for s in builder._missing_songs:
                print(f"    - {s}")

        # Print AAC file manifest for Hidden Springs services
        aac_manifest = builder.get_aac_manifest()
        if aac_manifest:
            print(f"\n  === AAC Files for Upload ===")
            max_slot = max(len(slot) for slot, _ in aac_manifest)
            for slot, filename in aac_manifest:
                print(f"  {slot + ':':<{max_slot + 1}} {filename}")

    # Step 6: Generate reading sheets (if requested)
    if args.reading_sheets and not is_hidden_springs:
        print("\n  === Generating reading sheets ===")

        # Build a reference builder to get the shared reading/POP data.
        # Use "9 am" so the psalm rubric reflects the Google Sheet's mode.
        rs_builder = BulletinBuilder(
            target_date=target_date,
            sheet_data=sheet_data,
            music_data=None,
            scripture_readings=scripture_readings,
            song_lookup_fn=song_lookup_fn,
            parish_ministries=parish_ministries,
            service_time="9 am",
        )
        rs_builder.resolve_all(prompt_fn=prompt_fn,
                               shared_resolutions=shared_resolutions)
        rs_data = rs_builder.get_reading_sheet_data()

        # Determine psalm rubrics for each service group
        rubric_8am = "Read in unison."
        rubric_9_11 = rs_builder.get_psalm_rubric_for_service("9 am")

        date_str = target_date.strftime("%Y-%m-%d")
        short_title = get_short_liturgical_title(schedule.title, schedule.proper)
        year_letter = get_lectionary_year(target_date.year)
        title_tag = f"{short_title}{year_letter}"

        if rubric_8am == rubric_9_11:
            # Same psalm mode for all services → one reading sheet
            doc = build_reading_sheet(rs_data, rubric_8am)
            rs_path = output_dir / f"{date_str} - {title_tag} - Readings and Prayers.docx"
            if rs_path.exists():
                rs_path.unlink()
            doc.save(str(rs_path))
            print(f"  Saved: {rs_path}")
        else:
            # Different psalm modes → two reading sheets
            doc_8 = build_reading_sheet(rs_data, rubric_8am)
            rs_path_8 = output_dir / f"{date_str} - {title_tag} - Readings and Prayers - 8am.docx"
            if rs_path_8.exists():
                rs_path_8.unlink()
            doc_8.save(str(rs_path_8))
            print(f"  Saved: {rs_path_8}")

            doc_9_11 = build_reading_sheet(rs_data, rubric_9_11)
            rs_path_9_11 = output_dir / f"{date_str} - {title_tag} - Readings and Prayers - 9 and 11.docx"
            if rs_path_9_11.exists():
                rs_path_9_11.unlink()
            doc_9_11.save(str(rs_path_9_11))
            print(f"  Saved: {rs_path_9_11}")

    print("\nDone.")


if __name__ == "__main__":
    main()
