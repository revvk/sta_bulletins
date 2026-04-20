"""
Shared bulletin-generation entry point used by both the CLI
(``generate.py``) and the local web UI (``web/app.py``).

Before this module existed, all the orchestration logic — fetching the
schedule, scripture, ministries, music; choosing service times; building
each bulletin; saving the .docx; collecting AAC manifests — lived
inline in ``generate.py``'s ``main()``. Lifting it here lets both
front-ends (argv parsing on the CLI side, an HTML form on the web side)
build the same parameters and consume the same structured result.

The function is deliberately small in surface area:

  - One dataclass-shaped input (``RunOptions``)
  - One dataclass-shaped output (``RunResult``)
  - Optional ``prompt_fn`` for interactive disambiguation (CLI passes
    the input() prompt, web passes None and lets defaults win)
  - Optional ``progress_fn`` for status lines (CLI uses print, web
    captures into an in-memory list)

No new behavior — this is a pure extraction. ``generate.py``'s output
should be byte-identical to before.
"""

from __future__ import annotations

import sys
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Callable, Optional

from bulletin.config import SERVICE_TIMES, get_lectionary_year
from bulletin.logic.rules import detect_special_service, get_short_liturgical_title
from bulletin.report import RunReport
from bulletin.sources.google_sheet import (
    BulletinData,
    LiturgicalScheduleRow,
    get_bulletin_data,
    get_hidden_springs_data,
)
from bulletin.sources.music_9am import fetch_9am_music
from bulletin.sources.music_11am import get_11am_music_slots
from bulletin.sources.parish_prayers import (
    format_ministries,
    get_ministries_for_date,
)
from bulletin.sources.scripture import fetch_readings
from bulletin.sources.songs import lookup_song
from bulletin.document.builder import BulletinBuilder
from bulletin.document.reading_sheet import build_reading_sheet
from bulletin.document.styles import prune_unused_styles


# ---------------------------------------------------------------------------
# Inputs / outputs
# ---------------------------------------------------------------------------

@dataclass
class RunOptions:
    """All knobs that ``generate.py`` exposes via argparse."""
    target_date: date
    service: str = "all"            # "all" | "8 am" | "9 am" | "11 am" | "7 pm" | "sunrise" | "hidden_springs"
    output_dir: Path = field(default_factory=lambda: Path("output"))
    output_path: Optional[Path] = None  # only honored with single service
    reading_sheets: bool = False
    force_fetch: bool = False


@dataclass
class GeneratedBulletin:
    """One generated .docx file plus the per-service AAC manifest (if any)."""
    service_time: str
    output_path: Path
    aac_manifest: list[tuple[str, str]]  # [(slot_name, aac_filename), …]


@dataclass
class RunResult:
    target_date: date
    services_requested: list[str]
    bulletins: list[GeneratedBulletin]
    reading_sheets: list[Path]
    report: RunReport


# Sentinel for "user cancelled" / "fatal error before generating anything".
class RunAborted(Exception):
    """Raised when the run cannot continue (bad date, missing schedule, etc).

    The CLI prints the message and exits non-zero. The web UI surfaces
    it as a flash on the form page.
    """


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def run_generation(
    options: RunOptions,
    *,
    prompt_fn: Optional[Callable] = None,
    progress_fn: Optional[Callable[[str], None]] = None,
    report: Optional[RunReport] = None,
) -> RunResult:
    """Run the full bulletin pipeline and return a structured result.

    Mirrors the logic that used to live in ``generate.py``'s ``main()``.
    Side effects (writing .docx files, printing progress) are preserved
    so CLI behavior is unchanged when called with the defaults
    (``progress_fn=print``).
    """
    if progress_fn is None:
        progress_fn = print

    if report is None:
        report = RunReport()

    target_date = options.target_date
    is_hidden_springs = options.service == "hidden_springs"
    is_sunrise = options.service == "sunrise"

    # ---- Step 1: Fetch sheet data ----
    hs_data = None
    if is_hidden_springs:
        progress_fn("  Fetching Hidden Springs planner data...")
        try:
            hs_row, hs_upcoming = get_hidden_springs_data(target_date)
        except ValueError as e:
            raise RunAborted(str(e)) from e
        hs_data = (hs_row, hs_upcoming)
        progress_fn(f"  Found: {hs_row.title} ({hs_row.service_type})")
        progress_fn(f"  Upcoming services: {len(hs_upcoming)}")

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
        progress_fn("  Fetching liturgical schedule from Google Sheets...")
        try:
            svc_filter = "sunrise" if is_sunrise else None
            sheet_data = get_bulletin_data(
                target_date, service_type_filter=svc_filter)
        except ValueError as e:
            raise RunAborted(str(e)) from e
        schedule = sheet_data.schedule

    progress_fn(f"  Found: {schedule.title} ({schedule.color})")
    if schedule.eucharistic_prayer:
        progress_fn(f"  Eucharistic Prayer: {schedule.eucharistic_prayer}")

    special = detect_special_service(schedule.title)
    is_weekday_special = special in ("maundy_thursday", "good_friday")

    if not is_hidden_springs:
        if is_sunrise:
            services = ["sunrise"]
        elif options.service != "all":
            services = [options.service]
        elif is_weekday_special:
            services = ["7 pm"]
            progress_fn(
                f"  Detected weekday special service: generating 7 pm bulletin")
        else:
            services = list(SERVICE_TIMES)

    # ---- Step 2: Scripture readings ----
    progress_fn("  Fetching scripture readings...")
    refs_to_fetch: dict[str, str] = {}
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

    if special == "palm_sunday":
        from bulletin.data.loader import load_palm_sunday
        palm_texts = load_palm_sunday()
        liturgy = palm_texts["liturgy_of_the_palms"]
        year_num = target_date.year
        remainder = year_num % 3
        lect_year = "A" if remainder == 1 else ("B" if remainder == 2 else "C")
        palm_gospel_ref = liturgy["palm_gospel"].get(lect_year, "Matthew 21:1-11")
        refs_to_fetch["palm_gospel"] = palm_gospel_ref

    scripture_readings: dict[str, dict] = {}
    if refs_to_fetch:
        try:
            scripture_readings = fetch_readings(
                refs_to_fetch, force_fetch=options.force_fetch, report=report)
            progress_fn(f"  Fetched {len(scripture_readings)} readings")
        except Exception as e:
            progress_fn(f"  Warning: Could not fetch scriptures: {e}")
            report.blocker(
                category="scripture",
                message=f"Could not fetch any scripture readings: {e}",
                fix_hint="Check your network connection and try again "
                         "with --force-fetch.")

    # ---- Step 3: Parish ministries ----
    progress_fn("  Looking up parish cycle of prayers...")
    try:
        ministries = get_ministries_for_date(target_date)
        parish_ministries = format_ministries(ministries)
        progress_fn(f"  Ministries: {parish_ministries}")
    except Exception as e:
        progress_fn(f"  Warning: Could not fetch ministries: {e}")
        report.warning(
            category="ministry",
            message=f"Could not fetch parish ministry rotation: {e}",
            fix_hint="Bulletin uses '[ministries]' placeholder in the "
                     "Prayers of the People — fill it in by hand in the "
                     "saved .docx.")
        parish_ministries = "[ministries]"

    # ---- Step 4: Service-specific music ----
    music_9am = None
    if "9 am" in services:
        progress_fn("  Fetching 9am music planning data...")
        try:
            music_9am = fetch_9am_music(target_date)
            if music_9am:
                progress_fn(f"  Found {len(music_9am.slots)} music slots")
            else:
                progress_fn(
                    "  Warning: No 9am music data found for this date")
                report.warning(
                    category="music",
                    message="No 9 am music data found for this date",
                    fix_hint="Check the 9am Music Planning Sheet — the row "
                             "for this date may be missing or the date "
                             "header may be misformatted.")
        except Exception as e:
            progress_fn(f"  Warning: Could not fetch 9am music: {e}")
            report.warning(
                category="music",
                message=f"Could not fetch 9 am music: {e}",
                fix_hint="9 am bulletin will have empty song slots; check "
                         "your network and the 9am Music Planning Sheet.")

    music_11am_slots = None
    if "11 am" in services or "7 pm" in services or "sunrise" in services:
        if sheet_data.music:
            music_11am_slots = get_11am_music_slots(sheet_data.music)
            label = "7 pm" if "7 pm" in services else "11am"
            progress_fn(
                f"  Found {len(music_11am_slots)} music slots for {label}")
        else:
            label = "7 pm" if "7 pm" in services else "11am"
            progress_fn(f"  Warning: No {label} music data found for this date")
            report.warning(
                category="music",
                message=f"No {label} music data found for this date",
                fix_hint="Check the Service Music tab — the row for this "
                         "date may be missing.")

    def song_lookup_fn(identifier: str, service: str):
        return lookup_song(identifier, service)

    # ---- Step 5: Build each bulletin ----
    output_dir = options.output_dir
    output_dir.mkdir(exist_ok=True, parents=True)

    bulletins: list[GeneratedBulletin] = []
    shared_resolutions = None

    for service_time in services:
        progress_fn(f"\n  === Assembling {service_time} bulletin ===")

        if service_time == "hidden_springs":
            music_data = None
        elif service_time == "sunrise":
            music_data = music_11am_slots
        elif service_time == "9 am":
            music_data = music_9am
        elif service_time in ("11 am", "7 pm"):
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
            report=report,
        )

        builder.resolve_all(prompt_fn=prompt_fn,
                            shared_resolutions=shared_resolutions)
        if shared_resolutions is None:
            shared_resolutions = builder.get_shared_resolutions()

        doc = builder.build()

        # Filename
        if options.output_path and len(services) == 1:
            output_path = options.output_path
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
            file_svc = service_time

            if builder.special_service == "good_friday":
                output_path = output_dir / f"{date_str} - {short_title}{year_letter} - {file_svc} - Bulletin.docx"
            else:
                # Maundy Thursday and standard Sunday services
                output_path = output_dir / f"{date_str} - {short_title}{year_letter} - {file_svc} (HEII-{ep_letter}) - Bulletin.docx"

        if output_path.exists():
            output_path.unlink()
        prune_unused_styles(doc)
        doc.save(str(output_path))
        progress_fn(f"  Saved: {output_path}")

        if builder._missing_songs:
            progress_fn(f"  Warning: Missing song lyrics for:")
            for s in builder._missing_songs:
                progress_fn(f"    - {s}")

        aac_manifest = builder.get_aac_manifest()
        if aac_manifest:
            progress_fn(f"\n  === AAC Files for Upload ===")
            max_slot = max(len(slot) for slot, _ in aac_manifest)
            for slot, filename in aac_manifest:
                progress_fn(f"  {slot + ':':<{max_slot + 1}} {filename}")

        bulletins.append(GeneratedBulletin(
            service_time=service_time,
            output_path=output_path,
            aac_manifest=list(aac_manifest),
        ))

    # ---- Step 6: Reading sheets ----
    reading_sheet_paths: list[Path] = []
    if options.reading_sheets and not is_hidden_springs:
        progress_fn("\n  === Generating reading sheets ===")

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

        rubric_8am = "Read in unison."
        rubric_9_11 = rs_builder.get_psalm_rubric_for_service("9 am")

        date_str = target_date.strftime("%Y-%m-%d")
        short_title = get_short_liturgical_title(schedule.title, schedule.proper)
        year_letter = get_lectionary_year(target_date.year)
        title_tag = f"{short_title}{year_letter}"

        if rubric_8am == rubric_9_11:
            doc = build_reading_sheet(rs_data, rubric_8am)
            rs_path = output_dir / f"{date_str} - {title_tag} - Readings and Prayers.docx"
            if rs_path.exists():
                rs_path.unlink()
            prune_unused_styles(doc)
            doc.save(str(rs_path))
            progress_fn(f"  Saved: {rs_path}")
            reading_sheet_paths.append(rs_path)
        else:
            doc_8 = build_reading_sheet(rs_data, rubric_8am)
            rs_path_8 = output_dir / f"{date_str} - {title_tag} - Readings and Prayers - 8am.docx"
            if rs_path_8.exists():
                rs_path_8.unlink()
            prune_unused_styles(doc_8)
            doc_8.save(str(rs_path_8))
            progress_fn(f"  Saved: {rs_path_8}")
            reading_sheet_paths.append(rs_path_8)

            doc_9_11 = build_reading_sheet(rs_data, rubric_9_11)
            rs_path_9_11 = output_dir / f"{date_str} - {title_tag} - Readings and Prayers - 9 and 11.docx"
            if rs_path_9_11.exists():
                rs_path_9_11.unlink()
            prune_unused_styles(doc_9_11)
            doc_9_11.save(str(rs_path_9_11))
            progress_fn(f"  Saved: {rs_path_9_11}")
            reading_sheet_paths.append(rs_path_9_11)

    return RunResult(
        target_date=target_date,
        services_requested=services,
        bulletins=bulletins,
        reading_sheets=reading_sheet_paths,
        report=report,
    )
