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

This file is the *CLI front-end*. The actual orchestration lives in
``bulletin.runner.run_generation``, which the local web UI also calls.
"""

import argparse
import sys
from datetime import datetime
from pathlib import Path

from bulletin.config import CHURCH_NAME, SERVICE_TIMES
from bulletin.runner import RunAborted, RunOptions, run_generation


def prompt_choice(question: str, options: list[str]) -> str:
    """Interactive prompt for user choices."""
    print(f"\n{question}")
    if not options:
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
    parser.add_argument("date", nargs="?",
                        help="Target date in YYYY-MM-DD format "
                             "(omit when using --funeral)")
    parser.add_argument("--service", "-s",
                        choices=SERVICE_TIMES + ["7 pm", "sunrise",
                                                 "hidden_springs", "all"],
                        default="all",
                        help="Which service to generate (default: all)")
    parser.add_argument("--funeral",
                        help="Generate a funeral / memorial bulletin from "
                             "a per-service YAML. Pass either a slug "
                             "('2026-01-31-cox') or a path to the YAML.")
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

    # ------------------------------------------------------------------
    # Funeral / memorial branch — entirely separate from the Sunday flow.
    # ------------------------------------------------------------------
    if args.funeral:
        _run_funeral(args.funeral, output_dir=Path("output"),
                     output_path=Path(args.output) if args.output else None)
        return

    if not args.date:
        parser.error("date is required (or pass --funeral <slug>)")

    try:
        target_date = datetime.strptime(args.date, "%Y-%m-%d").date()
    except ValueError:
        print(f"Error: Invalid date format '{args.date}'. Use YYYY-MM-DD.")
        sys.exit(1)

    print(f"Generating bulletin for {target_date.strftime('%B %-d, %Y')}...")

    options = RunOptions(
        target_date=target_date,
        service=args.service,
        output_dir=Path("output"),
        output_path=Path(args.output) if args.output else None,
        reading_sheets=args.reading_sheets,
        force_fetch=args.force_fetch,
    )
    prompt_fn = None if args.no_prompt else prompt_choice

    try:
        result = run_generation(options, prompt_fn=prompt_fn)
    except RunAborted as e:
        print(f"Error: {e}")
        sys.exit(1)

    # Unified post-generation TODO report. Prints nothing if the run had
    # no warnings/blockers/manual items, so clean runs stay quiet.
    result.report.print_console()

    print("\nDone.")


def _run_funeral(slug_or_path: str, *, output_dir: Path,
                  output_path: Path | None) -> None:
    """Generate a funeral / memorial bulletin from a per-service YAML.

    Loads the YAML, fetches scripture for every reading reference it
    names, then runs ``FuneralBuilder``. Output filename mirrors the
    convention used by St. Andrew's existing Pages bulletins,
    e.g. ``2026-01-31 - Burial of the Dead - Annette Cox.docx``.
    """
    from bulletin.sources.funeral_data import load_service
    from bulletin.sources.scripture import fetch_reading
    from bulletin.sources.songs import lookup_song
    from bulletin.sources.music_11am import parse_11am_identifier
    from bulletin.document.funeral_builder import FuneralBuilder

    fd = load_service(slug_or_path)
    print(f"Generating funeral bulletin: {fd.slug}")
    print(f"  {fd.cover_subtitle_resolved} — Rite {fd.rite}, "
          f"HC={fd.holy_eucharist['enabled']}, "
          f"commendation={fd.include_commendation}, "
          f"committal={fd.include_committal}")

    # ----- Scripture --------------------------------------------------
    # Psalms come from the BCP Coverdale psalter (bulletin/data/bcp_texts/
    # psalms.yaml) so they render with the same hanging-indent verse
    # layout as Sunday bulletins. Other readings come from oremus.
    from bulletin.sources.psalms import get_psalm
    scripture = {}
    psalm_ref = fd.readings.get("psalm")
    if psalm_ref:
        try:
            scripture[psalm_ref] = get_psalm(psalm_ref).to_lines()
        except Exception as e:
            print(f"  Warning: could not look up {psalm_ref}: {e}")
    for ref in (fd.readings.get("first"),
                 fd.readings.get("second"),
                 fd.readings.get("gospel")):
        if not ref:
            continue
        try:
            scripture[ref] = fetch_reading(ref)
        except Exception as e:
            print(f"  Warning: could not fetch {ref}: {e}")

    # ----- Song lookup -----------------------------------------------
    # Funerals draw from the 11am music pool (full hymnals + songbook).
    # Wrap lookup_song with a hymnal-stub fallback: if the title
    # carries a "#NNN" prefix and the catalog has no full lyrics, we
    # synthesize a header-only stub so the bulletin still prints the
    # title + hymnal reference (matches the 11am Sunday pipeline's
    # _lookup_slot behavior).
    def song_lookup_fn(title: str, service: str):
        result = lookup_song(title, service="11am")
        if result and result.get("sections"):
            return result
        parsed = parse_11am_identifier(title)
        if parsed.get("hymnal_number"):
            return {
                "title":         result["title"] if result else (parsed.get("title") or title),
                "hymnal_number": parsed["hymnal_number"],
                "hymnal_name":   parsed.get("hymnal_name") or "Hymnal 1982",
                "tune_name":     (result or {}).get("tune_name"),
                "sections":      [],   # header-only render
            }
        return result   # may be None — renderer prints "[Song lyrics not found]"

    builder = FuneralBuilder(fd, scripture, song_lookup_fn)
    doc = builder.build()

    out = output_path or (output_dir / builder.output_filename())
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    print(f"\nWrote: {out}")


if __name__ == "__main__":
    main()
