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


if __name__ == "__main__":
    main()
