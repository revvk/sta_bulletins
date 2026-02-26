#!/usr/bin/env python3
"""
One-time extraction script: scrape Collects of the Day from lectionarypage.net.

The BCP collects are fixed text — the same for a given Sunday regardless of
lectionary year (A, B, C).  We scrape the 2026 liturgical calendar page to
get the full list of Sunday / principal feast URLs, fetch each one, and
extract the collect text.

Output: bulletin/data/bcp_texts/collects.yaml

Usage:
    python tools/extract_collects.py
"""

import re
import sys
import time
from pathlib import Path

import requests
import yaml
from bs4 import BeautifulSoup

BASE_URL = "https://www.lectionarypage.net"
CALENDAR_URL = f"{BASE_URL}/CalndrsIndexes/Calendar2026.html"
OUTPUT_PATH = Path(__file__).parent.parent / "bulletin" / "data" / "bcp_texts" / "collects.yaml"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}

# Delay between requests to be polite
REQUEST_DELAY = 1.0


def fetch_calendar_links() -> list[tuple[str, str]]:
    """Fetch all Sunday/feast page links from the 2026 calendar."""
    r = requests.get(CALENDAR_URL, headers=HEADERS, timeout=30)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")

    links = []
    seen = set()
    for a in soup.find_all("a", href=True):
        href = a["href"]
        text = a.get_text(strip=True)
        if not text or "Year" not in href:
            continue
        # Normalize relative paths
        if href.startswith("../"):
            href = href[3:]
        # Skip weekdays of Lent, minor saints, special services
        if any(skip in href for skip in (
            "WeekdaysOfLent", "HolyDays", "SpecServ", "LFF",
        )):
            continue
        if href not in seen:
            seen.add(href)
            links.append((href, text))

    return links


def extract_collect(url: str) -> str | None:
    """Fetch a lectionary page and extract the Collect text."""
    r = requests.get(url, headers=HEADERS, timeout=30)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")

    # Strategy: find an <h2> or <b> element containing "The Collect" or "Collect"
    # then take the next <p> sibling as the collect text.
    for tag in soup.find_all(["h2", "h3", "b", "strong"]):
        text = tag.get_text(strip=True)
        if re.match(r"^The Collect$", text, re.IGNORECASE):
            nxt = tag.find_next_sibling()
            while nxt and nxt.name != "p":
                nxt = nxt.find_next_sibling()
            if nxt:
                collect = nxt.get_text(strip=True)
                # Some pages have "or" followed by an alternate collect —
                # just take the first one
                if collect and len(collect) > 20:
                    return collect
            break

    # Fallback: search for <p> containing text that looks like a collect
    # (starts with "O God" or "Almighty God" or similar, ends with "Amen.")
    for p in soup.find_all("p"):
        text = p.get_text(strip=True)
        if text.endswith("Amen.") and len(text) > 80:
            # Likely a collect
            if any(text.startswith(opener) for opener in (
                "Almighty", "O God", "O Lord", "Grant", "Most",
                "Heavenly", "Lord God", "Eternal", "Merciful",
                "Almighty and everliving", "Almighty and everlasting",
            )):
                return text

    return None


def extract_page_title(url: str, soup: BeautifulSoup = None) -> str | None:
    """Try to extract the liturgical day title from the page itself."""
    if soup is None:
        r = requests.get(url, headers=HEADERS, timeout=30)
        soup = BeautifulSoup(r.text, "html.parser")

    # The title is usually in the first <h2> or the <title> tag
    title_tag = soup.find("title")
    if title_tag:
        t = title_tag.get_text(strip=True)
        # Clean up: "The Lectionary Page - Second Sunday in Lent" -> "Second Sunday in Lent"
        if " - " in t:
            t = t.split(" - ", 1)[1].strip()
        return t
    return None


def normalize_title(title: str) -> str:
    """Normalize a liturgical day title for use as a YAML key.

    Keeps the human-readable form but cleans up inconsistencies.
    """
    # Fix concatenated titles from the calendar page
    # e.g. "First Sunday after the EpiphanyThe Baptism of Our Lord"
    title = re.sub(r"([a-z])([A-Z])", r"\1 / \2", title)

    # Remove extra whitespace
    title = " ".join(title.split())

    return title.strip()


def main():
    print("Fetching calendar links...")
    links = fetch_calendar_links()
    print(f"  Found {len(links)} Sunday/feast pages")

    collects = {}
    errors = []

    for i, (href, link_text) in enumerate(links):
        url = f"{BASE_URL}/{href}"
        title = normalize_title(link_text)

        print(f"  [{i+1}/{len(links)}] {title}...", end=" ", flush=True)

        try:
            collect = extract_collect(url)
            if collect:
                collects[title] = collect
                print(f"OK ({len(collect)} chars)")
            else:
                print("no collect found (skipped)")
        except Exception as e:
            print(f"ERROR: {e}")
            errors.append((title, str(e)))

        if i < len(links) - 1:
            time.sleep(REQUEST_DELAY)

    # Write YAML
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        yaml.dump(
            collects,
            f,
            default_flow_style=False,
            allow_unicode=True,
            width=120,
            sort_keys=False,
        )

    print(f"\nDone! Extracted {len(collects)} collects to {OUTPUT_PATH}")
    if errors:
        print(f"Errors ({len(errors)}):")
        for title, err in errors:
            print(f"  - {title}: {err}")


if __name__ == "__main__":
    main()
