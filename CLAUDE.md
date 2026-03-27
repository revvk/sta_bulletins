# Project Overview
This project is a script to build bulletins for St. Andrew's Episcopal Church, taking data from the planning spreadsheets, stored data (mostly in YAML), and various other document templates and images.

## Project Pieces Completed:
9 am Sunday bulletin - generates successfully!
11 am Sunday bulletin - generates successfully!
8 am Sunday bulletin - generates successfully!
Reading sheets - generate successfully!
Holy Week bulletins - in progress.

## Scripture Poetry Formatting (Resolved)
* Scripture readings with poetry (e.g., Matthew 21:5/9, Philippians 2:6-11, Ephesians 5:14) are now properly detected and rendered with correct indentation.
* The solution uses Oremus's own HTML `<br>` tag CSS classes to detect poetry structure and indent levels — no BibleGateway needed.
* Poetry is rendered with three indent levels via dedicated styles: "Reading (Poetry)", "Reading (Poetry Indent 1)", "Reading (Poetry Indent 2)".

## Maundy Thursday (Special Service) - Resolved
* Front cover template now used (front_cover_maundy_thursday.docx)
* Foot washing invitation text: sourced from BCP p. 274 — flagged in YAML comment for review
* Anthem spacers between Celebrant and congregation sections: fixed
* Doxology and Sanctus wrapped in 1x1 table cells to prevent page splits (systemic fix across all bulletins)
* Fraction anthem now uses Agnus Dei with music notation images (same as 11am Lent)
* "Stripping the Altar" changed to "Stripping of the Altar"

## Good Friday (Special Service) - Resolved
* Removed space between People's line and "Let us pray" in opening
* Isaiah reading cache cleared; re-fetched with poetry formatting
* Smart quotes applied to John passion gospel YAML (Matthew already had them)
* Passion gospel part labels: lowercase + small caps (both Palm Sunday and Good Friday), 9pt, 7pt space after
* Passion gospel opening: rubric with John 19:17 reference, preamble Narrator line, responses rubric, then text
* Solemn Collects: rubric updated ("The Deacon reads the biddings..."), celebrant collects formatted with Celebrant label and tab
* Veneration anthem: restructured with interleaved bold/normal sections per BCP formatting
* Veneration hymn rubric added for musical piece during devotion

## Both Maundy Thursday and Good Friday - Resolved
* Inside back covers inserted before back cover (inside_back_cover_maundy_thursday.docx, inside_back_cover_good_friday.docx)

## Hidden Springs Bulletins - In Progress
* LOW (Liturgy of the Word) bulletin generates successfully for March 25 Annunciation
* Uses `--service hidden_springs` flag: `python generate.py 2026-03-25 --service hidden_springs`
* Data fetched from Hidden Springs Planner sheet (GID 2077691139)
* Large-print styles (LP_) on US Letter paper (8.5"x11") with 16pt body text
* Front/back cover templates: senior_living_front_cover.docx / senior_living_back_cover.docx
* Back cover auto-populates next 3 upcoming Wednesday services (bold date/time, normal title, italic clergy role)
* Title formatting: feast days use title as-is; regular weeks prepend "Wednesday after the [Sunday Name]"
* Service time: 11:00 am (displayed in footer, front cover, and upcoming services)
* Songs render in 1×N tables (one row per verse/chorus) for large-print readability
* Gloria renders as four separate keep-together sections matching BCP structure
* Canticle support: Canticles 9, 13, 15–21 in canticles.yaml, reusing psalm rendering pipeline

### Hidden Springs Music Library (Resolved)
* `hidden_springs_songs.yaml` — 104-entry catalog mapping all 113 AAC files to songs with lyrics
* Covers Hymnal 1982 (56 songs), Other (22 songs), Prelude/Postlude (26 pieces)
* 53 songs have full lyrics (from songs.yaml + Google Sheet); 31 are instrumental/no-lyrics
* Scaffold script: `scripts/build_hs_catalog.py` builds catalog from AAC files + lyrics sources
* HS-specific lookup chain: searches `hidden_springs_songs.yaml` first, falls back to `songs.yaml`
* AAC variant tracking: verse counts (`[Nv]`) and instruments (`[PIANO]`/`[ORGAN]`) parsed from planner
* Verse slicing: when `[Nv]` specified, only prints that many verses (matching AAC file)
* AAC manifest printed after generation — flat list of exactly which files to upload to Squarespace
* 7 songs still missing lyrics (render as title-only): His eye is on the sparrow, I come to the garden alone, Jesus loves me, Love lifted me, Praise God from whom all blessings flow, The King is coming, This is my Father's world

### Hidden Springs Outstanding:
* HE-II (Holy Eucharist) service type — needs separate builder for communion section
* Hidden Springs-specific POP forms (custom prayers referencing HS ministries)
* Front cover template has typo: "Chruch" → "Church" (needs template .docx fix)
* Special BOS blessings not rendering (e.g., Easter Season, Pentecost blessings from BOS)
* 7 songs still need lyrics added to hidden_springs_songs.yaml

## Project Pieces Outstanding:
* Hidden Springs HE-II services
* Hidden Springs-specific POP forms
* 7 missing HS song lyrics

# Remember:
All the the propers (which is the Episcopal word for the readings and prayers that are assigned for that day) and all the formatting styles are consistent across the 8 am, 9 am, and 11 am bulletins, as well as special service bulletins.

Architect the code so that as much of the code base can be reused as possible. We will need to be able to specify when running generate.py which bulletin to generate, or generate all of them.
