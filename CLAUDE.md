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

## Project Pieces Outstanding (detailed notes for next piece to work on are below):
* Hidden Springs Bulletins (not started)

# Remember:
All the the propers (which is the Episcopal word for the readings and prayers that are assigned for that day) and all the formatting styles are consistent across the 8 am, 9 am, and 11 am bulletins, as well as special service bulletins.

Architect the code so that as much of the code base can be reused as possible. We will need to be able to specify when running generate.py which bulletin to generate, or generate all of them.
