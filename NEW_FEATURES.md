# Hidden Springs Bulletin Creation

Hidden Springs is a senior living center. St. Andrew's holds a weekly service on Wednesday mornings. Most of the congregants are not Episcopalian, and so we have monthly communion on the first Sunday of the month rather than weekly communion. The changes to the bulletin generation derive from these basic facts:

* Because the congregation is elderly, the bulletin is printed on bigger, US Letter, sized paper with bigger fonts.
* Preface all new styles with LP_ for large print, keeping the rest of the style name the same.
* There are custom front and back covers in the /templates folder. They are senior_living_front_cover.docx and senior_living_back_cover.docx because the general format is also used for occasional services at other senior living centers.
* Generally speaking, the service uses the propers for the Sunday prior, and the name of the day for the front cover is "Wednesday after the [Sunday Name].
** The exception is when the Wednesday is a feast day, in which case we use the Feast Day propers. In that case, the name of the day is just the name of the feast. March 25, the Feast of the Annunciation, is one such case.
* The propers for the service can be found in the "Hidden Springs Planner" sheet in the Rota & Liturgical Schedule Google Sheet.
* For the most part, except on first Sundays when we have Holy Communion, the service follows a pattern we call Liturgy of the Word (LOW). This is not actually in the BCP, but is more or less the proanophora from Holy Eucharist Rite II. The service type (LOW or HEII) is specified in column A of the "Hidden Springs Planner" sheet.
* I've included five bulletin examples with the LOW in /source_documents/Hidden_Springs/ — one each from the seasons of Advent (LOW), Epiphany (LOW), Lent (HEII), Easter (LOW), and after Pentecost (LOW). You should be able to derive the structure and seasonal variations from those, but please ask any questions you need. HEII follows the basic structure of the existing bulletins.
* The service time is 11:00 am.
* Hidden Springs has special versions of all the prayers of the people that reference their specific ministry context. These will be in the pop.yaml file, though if one is not found the fallback should be to go with the generic form of that prayer.
* There is no offering collected at the service except on first Sunday of the month communion Sundays. And in any case, no rubrics are needed about giving or the QR code.

## Completed

### LOW Bulletin Generation
* Annunciation (March 25) LOW bulletin generates correctly
* Uses `--service hidden_springs` flag: `python generate.py 2026-03-25 --service hidden_springs`
* Large-print styles, US Letter paper, custom front/back covers
* Back cover: upcoming 3 Wednesday services auto-populated with formatted bold/italic runs
* Canticle support (Canticles 9, 13, 15–21)
* Songs in 1×N tables (multi-row) for large-print readability
* Gloria as four separate keep-together BCP sections

### Music Library & AAC File Management
* `hidden_springs_songs.yaml` catalogs all 113 AAC files across 104 song entries
* Three categories: Hymnal 1982 (56), Other (22), Prelude/Postlude (26)
* Lyrics populated from songs.yaml (20 songs) and Google Sheet (33 songs); 31 are instrumental/no-lyrics
* Scaffold script: `scripts/build_hs_catalog.py` builds catalog from AAC filenames + lyrics sources
* HS-specific lookup chain: hidden_springs_songs.yaml first, songs.yaml fallback
* AAC variant tracking: verse counts [Nv] and instruments [PIANO]/[ORGAN] parsed from planner fields
* Verse slicing: when [Nv] specified, only that many verses are printed (matching the AAC file)
* AAC manifest printed after bulletin generation — flat list of files to upload to Squarespace
* Lyrics Google Sheet: https://docs.google.com/spreadsheets/d/1vX9llfMg0bAWZiSM10RaAsgExyKIN5YflKaaaEPoOOI/edit?usp=sharing

## Outstanding

### HE-II (Holy Eucharist) Service Type
* First-Wednesday communion services need separate builder wiring for the communion section
* Should follow the basic structure of existing 9am/11am HEII bulletins
* Offertory, Doxology, and Communion music slots need to be wired into the HS lookup chain

### Hidden Springs-Specific POP Forms
* Custom prayers of the people referencing HS-specific ministry context
* Should be in pop.yaml with fallback to generic forms

### Missing Song Lyrics (7 songs)
* His eye is on the sparrow
* I come to the garden alone
* Jesus loves me! this I know
* Love lifted me
* Praise God, from whom all blessings flow
* The King is coming (copyright — church has license)
* This is my Father's world

### Other
* Front cover template has typo: "Chruch" → "Church" (needs template .docx fix)
* Special BOS blessings not rendering (e.g., Easter Season, Pentecost blessings from BOS)
