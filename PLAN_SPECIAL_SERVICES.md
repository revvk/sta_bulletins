# Implementation Plan: Proper Liturgies for Special Days & Special Covers

## Scope (Phase 1): Maundy Thursday + Good Friday
Palm Sunday will come after, since it reuses more of the normal Sunday flow.

## User Decisions
- Palm Sunday: All three services (8am, 9am, 11am) follow the Palm Sunday liturgy
- Maundy Thursday & Good Friday: Service time is "7 pm" on cover and filename
- Passion Gospels: User will provide pre-tagged voice parts as separate files
- Build order: Maundy Thursday + Good Friday first, then Palm Sunday

## Architecture: Option C — Per-Service Section Modules

Create per-service modules (`maundy_thursday.py`, `good_friday.py`, `palm_sunday.py`) in
`bulletin/document/sections/` that compose from existing building blocks. A lightweight
routing layer in `BulletinBuilder.build()` delegates to the appropriate flow.

**Why Option C:**
- Each special service is a fundamentally unique liturgical flow (not parameterized variations)
- The actual content primitives (add_body, add_rubric, add_dialogue, etc.) are all reusable
- Keeps the builder clean; scales well as more special services are added
- Each module is self-contained and easy to maintain

## Phase 1: Foundation

### 1a. Special Cover Selection
- **`templates.py`**: Add `cover_template` parameter to `load_front_cover()` (currently hardcodes `front_cover.docx`)
- **`builder.py`**: Add `_COVER_TEMPLATES` mapping (liturgical title keyword → template filename)
  - "palm sunday" → `front_cover_palm_sunday.docx`
  - "good friday" → `front_cover_good_friday.docx`
- Templates already exist in `templates/` directory

### 1b. Promote Shared Helper Functions
Make private helpers in `word_of_god.py` and `holy_communion.py` importable:
- From `word_of_god.py`: `_add_reading`, `_add_psalm`, `_add_gospel`, `_add_confession`,
  `_add_nicene_creed`, `_add_song_smart`, `_add_body_with_amen`, `_add_celebrant_with_cross`,
  `_add_standard_opening`
- From `holy_communion.py`: `_add_prayer_a_or_b`, `_add_sanctus_text`,
  `_add_offertory_rubric`, `_add_doxology_amen`, `_add_agnus_dei_spoken`
- Extract duplicate `_add_song_smart` into a shared location

### 1c. Service Type Detection
- **`rules.py`**: Add `detect_special_service(title)` → "palm_sunday" | "maundy_thursday" | "good_friday" | None
- **`builder.py`**: Store `self.special_service` in `__init__()`, apply rule overrides:
  - Maundy Thursday: FESTAL acclamation (not Lenten), no penitential order
  - Good Friday: Different acclamation ("Blessed be our God"), no Eucharist

## Phase 2: Maundy Thursday

### Data: `bulletin/data/bcp_texts/holy_week_maundy_thursday.yaml`
- Foot washing rubrics and music references
- Maundy Thursday Anthem text ("The Lord Jesus, after he had supped...")
- Short bidding-prayer POP form (BCP p. 383 style)
- Stripping of altar rubrics

### Module: `bulletin/document/sections/maundy_thursday.py`
Flow:
1. Word of God (mostly standard, FESTAL acclamation)
   - Processional → Opening Acclamation (FESTAL) → Collect for Purity → Song of Praise
   - Collect of the Day → First Reading → Psalm → Sequence Hymn
   - Gospel (normal, NOT multi-voice) → "The Homily" (heading change)
   - NO Nicene Creed
2. Foot Washing section (UNIQUE)
   - Rubric + music during washing + Maundy Thursday Anthem
3. Short bidding-prayer POP → Confession → Peace
4. Holy Communion (modified ending)
   - Offertory → Doxology → Great Thanksgiving → Sanctus → Eucharistic Prayer
   - Lord's Prayer → Breaking of Bread → Communion
   - Post-Communion Prayer
   - Post-Communion Hymn (new slot, replaces closing hymn)
   - NO blessing, NO dismissal
5. Stripping of the Altar + Psalm 22 → silence

## Phase 3: Good Friday

### Data: `bulletin/data/bcp_texts/holy_week_good_friday.yaml`
- Opening acclamation ("Blessed be our God" / "For ever and ever. Amen.")
- Solemn Collects (BCP pp. 277-280) — 9 bidding prayers
- Veneration of the Cross antiphon
- Closing prayer ("Lord Jesus Christ, Son of the living God...")

### Module: `bulletin/document/sections/good_friday.py`
Flow (NOT a Eucharist):
1. Entrance in silence (no prelude, no processional)
2. Opening Acclamation → Collect of the Day
3. First Reading → Psalm → Passion Gospel (John, multi-voice) → Sermon
4. Solemn Collects (UNIQUE — 9 bidding prayers with silence + collect pattern)
5. Veneration of the Cross (hymns + antiphon)
6. Lord's Prayer (standalone, no Eucharistic context)
7. Closing Prayer → silence (no blessing, no dismissal)

## Phase 4: Passion Gospel Rendering

### Data format: `bulletin/data/passion_gospels/{gospel}_year_{letter}.yaml`
```yaml
title: "The Passion of our Lord Jesus Christ according to John"
characters:
  narrator: { label: "Narrator", style: "normal" }
  jesus: { label: "✠ Jesus", style: "normal" }
  assembly: { label: "Assembly", style: "bold" }
segments:
  - voice: narrator
    text: "..."
  - voice: assembly
    text: "Crucify him!"
```
User will provide the pre-tagged text files; we parse into this YAML schema.

### Module: `bulletin/document/sections/passion_gospel.py`
- Renders multi-voice gospel with character labels
- Assembly parts in bold (congregation reads them)
- No opening/closing gospel responses

## Phase 5: Palm Sunday (after Maundy Thu + Good Friday)

### Data: `bulletin/data/bcp_texts/holy_week_palm_sunday.yaml`
- Liturgy of the Palms (opening dialogue, palm blessing, procession)
- Great Litany (BCP pp. 148-154) — replaces Creed + POP + Confession

### Module: `bulletin/document/sections/palm_sunday.py`
Flow:
1. Liturgy of the Palms (UNIQUE)
   - Special opening acclamation → Opening Prayer → Palm Gospel → Blessing of Palms → Procession
2. Word of God (modified)
   - NO Opening Acclamation/Collect for Purity/Song of Praise
   - Collect → Reading → Psalm → Sequence → Passion Gospel (Synoptic)
   - Silence (no sermon) → Great Litany → Peace
3. Holy Communion (standard with Lent rules)

## Phase 6: CLI Changes

### `generate.py` modifications:
- Detect weekday feasts from Google Sheet `service_type` field
- For Sunday dates (including Palm Sunday): generate 8am, 9am, 11am as usual
- For weekday dates (Maundy Thursday, Good Friday): generate one bulletin at "7 pm"
- Music for weekday services: use Service Music tab (same as 11am)
- Filename: e.g., `2026-04-02 - Maundy Thursday A - 7 pm (HEII-A) - Bulletin.docx`

## What Already Exists
- All Holy Week collects in collects.yaml
- Holy Week proper preface in proper_prefaces.yaml
- Palm Sunday through Maundy Thursday blessing (Prayer over the People)
- All eucharistic prayers (A, B, C)
- Google Sheet data for weekday feasts (dates, titles, readings, clergy, music)
- Special cover templates: `front_cover_palm_sunday.docx`, `front_cover_good_friday.docx`
- Rules engine already detects Holy Week via `_is_holy_week()`
