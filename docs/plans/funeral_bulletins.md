# Plan: Funeral & Memorial Bulletin Support

## Context

Funerals and memorials at St. Andrew's currently bypass the generator. Each
one is a fresh hand-edit of a Pages template — typing in scripture passages,
copy-pasting prayers from the BCP, looking up hymn lyrics, formatting the
participants page. That's exactly the work the Sunday/special-service
generator was built to eliminate.

The good news from a design standpoint: every funeral bulletin sampled —
Rite I and Rite II, with Eucharist and without, full burial and
memorial-only, even graveside — is built from **the exact same paragraph
styles, dialogue structure, song tables, and footer treatment** that the
existing builder already produces. Page setup matches, fonts match, the
rubric/heading hierarchy matches. The structural difference between a
Sunday bulletin and a funeral bulletin is *what* you put inside the same
containers, not how the containers themselves work.

So this is mostly a content-and-routing project, not a styling project.
The core work is: a YAML schema for funeral-specific decisions, a YAML
library of funeral-specific BCP texts (which the generator doesn't yet
know about), a builder module that wires them together, a new front cover
template, and a web planner page that walks the priest or family through
the choices.

## What an actual funeral bulletin looks like

After reading the planner, the guide, and bulletins from Cox (Rite II HC),
Stuhler (Rite I HC memorial), Owens (Rite II No-HC), Rucker (graveside),
and JB Smith (memorial No-HC), the body of every variant follows the same
backbone:

```
COVER             — name, dates of life, "The Burial of the Dead" /
                    "Memorial Service", decorative "Alleluia" in gray
                    Garamond italic, ichthys mark, service date in
                    vertical strip, "Invite Involve Instruct Inspire"
                    footer.

INSIDE COVER      — (optional) photo + biographical paragraphs,
                    two-column.

WORD OF GOD       — Seating of the Family, Prelude, [stand], Opening
                    Anthem, Opening Hymn, The Collect, [seated],
                    1st Reading, Psalm, [optional 2nd Reading], [stand],
                    Sequence Hymn, [stand], Gospel, [seated],
                    [optional Words of Remembrance + Hymn], Homily,
                    Apostles' Creed, Prayers for the Departed,
                    [optional special prayer — Daughters of the King,
                    veteran, parent-of-deceased-child, etc.], The Peace.

THE END OF THE
SERVICE           — Five independent toggles, ordered as below when
                    present:

                    1. Holy Eucharist?
                       If yes: Offertory Anthem (with soloist credit),
                       Great Thanksgiving (Eucharistic Prayer
                       I/II/A/B/C/D depending on rite), Lord's Prayer,
                       Breaking, Invitation, Communion Music (mix of
                       hymnal refs and full-lyric songs),
                       Postcommunion Prayer.

                    2. Commendation?
                       Anthem ("Give rest, O Christ…") and commendation
                       prayer ("Into your hands, O merciful Savior…").
                       Commends the soul. Used when remains are present
                       but committal does NOT immediately follow at the
                       church — i.e. the body is at the funeral but the
                       burial is delayed or happens elsewhere later.

                    3. Closing music + dismissal block?
                       (Closing Hymn → Dismissal → Postlude)
                       Always present except when committal happens
                       in-building immediately, in which case the
                       procession hymn from step 4 IS the closing music
                       and the dismissal lives inside the committal.

                    4. Committal?
                       Procession hymn (when the committal happens at
                       the church columbarium and everyone walks
                       together) OR no procession music (when committal
                       is at a separate location).
                       Then THE COMMITTAL: Anthem ("Everyone the Father
                       gives to me…"), commendation/committal prayer
                       ("In sure and certain hope…"), Lord's Prayer,
                       "Rest eternal" antiphon, dismissal blessing.
                       Used when the committal is part of THIS bulletin
                       — i.e. everyone present is going to it.

                    5. Blessing.
                       The "God of peace, who brought again from the
                       dead our Lord Jesus Christ…" dismissal blessing.
                       Always present somewhere, but its position
                       depends on which of the above are in play —
                       see below.

BACK              — (optional) Reception notice in italic centered text.
                    Participants page (Celebrant, Assisting Clergy,
                    Deacon, Readers, Crucifer, Chalice Bearers,
                    Musicians). "The text of the service…" boilerplate.
                    Logo.
```

### The Commendation × Committal × HC matrix

These three are independent. Every combination is liturgically valid:

| HC  | Commendation | Committal | When it happens                                                | Bulletin shape after The Peace                                                                                     |
|:---:|:------------:|:---------:|----------------------------------------------------------------|--------------------------------------------------------------------------------------------------------------------|
| Yes | No           | No        | Memorial with HC, no remains present, no committal             | HC → Blessing → Closing Hymn → Dismissal → Postlude                                                                |
| Yes | No           | Yes       | Funeral with HC; committal immediately at the church           | HC → Procession hymn → **The Committal** (anthem + prayer + LP + Rest eternal + Blessing-as-dismissal) → Postlude  |
| Yes | Yes          | No        | Funeral with HC; remains present but committal delayed/elsewhere | HC → **The Commendation** → Blessing → Closing Hymn → Dismissal → Postlude                                       |
| Yes | Yes          | Yes       | Funeral with HC; commend at church, then go elsewhere/later for committal | HC → **The Commendation** → Closing Hymn → Dismissal → Postlude → (separate Committal page, often a tear-off insert) |
| No  | No           | No        | Memorial without HC, no remains present                        | Blessing → Closing Hymn → Dismissal → Postlude                                                                     |
| No  | No           | Yes       | Funeral without HC; committal immediately at the church        | Procession hymn → **The Committal** → Postlude                                                                     |
| No  | Yes          | No        | Funeral without HC; remains present but committal delayed/elsewhere (the Owens case) | **The Commendation** → Blessing → Closing Hymn → Dismissal → Postlude                                              |
| No  | Yes          | Yes       | Funeral without HC; commend at church, then elsewhere for committal | **The Commendation** → Closing Hymn → Dismissal → Postlude → (separate Committal section/insert)                  |

The schema captures this as three independent flags (`holy_eucharist`,
`include_commendation`, `include_committal`); the builder's job is just
to walk that table and emit the right blocks in the right order.

### Rite I vs Rite II — entirely separate liturgies

Rite I and Rite II are **complete, independent liturgies in the BCP**, not
variations of one another. Different opening anthems, different collect
language, different Apostles' Creed wording, different Prayers for the
Departed (totally different forms — Rite I is a long sequence each ending
in **Amen**; Rite II is the Lazarus-style litany ending in **Hear us,
Lord**), different commendation and committal prayer wording, different
Sanctus, different Eucharistic Prayers (Rite I: Prayer I and Prayer II;
Rite II: Prayer A, B, C, D), different postcommunion prayer.

`funeral_texts.yaml` will store the two rites as separate top-level
sub-trees (`rite_I:` and `rite_II:`), each transcribed verbatim from the
1979 BCP. **No derivation.** The text for each rite must come straight
from the BCP source — Andrew has indicated he'll provide the BCP PDF for
this; without it the safe move is to use the published BCP text via
another verified source rather than paraphrase.

The same separation principle applies to the two new Eucharistic Prayers
needed for Rite I (`prayer_i`, `prayer_ii`) and the still-missing Rite II
Prayer D (`prayer_d`) in `eucharistic_prayers.yaml` — those get
transcribed from the BCP verbatim, not adapted from prayer_a.

## What gets reused unchanged

- All paragraph styles in `bulletin/document/styles.py` (with the
  keep_with_next flags — funerals will benefit immediately).
- `bulletin/document/formatting.py` helpers: `add_heading`, `add_heading2`,
  `add_introductory_rubric`, `add_rubric`, `add_celebrant_line`,
  `add_people_line`, `add_dialogue`, `add_body`, `add_song`,
  `add_song_two_column`, `add_no_split_block`, the cross-symbol helper.
- Scripture rendering with poetry indent detection (some burial readings
  — Isaiah 25, Isaiah 61, Lamentations 3, Wisdom 3 — are poetry).
- Psalm rendering (responsive / unison / antiphonal modes — burial
  bulletins use unison most often).
- The Apostles' Creed (Rite II), the Lord's Prayer, the Sanctus,
  Eucharistic Prayers A, B, and C — already in the bcp_texts YAML files.
- Page setup, footers (just a different left footer text: "Burial of
  the Dead, Rite I/II" / "Memorial Service" / etc.).
- Songs — the funeral hymn list in Appendix B of the guide is mostly
  already in `songs.yaml` since they're standard Hymnal 1982 entries.
  A few St. Andrew's Songbook entries may need adding.
- The cover-template floating-shape pin and the keep-with-next style
  work that already shipped.

## What's new

### New data files

| Path | Purpose |
|---|---|
| `bulletin/data/funerals/funeral_texts.yaml` | The fixed BCP texts unique to the burial office, organized as `rite_I:` and `rite_II:` independent sub-trees, each containing: opening anthems, prayers for the departed, commendation anthem, commendation prayer, committal anthem, committal prayer, "Rest eternal" antiphon, "God of peace" dismissal blessing, the procession-to-columbarium rubric. Each text transcribed verbatim from the BCP — no derivation between the two rites. |
| `bulletin/data/funerals/special_prayers.yaml` | A library of optional add-on prayers: "Prayer for a Daughter Who Has Died" (Daughters of the King), "Prayer for a Veteran", "Prayer for a Parent of a Deceased Child", etc. Each entry knows its placement (`after_prayers_for_departed` is currently the only slot used). |

**Burial readings are not stored separately.** The five-OT / six-NT /
five-Gospel / six-Psalm menu the planner walks through is just a list
of BCP suggestions; the actual reading text gets fetched and cached
through the existing scripture pipeline like any Sunday reading. The
short menu of suggested options can live inline in `funeral_texts.yaml`
or as a constant in the web planner module — whichever is convenient
when that step is built.
| `bulletin/data/funerals/services/<slug>.yaml` | One per funeral. Schema below. Lives next to the generated `.docx`; the web planner writes it, the CLI reads it. Committable so we have a record of every service built. |
| `bulletin/data/bcp_texts/eucharistic_prayers.yaml` (extension) | Add `prayer_d` (Rite II) and `prayer_i`, `prayer_ii` (Rite I), each transcribed verbatim from the BCP. |

### Per-service YAML — the schema

```yaml
# bulletin/data/funerals/services/2026-01-31-cox.yaml
service_kind: burial          # burial | memorial | committal
rite: II                      # I | II
holy_eucharist:
  enabled: true
  prayer: A                   # A | B | C | D for Rite II; I | II for Rite I
include_commendation: false   # see matrix above
include_committal: true       # see matrix above

deceased:
  full_name: 'Dorothy "Annette" Woodall Huddleston Cox'
  preferred_name: Annette     # used in prayers ("for our sister Annette")
  pronoun: she                # she / he / they — drives commendation/committal pronouns
  born:  1932-11-30
  died:  2026-01-08
  bio: |                      # optional; rendered on the inside front cover
    Dorothy "Annette" Woodall Huddleston Cox returned home...
    [free-form paragraphs; two-column layout]
  photo: photos/annette-cox.jpg   # optional

service:
  date: 2026-01-31
  start_time: '11:00 am'
  location: Great Hall        # Great Hall | Chapel | <funeral home name>
  committal_location: columbarium   # columbarium | cemetery | scattering | none
                                    # only meaningful when include_committal: true

participants:
  celebrant: Andrew Van Kirk
  assisting:
    - Logan Hurst
    - Paulette Magnuson
  deacon: Katie Gerber
  readers:
    - Lee Herrin
    - Cody Hutchinson
    - Roland Towery
  crucifer: Sandi Krummel
  chalice_bearers: []         # empty when no HC
  musician: Laura Bray
  vocalist: Laura Bray
  pall_bearers: []            # listed only if family wants them in the bulletin
  words_of_remembrance:
    - Jeffrey Lon Huddleston  # one or more names; omit field if no remarks

readings:
  first:    'Isaiah 25:6-9'   # any reference, but the planner UI restricts to the burial menu
  psalm:    'Psalm 139:1-11'
  psalm_mode: unison          # unison | responsive | antiphonal | by_half_verse
  second:   '2 Corinthians 4:16-5:9'   # omit for one-reading service
  gospel:   'John 14:1-6'

music:
  # Two slot kinds throughout the music section:
  #   *_hymn   — congregational, full lyrics printed (or hymnal ref).
  #              Schema value: a string (the song title).
  #   *_anthem — soloist or choir, just title + composer/arranger.
  #              Schema value: { title, arranger, soloist } dict.
  # The post-Words-of-Remembrance and post-Homily slots support
  # EITHER kind via paired _hymn / _anthem fields; both are
  # independently optional, both default to null.

  opening_anthem: from_bcp    # 'from_bcp' is the default (renders from funeral_texts.yaml)
                              # could be a song key in songs.yaml for a sung opening
  opening_hymn:    "Love divine, all loves excelling"
  sequence_hymn:   "Amazing grace! how sweet the sound"

  # After the Words of Remembrance / Eulogy. Both optional and
  # independent — a service can have either, neither, or both.
  remembrance_hymn:  "Great is thy faithfulness"     # congregational hymn (Cox case)
  remembrance_anthem: null                           # OR an anthem (sung by soloist/choir)

  # After the Homily. Same pattern — both optional and independent.
  # Cox has neither; Owens has the homily_anthem slot ("Surely It Is
  # God Who Saves Me" by Jack Noble White).
  homily_hymn:   null
  homily_anthem: null

  # Offertory anthem — only meaningful when holy_eucharist.enabled.
  offertory_anthem:
    title: "Blessed Assurance"
    arranger: "arr. William Cutter"
    soloist:  "Lexi Johnson, Soprano"

  communion_music:
    - "Be thou my vision, O Lord of my heart"
    - "Just as I am, without one plea"
    - "It Is Well With My Soul"

  procession_hymn: "Lift high the cross"             # only when committal happens at the church
  closing_hymn:    null                              # No-committal-at-church variants use this

collects:
  adult:
    choice: option_3                  # option_1 | option_2 | option_3 (Rite II)
                                      # ignored for Rite I (only one Adult option)
  add_for_those_who_mourn: false      # Rite II only — append the addendum
                                      # after whichever collect was chosen

prayers_for_the_departed:
  include_baptism_petition:   true    # Rite II — include the baptized-person petition?
  include_communion_petition: true    # Rite II — include the communicant petition?
  conclusion: commend                 # Rite II — commend | father_of_all
                                      # (Rite I has no choice here)

special_prayers:                      # references entries in special_prayers.yaml
  - key:         daughters_of_the_king
    reader:      'Diane Victor'
    reader_role: "President of the St. Andrew's Chapter of the Daughters of the King"

reception:
  shown: true
  text: |                                            # if shown, italic centered text
    All are invited to join the family in
    Michie Hall at the reception hosted by
    The Daughters of the King.

interment_notice: null         # optional italic centered notice for No-Committal services,
                               # e.g. "Wanda's remains will be interred on April 14
                               # in Dreamland Cemetery, Canyon, Texas."
```

Every field has a sensible default, so a minimal funeral YAML can be ten
lines. The web planner writes the full file; a power user can also
hand-edit.

Three reference per-service YAMLs are committed under
`bulletin/data/funerals/services/`:

| File | Case | Distinguishing feature |
|---|---|---|
| `2026-01-31-cox.yaml`     | Rite II HC, body present, committal at the church | Words of Remembrance + DOK prayer + procession to columbarium |
| `2025-03-07-stuhler.yaml` | Rite I HC, Memorial (no body)                    | Memorial form; Rite I `{patron_phrase}` for petition 10 |
| `2026-04-10-owens.yaml`   | Rite II No-HC, body present, deferred interment  | `interment_notice` populated; post-homily anthem |

These three between them cover the structurally meaningful rows of
the HC × Commendation × Committal matrix — every other row is just a
combination of features already exercised here. The exercise of
writing them surfaced the schema additions called out above:

- `collects.adult.choice` / `collects.add_for_those_who_mourn` (Rite II)
- `prayers_for_the_departed.include_baptism_petition` /
  `include_communion_petition` / `conclusion` (Rite II)
- `prayers_for_the_departed.include_optional_petitions` /
  `patron_phrase` (Rite I — Stuhler exposed these)
- `music.remembrance_anthem` and `music.homily_hymn` /
  `music.homily_anthem` (Owens exposed the post-homily anthem slot;
  matching paired hymn/anthem fields added for symmetry)

None of the original v1 sketch was wrong; the YAMLs simply revealed
real decisions families make that the schema needed to express. The
plan and the per-service YAMLs are now in sync.

### New code modules

| Path | Purpose |
|---|---|
| `bulletin/document/sections/burial.py` | The funeral builder. Mirrors `add_word_of_god` / `add_holy_communion` but specialized for the burial office. Reads the per-service dict, looks up scripture/prayers/songs, walks the HC×Commendation×Committal matrix, calls existing helpers. ~400 lines, similar size to `holy_communion.py`. |
| `bulletin/document/builder.py` | Two new branches: `is_funeral` flag set when input is a funeral YAML, `_build_funeral(doc)` calls into `burial.py` analogous to `_build_maundy_thursday(doc)`. Front cover template selection and back cover assembly already follow the existing pattern. |
| `bulletin/sources/funeral_data.py` | Loader for the per-service YAML. Validates the schema, resolves date strings to `datetime`, expands defaults, returns a typed dataclass. |
| `bulletin/data/loader.py` | Add `load_funeral_texts()` and `load_special_prayers()` mirrors of the existing loader functions. |
| `templates/front_cover_funeral.docx` | **Two-page** template — outer cover on page 1, inner photo+bio cover on page 2. Same layout family as the Sunday cover but with the deceased name + dates as the visual anchor. Outer page has the "Alleluia" italic treatment in light gray, ichthys mark, vertical date strip, "Invite Involve Instruct Inspire" footer; inner page has the photo above and a two-column bio below. Text placeholders: `{{DATE}}`, `{{SUBTITLE}}` ("The Burial of the Dead" / "Memorial Service" / "Graveside Service"), `{{NAME}}` (single string — text frame wraps long names automatically), `{{LIFE_DATES}}`, `{{BIO}}`. The photo is an embedded image whose binary the builder swaps out at generation time (no `{{PHOTO}}` text placeholder needed). For services with neither bio nor photo, the user deletes the inner page in the post-generation hand-edit pass — simpler than conditionally assembling a one-page-or-two-page document. |

### CLI

```
python generate.py --funeral bulletin/data/funerals/services/2026-01-31-cox.yaml
```

Or, when the web flow is in place:

```
python generate.py --funeral 2026-01-31-cox            # by slug
```

Output: `2026-01-31 - Burial of the Dead - Annette Cox.docx` next to
the Sunday output dir, plus the same Run Report telling Andrew which
scripture passages were fetched and which song lyrics were missing.

### Web flow

A new "Funerals" section in the existing web app, sitting alongside
Generate / Songs / Prayers:

- **`/funerals`** — List of all per-service YAMLs, with date, name,
  status (generated / not generated yet).
- **`/funerals/new`** — The planner form. Mirrors the structure of the
  paper planner (pages 6-8 of `Funeral Planner v2024.pdf`):
  1. Deceased details (name, cover spelling, preferred name, pronoun,
     dates, bio, photo upload)
  2. Service kind (burial / memorial / graveside) + date + location
  3. Rite I or Rite II radio
  4. Holy Communion yes/no, then conditional prayer letter
  5. Commendation yes/no, Committal yes/no (with helper text from the
     matrix to make it clear what each combination produces)
  6. Readings — checkboxes presenting only the BCP-suggested burial
     options (5 OT, 6 NT, 5 Gospels, 6 psalms), with the one-line
     descriptions from the planner. "Other" text field for non-
     suggested readings. The list of suggestions is small enough to
     live inline in `funeral_texts.yaml` or as a constant in the
     planner module — pick one when building this step. The reading
     text itself is fetched and cached through the existing scripture
     pipeline; nothing new on the data side.
  7. Music — text inputs (with autocomplete from `songs.yaml`) for
     each music slot. Slots conditional on rite + HC choice + committal
     location.
  8. Participants — text inputs.
  9. Special prayers — checkboxes from `special_prayers.yaml`.
  10. Reception — toggle + free-form text.
  11. Interment notice — toggle + free-form text (for the "remains
      will be interred…" blurb on No-Committal bulletins).
- **`/funerals/<slug>/edit`** — Same form, populated from existing YAML.
- **`/funerals/<slug>/generate`** — Runs `BulletinBuilder` with the
  funeral YAML, returns the Run Report and a Reveal-in-Finder button
  for the .docx.

The form should save partial state as the priest works through it (the
planning conversation with a family often spans multiple sittings).

## Build order

Each step ships independently and leaves the existing system untouched:

1. **Funeral YAML data**, no code.
   - Author `funeral_texts.yaml` with `rite_I:` and `rite_II:` sub-trees,
     each transcribed verbatim from the BCP (opening anthems, prayers
     for the departed, commendation, committal, blessings).
   - Author `special_prayers.yaml` (Daughters of the King to start).
   - Add Eucharistic Prayer D (Rite II) and Prayers I + II (Rite I)
     to `eucharistic_prayers.yaml`, also transcribed verbatim.
   - Reviewable by Andrew without any UI.

2. **One real per-service YAML**, hand-written. Pick the Cox bulletin
   (because we already have its rendered PDF as a reference) and write
   `bulletin/data/funerals/services/2026-01-31-cox.yaml` by hand. This
   forces the schema to be complete and gives us a fixed regression
   target for step 4.

3. **Funeral cover template** (single two-page `.docx` —
   `templates/front_cover_funeral.docx`). Page 1 outer cover, page 2
   inner photo+bio cover. Verified by loading through python-docx
   and substituting all five placeholders via the existing
   `_replace_all_placeholders` machinery (already shown to handle
   the run-split tags Word inserts when the spell-checker has
   touched the placeholder text). For services with no bio/photo,
   the inner page is deleted in the post-generation hand-edit pass.

4. **`bulletin/document/sections/burial.py` + builder hook**. The
   funeral builder. Generate the Cox bulletin from the YAML and diff
   against the reference PDF. Goal: byte-equivalent text, layout
   deviations only on whitespace tuning.

5. **CLI**. `--funeral` flag. Now Andrew can generate any future
   funeral by hand-writing a 30-line YAML.

6. **Web planner page**. The form, save/load, generate button. This
   is the largest single piece of UI work but reuses the existing
   FastAPI/Jinja/HTMX scaffold.

7. **One more example end-to-end**. Pick a Rite I memorial without HC
   (Stuhler family, since we have a PDF reference). Build its YAML
   through the web form, generate, compare. Catches Rite-I-specific
   rendering issues that the Cox case doesn't exercise.

## What's intentionally out of scope (for v1)

- **Reception cards / committal-only sub-bulletins**. These exist as
  separate physical printouts at some funerals (e.g. Cox includes a
  "Readings and Prayers" booklet for the family). The data is already
  in the per-service YAML so this is a pure-rendering follow-up.
- **The full funeral planner** (Parts 1 + 2 of `Funeral Planner v2024.pdf`
  — personal/estate info). That's a personal-document workflow, not
  a bulletin-generation workflow. Out of scope here. The web planner
  only covers Part 3 ("Burial Service Planner", pages 6-8).
- **Auto-import of historical funerals**. The 50+ existing Pages
  bulletins won't be retroactively imported. They're the reference
  corpus for visual fidelity, not an input dataset.
- **Funeral-specific propers fetching from a calendar**. Funerals
  don't follow the lectionary; readings are family-chosen. The BCP
  burial-menu checkboxes (in the planner UI) plus an "Other" override
  are the entire selection mechanism; the reading text itself flows
  through the existing scripture cache.
- **Multiple-deceased services** (joint funerals — rare but they exist).
  Schema would need `deceased: [list]`. Defer.

## Verification

1. **Cox bulletin reproduction.** Generate from the hand-written YAML;
   visually diff against the reference PDF. Aim for parity on all 16
   content pages (the 4 blank pages between the Postlude and
   Participants are intentional whitespace and should match too).
2. **Stuhler bulletin reproduction (Rite I, with HC, memorial — no
   committal).** Tests the Rite I branch and the no-body branch.
3. **Owens reproduction (Rite II, no HC, with deferred-interment
   notice).** Tests the No-HC branch and the optional interment-notice
   block. Specifically: `include_commendation: true`,
   `include_committal: false`.
4. **CLI smoke** — `python generate.py --funeral <yaml>` produces an
   openable .docx and a Run Report.
5. **Web smoke** — Build a brand-new fictional funeral end-to-end
   through the planner form. Confirm the generated YAML round-trips
   through edit and re-save without losing any field.
6. **Funeral hymn coverage** — Cross-check Appendix B of the funeral
   guide against `songs.yaml`. Surface missing entries in the Run Report.

## Open questions before step 1

- **Where is the BCP PDF?** Step 1 needs verbatim BCP text for both
  rites. Andrew indicated he'll provide it; if it's not already in
  `source_documents/`, point at the file path or drop it in
  `source_documents/funerals_and_memorials/` and I'll work from there.
