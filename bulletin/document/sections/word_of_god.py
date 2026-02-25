"""
The Word of God section of the bulletin.

Covers everything from the processional hymn through The Peace:
  - Processional hymn
  - Opening Acclamation
  - Collect for Purity (standard) or Penitential Order (Lent)
  - Song of Praise / Kyrie / Advent Wreath
  - Collect of the Day
  - Children's Sermon
  - Scripture readings (Epistle/OT)
  - Psalm / Canticle
  - Sequence Hymn
  - Gospel
  - Sermon
  - Nicene Creed
  - Prayers of the People
  - Confession & Absolution (if not in Penitential Order)
  - The Peace
"""

from docx import Document

from bulletin.config import CROSS_SYMBOL
from bulletin.data.loader import load_common_prayers
from bulletin.document.formatting import (
    add_spacer, add_heading, add_heading2, add_rubric,
    add_introductory_rubric, add_body, add_celebrant_line,
    add_people_line, add_dialogue, add_cross_symbol,
    add_hymn_header, add_song, add_song_two_column,
    add_scripture_text,
)
from bulletin.logic.rules import SeasonalRules


def add_word_of_god(doc: Document, rules: SeasonalRules, data: dict):
    """Add the entire Word of God section.

    Args:
        doc: The Document.
        rules: SeasonalRules for this service.
        data: Dict with keys:
            - processional: song data dict (or None)
            - song_of_praise: song data dict (or None)
            - sequence_hymn: song data dict (or None)
            - collect_text: str (the collect of the day text)
            - reading_1_ref: str (e.g., "1 Corinthians 10:1-13")
            - reading_1_text: ScriptureReading
            - psalm_ref: str (e.g., "Psalm 63:1-8")
            - psalm_text: list of str lines
            - psalm_rubric: str (e.g., "Read in unison.")
            - gospel_ref: str (e.g., "Luke 13:1-9")
            - gospel_book: str (e.g., "Luke")
            - gospel_text: ScriptureReading
            - preacher: str (e.g., "The Rev. Andrew Van Kirk")
            - pop_elements: list of POP element dicts
            - pop_concluding_rubric: str
            - advent_wreath_verse: str | None (for Advent)
            - advent_hymnal_ref: str | None
            - penitential_sentence: str | None (for Lent 2-5)
    """
    prayers = load_common_prayers()

    # --- Penitential Order (Lent) or standard Word of God ---
    if rules.use_penitential_order:
        _add_penitential_order(doc, rules, data, prayers)
    else:
        _add_standard_opening(doc, rules, data, prayers)

    # --- Collect of the Day ---
    add_heading2(doc, "Collect of the Day")
    add_celebrant_line(doc, "Celebrant", "The Lord be with you.")
    add_people_line(doc, "People", "And also with you.")
    add_people_line(doc, "Celebrant", "Let us pray.")
    add_spacer(doc)
    _add_body_with_amen(doc, data["collect_text"])

    # --- Be seated for readings ---
    add_introductory_rubric(doc, "Be seated.")

    # --- Children's Sermon ---
    add_heading2(doc, "Children's Sermon")

    # --- First Reading ---
    _add_reading(doc, data["reading_1_ref"], data["reading_1_text"])

    # --- Psalm ---
    _add_psalm(doc, data["psalm_ref"], data.get("psalm_rubric", ""),
               data["psalm_text"])

    # --- Sequence Hymn ---
    add_introductory_rubric(doc, "Please stand.")
    add_heading2(doc, "Sequence Hymn")
    _add_song_smart(doc, data.get("sequence_hymn"))

    # --- Gospel ---
    _add_gospel(doc, data["gospel_ref"], data["gospel_book"],
                data["gospel_text"])

    # --- Sermon ---
    add_introductory_rubric(doc, "Be seated.")
    add_heading2(doc, "Sermon")
    p = doc.add_paragraph(style="Body")
    run = p.add_run(data.get("preacher", ""))
    run.bold = True

    # --- Nicene Creed ---
    add_introductory_rubric(doc, "Please stand.")
    add_heading2(doc, "The Nicene Creed")
    _add_nicene_creed(doc, prayers)

    # --- Prayers of the People ---
    add_heading2(doc, "Prayers of the People")
    _add_pop(doc, data.get("pop_elements", []))
    add_rubric(doc, data.get("pop_concluding_rubric",
                              "The Celebrant concludes with a suitable Collect."))

    # --- Confession & Absolution (if not already done in Penitential Order) ---
    if not rules.no_confession_after_pop:
        _add_confession(doc, prayers)

    # --- The Peace ---
    p = doc.add_paragraph(style="Body - Introductory Rubric")
    run = p.add_run("Please stand.")
    run.bold = True
    add_heading2(doc, "The Peace")
    p = doc.add_paragraph(style="Body - Celebrant")
    run = p.add_run("Celebrant")
    p.add_run("\t")
    run2 = p.add_run("The peace of the Lord be always with you.")
    run2.bold = True
    add_people_line(doc, "People", "And also with you.")


def _add_standard_opening(doc: Document, rules: SeasonalRules,
                          data: dict, prayers: dict):
    """Standard (non-Lent) opening: Word of God → Processional → Acclamation → Collect for Purity → Song of Praise."""
    add_heading(doc, "The Word of God")
    add_spacer(doc)

    # Processional
    add_introductory_rubric(doc, "Please stand.")
    add_heading2(doc, "Processional")
    _add_song_smart(doc, data.get("processional"))

    # Opening Acclamation
    add_introductory_rubric(doc, "Remain standing.")
    add_heading2(doc, "Opening Acclamation")
    cel_text = rules.acclamation_celebrant.replace("{cross}", CROSS_SYMBOL)
    _add_celebrant_with_cross(doc, "Celebrant", cel_text)
    add_people_line(doc, "People", rules.acclamation_people)
    add_spacer(doc)

    # Collect for Purity
    if rules.include_collect_for_purity:
        _add_body_with_amen(doc, prayers["collect_for_purity"])

    # Song of Praise / Kyrie / Advent Wreath
    add_heading2(doc, rules.song_of_praise_label)
    if rules.is_advent:
        _add_advent_wreath(doc, data)
    else:
        _add_song_smart(doc, data.get("song_of_praise"))


def _add_penitential_order(doc: Document, rules: SeasonalRules,
                           data: dict, prayers: dict):
    """Lenten opening: Penitential Order → Processional → Acclamation → [Decalogue/Sentence] → Confession → Word of God → Kyrie."""
    add_heading(doc, "A Penitential Order")
    add_spacer(doc)

    # Processional
    add_introductory_rubric(doc, "Please stand.")
    add_heading2(doc, "Processional")
    _add_song_smart(doc, data.get("processional"))

    # Opening Acclamation
    add_introductory_rubric(doc, "Remain standing.")
    add_heading2(doc, "Opening Acclamation")
    add_celebrant_line(doc, "Celebrant", rules.acclamation_celebrant)
    add_people_line(doc, "People", rules.acclamation_people)
    add_spacer(doc)

    # Decalogue (Lent 1) or Scripture Sentence (Lent 2-5)
    if rules.use_decalogue:
        _add_decalogue(doc, prayers)
    else:
        sentence = data.get("penitential_sentence", "")
        if sentence:
            add_body(doc, sentence)

    # Confession
    add_introductory_rubric(doc, "Please kneel or remain standing.")
    _add_confession(doc, prayers)

    # Then "The Word of God" section heading
    add_heading(doc, "")
    add_heading(doc, "The Word of God")
    add_heading(doc, "")

    # Kyrie
    add_introductory_rubric(doc, "Please stand.")
    add_heading2(doc, "Kyrie")
    _add_song_smart(doc, data.get("song_of_praise"))


def _add_confession(doc: Document, prayers: dict):
    """Add the Confession of Sin and Absolution."""
    add_heading2(doc, "Confession of Sin")
    add_rubric(doc, "The Deacon or a Priest says")
    add_celebrant_line(doc, "", prayers["confession_invitation"])

    # Confession text (bold, recited by people)
    text = " ".join(prayers["confession"])
    doc.add_paragraph(text, style="Body - People Recitation")
    add_spacer(doc)

    # Absolution
    add_rubric(doc, "The Priest says the Absolution.")
    abs_text = prayers["absolution"]
    p = doc.add_paragraph(style="Body")
    # Insert cross at "forgive you all your sins"
    if "forgive you" in abs_text:
        before, after = abs_text.split("forgive you", 1)
        p.add_run(before)
        add_cross_symbol(p)
        run = p.add_run(" forgive you" + after)
    else:
        p.add_run(abs_text)


def _add_decalogue(doc: Document, prayers: dict):
    """Add the Decalogue (Ten Commandments) for Lent 1."""
    add_body(doc, prayers["decalogue_intro"])
    for item in prayers["decalogue"]:
        add_celebrant_line(doc, "", item["commandment"])
        p = doc.add_paragraph(style="Body - People")
        p.add_run(item["response"])
    add_spacer(doc)


def _add_reading(doc: Document, reference: str, reading):
    """Add a scripture reading with responses."""
    add_heading2(doc, f"The Scriptures: {reference}")

    # Add the reading text
    if hasattr(reading, "paragraphs"):
        for i, para in enumerate(reading.paragraphs):
            add_scripture_text(doc, para, indent=(i > 0))
        if reading.has_poetry:
            for line in reading.poetry_lines:
                add_scripture_text(doc, line, style="Reading (Poetry)")
    else:
        # Fallback: reading is a plain string
        add_scripture_text(doc, str(reading))

    add_spacer(doc)
    add_rubric(doc, "After the reading, the Reader will say")
    add_celebrant_line(doc, "", "The Word of the Lord.")
    add_people_line(doc, "People", "Thanks be to God.")


def _add_psalm(doc: Document, reference: str, rubric: str, lines):
    """Add a psalm or canticle.

    Each entry in *lines* is a single verse formatted as:
      "first half text\\n\\tsecond half line 1\\n\\tsecond half line 2"
    We render each verse as one "Psalm"-styled paragraph, using
    line breaks (not new paragraphs) within a verse so the hanging
    indent applies to the second-half lines.
    """
    p = doc.add_paragraph(style="Heading 2")
    run = p.add_run(reference)
    run.italic = True

    if rubric:
        add_rubric(doc, rubric)

    if isinstance(lines, list):
        for verse_text in lines:
            p = doc.add_paragraph(style="Psalm")
            if "\n" in verse_text:
                sub_lines = verse_text.split("\n")
                for i, sub in enumerate(sub_lines):
                    if i > 0:
                        run = p.add_run()
                        run.add_break()
                    p.add_run(sub)
            else:
                p.add_run(verse_text)
    elif hasattr(lines, "paragraphs"):
        for para in lines.paragraphs:
            doc.add_paragraph(para, style="Psalm")


def _add_gospel(doc: Document, reference: str, book: str, reading):
    """Add the Gospel reading with announcement and response."""
    add_introductory_rubric(doc, "Remain standing.")
    p = doc.add_paragraph(style="Heading 2")
    run = p.add_run(f"The Gospel: {reference}")
    run.italic = True

    add_rubric(doc, "The Deacon or a Priest reads the Gospel, first saying")

    # Announcement
    p = doc.add_paragraph(style="Body - Celebrant")
    run = p.add_run(f"The Holy Gospel of our Lord Jesus Christ according to {book}.")
    run.italic = True
    add_people_line(doc, "People", "Glory to you, Lord Christ.")
    add_spacer(doc)

    # Gospel text
    if hasattr(reading, "paragraphs"):
        for i, para in enumerate(reading.paragraphs):
            add_scripture_text(doc, para, indent=(i > 0))
    else:
        add_scripture_text(doc, str(reading))

    add_spacer(doc)
    add_rubric(doc, "After the Gospel, the Reader will say")
    add_celebrant_line(doc, "", "The Gospel of the Lord.")
    add_people_line(doc, "People", "Praise to you, Lord Christ.")


def _add_nicene_creed(doc: Document, prayers: dict):
    """Add the Nicene Creed in the three-article format."""
    creed_lines = prayers["nicene_creed"]
    # Group into three articles (separated by blank lines)
    articles = []
    current = []
    for line in creed_lines:
        if line == "":
            if current:
                articles.append(" ".join(current))
                current = []
        else:
            current.append(line.strip())
    if current:
        articles.append(" ".join(current))

    for article in articles:
        doc.add_paragraph(article, style="Body - People Recitation (Creed)")


def _add_pop(doc: Document, elements: list[dict]):
    """Add Prayers of the People elements."""
    for elem in elements:
        etype = elem.get("type", "leader")
        text = elem.get("text", "")

        if etype == "leader":
            add_body(doc, text)
        elif etype == "people":
            p = doc.add_paragraph(style="Body")
            run = p.add_run(text)
            run.bold = True
        elif etype == "rubric":
            add_rubric(doc, text)
        elif etype == "both":
            add_body(doc, elem.get("leader_text", ""))
            p = doc.add_paragraph(style="Body")
            run = p.add_run(elem.get("people_text", ""))
            run.bold = True

        # Add spacer between petitions
        if etype in ("people", "both"):
            add_spacer(doc)


def _add_advent_wreath(doc: Document, data: dict):
    """Add the Advent wreath lighting with O Come Emmanuel."""
    hymn_ref = data.get("advent_hymnal_ref", "#56 (Hymnal 1982)")
    p = doc.add_paragraph(style="Body")
    p.add_run("O come, O come, Emmanuel")
    run = p.add_run(f"\t{hymn_ref}")
    run.italic = True

    # Verse 1 (always)
    advent_v1 = [
        "O come, O come, Emmanuel,",
        "and ransom captive Israel,",
        "that mourns in lonely exile here",
        "until the Son of God appear.",
    ]
    for line in advent_v1:
        doc.add_paragraph(line, style="Body - Lyrics")
    add_spacer(doc)

    # Chorus
    chorus = ["Rejoice! Rejoice! Emmanuel", "shall come to thee, O Israel!"]
    for line in chorus:
        p = doc.add_paragraph(style="Body - Lyrics")
        run = p.add_run(line)
        run.italic = True

    # Additional O Antiphon verse if provided
    advent_verse = data.get("advent_wreath_verse")
    if advent_verse and isinstance(advent_verse, list):
        add_spacer(doc)
        for line in advent_verse:
            doc.add_paragraph(line, style="Body - Lyrics")
        add_spacer(doc)
        for line in chorus:
            p = doc.add_paragraph(style="Body - Lyrics")
            run = p.add_run(line)
            run.italic = True


def _add_song_smart(doc: Document, song_data: dict | None):
    """Add a song, choosing two-column layout for 3+ verse songs."""
    if not song_data:
        add_body(doc, "[Song lyrics not found]")
        return

    # Count verses (non-chorus sections)
    verses = [s for s in song_data.get("sections", []) if s["type"] == "verse"]
    max_line_len = max(
        (len(line) for s in song_data.get("sections", []) for line in s["lines"]),
        default=0
    )

    # Two-column if 3+ verses and lines aren't too long for half-width
    # ~42 chars is roughly half of 7" page at 12pt Gill Sans Light
    if len(verses) >= 3 and max_line_len <= 48:
        add_song_two_column(doc, song_data)
    else:
        add_song(doc, song_data)


def _add_body_with_amen(doc: Document, text: str):
    """Add body text, making the final 'Amen.' bold."""
    if text.rstrip().endswith("Amen."):
        body = text.rstrip()[:-5]
        p = doc.add_paragraph(style="Body")
        p.add_run(body)
        run = p.add_run("Amen.")
        run.bold = True
    else:
        add_body(doc, text)


def _add_celebrant_with_cross(doc: Document, label: str, text: str):
    """Add a celebrant line, rendering {cross} / ✠ as a bold-italic cross symbol."""
    p = doc.add_paragraph(style="Body - Celebrant")
    if label:
        p.add_run(label)
        p.add_run("\t")

    if CROSS_SYMBOL in text:
        parts = text.split(CROSS_SYMBOL)
        p.add_run(parts[0])
        add_cross_symbol(p)
        if len(parts) > 1:
            p.add_run(parts[1])
    else:
        p.add_run(text)
