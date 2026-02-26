"""
The Holy Communion section of the bulletin.

Covers everything from the Offertory through the Dismissal:
  - Offertory
  - Offering Music
  - Doxology
  - The Great Thanksgiving (Sursum Corda + Proper Preface + Sanctus + Eucharistic Prayer)
  - Lord's Prayer
  - Breaking of the Bread
  - Invitation
  - Holy Communion (rubric)
  - Communion Music
  - Closing Prayer
  - Prayer for Lay Eucharistic Visitor (optional)
  - Blessing / Prayer over the People
  - Closing Hymn
  - Dismissal
  - Postlude
"""

from docx import Document

from bulletin.config import CROSS_SYMBOL
from bulletin.data.loader import (
    load_common_prayers, load_eucharistic_prayers, load_blessings,
    get_proper_preface_text,
)
from bulletin.document.formatting import (
    add_spacer, add_heading, add_heading2, add_rubric,
    add_introductory_rubric, add_body, add_celebrant_line,
    add_people_line, add_cross_symbol, add_song, add_song_two_column,
    add_hymn_header,
)
from bulletin.logic.rules import SeasonalRules


def add_holy_communion(doc: Document, rules: SeasonalRules, data: dict):
    """Add the entire Holy Communion section.

    Args:
        doc: The Document.
        rules: SeasonalRules for this service.
        data: Dict with keys:
            - offertory_song: song data dict (or None)
            - sanctus_song: song data dict (or None) — the Sanctus setting
            - communion_songs: list of song data dicts
            - closing_hymn: song data dict (or None)
            - postlude_songs: list of song data dicts (medley for postlude)
            - eucharistic_prayer: "A", "B", or "C"
            - proper_preface_text: str (already resolved)
            - fraction_song: song data dict | None (Lent: Agnus Dei)
            - blessing_text: str | None
            - dismissal_deacon: str
            - dismissal_people: str
            - post_communion_prayer: "a" or "b"
            - include_lev: bool (Lay Eucharistic Visitor prayer)
    """
    prayers = load_common_prayers()
    ep_data = load_eucharistic_prayers()

    # Section heading
    add_heading(doc, "")
    add_heading(doc, "The Holy Communion")
    add_heading(doc, "")

    # --- Offertory ---
    add_spacer(doc)
    add_rubric(doc, "Be seated.")
    add_heading2(doc, "Offertory")
    add_rubric(doc, "Please place your offerings and Connection Cards in the "
               "offering plates as they are passed. You can also easily give "
               "online at standrewsmckinney.org/give.")
    add_spacer(doc)

    # Offering Music
    add_heading2(doc, "Offering Music")
    _add_song_smart(doc, data.get("offertory_song"))

    # Doxology
    add_spacer(doc)
    add_introductory_rubric(doc, "Please stand.")
    p = doc.add_paragraph(style="Heading 2")
    p.add_run("Doxology")
    run = p.add_run("\tOld 100th")
    run.italic = True
    _DOXOLOGY = [
        "Praise God, from Whom all blessings flow;",
        "Praise Him, all creatures here below;",
        "Praise Him above, ye heavenly host:",
        "Praise Father, Son, and Holy Ghost.",
    ]
    for line in _DOXOLOGY:
        doc.add_paragraph(line, style="Body - Lyrics")

    # --- The Great Thanksgiving ---
    add_spacer(doc)
    add_rubric(doc, "Remain standing.")
    add_heading2(doc, "The Great Thanksgiving")

    # Sursum Corda
    for exchange in ep_data["sursum_corda"]:
        p = doc.add_paragraph(style="Body - Celebrant")
        run = p.add_run("Celebrant")
        p.add_run("\t")
        run2 = p.add_run(exchange["celebrant"])
        run2.bold = True
        add_people_line(doc, "People", exchange["people"])
    add_spacer(doc)

    # Eucharistic Prayer
    prayer_key = data.get("eucharistic_prayer", "A").upper()
    if prayer_key == "C":
        _add_prayer_c(doc, ep_data, data, prayers)
    else:
        _add_prayer_a_or_b(doc, ep_data, data, prayers, prayer_key)

    # --- Lord's Prayer ---
    add_rubric(doc, "Please stand and, as you are comfortable, join hands "
               "with those around you.")
    add_body(doc, prayers["lords_prayer_intro"]["option_1"])
    text = " ".join(line.strip() for line in prayers["lords_prayer"])
    doc.add_paragraph(text, style="Body - People Recitation")

    # --- Breaking of the Bread ---
    add_spacer(doc)
    add_heading2(doc, "Breaking of the Bread")
    if rules.use_fraction_anthem:
        # Lent: sung fraction anthem (Agnus Dei)
        fraction_song = data.get("fraction_song")
        if fraction_song:
            _add_song_smart(doc, fraction_song)
        else:
            # Default Agnus Dei text
            for line in [
                "Lamb of God, you take away the sins of the world: have mercy on us.",
                "Lamb of God, you take away the sins of the world: have mercy on us.",
                "Lamb of God, you take away the sins of the world: grant us peace.",
            ]:
                p = doc.add_paragraph(style="Body - Celebrant")
                run = p.add_run(line)
                run.bold = True
    else:
        add_celebrant_line(doc, "Celebrant", rules.fraction_celebrant)
        add_people_line(doc, "People", rules.fraction_people)

    # --- Invitation ---
    add_spacer(doc)
    add_heading2(doc, "Invitation")
    add_rubric(doc, "Facing the people, the Celebrant says the following Invitation")
    add_body(doc, prayers["invitation_to_communion"] + " " +
             prayers["invitation_addition"])

    # --- Holy Communion ---
    add_spacer(doc)
    add_heading2(doc, "Holy Communion")
    add_rubric(doc, "All baptized Christians are invited to the altar rail to "
               "receive Holy Communion. If you are not going to receive, you "
               "are still welcome at the rail for a blessing; just cross your "
               "arms over your chest.")

    # --- Communion Music ---
    add_spacer(doc)
    add_heading2(doc, "Communion Music")
    comm_songs = data.get("communion_songs", [])
    for i, song in enumerate(comm_songs):
        if i > 0:
            add_spacer(doc)
        _add_song_smart(doc, song)

    # --- Closing Prayer ---
    add_spacer(doc)
    add_heading2(doc, "Closing Prayer")
    p = doc.add_paragraph(style="Body - Celebrant")
    run = p.add_run("Celebrant")
    p.add_run("\t")
    run2 = p.add_run("Let us pray.")
    run2.bold = True

    prayer_choice = data.get("post_communion_prayer", "a")
    if prayer_choice == "b":
        text = " ".join(prayers["post_communion_prayer_b"])
    else:
        text = " ".join(prayers["post_communion_prayer_a"])
    doc.add_paragraph(text, style="Body - People Recitation")

    # --- Prayer for Lay Eucharistic Visitor (optional) ---
    if data.get("include_lev", True):
        add_spacer(doc)
        add_heading2(doc, "Prayer for Lay Eucharistic Visitor")
        p = doc.add_paragraph(style="Body - Rubric")
        run = p.add_run("The Celebrant commissions the Lay Eucharistic Visitor, saying")
        run.italic = True
        p = doc.add_paragraph(style="Body - Celebrant")
        run = p.add_run(
            "In the name of this congregation, I send you forth bearing these "
            "holy gifts that those to whom you go may share with us in the "
            "communion of Christ's body and blood."
        )
        add_people_line(doc, "People",
                        "We who are many are one body because we all share "
                        "one bread and one cup.")

    # --- Blessing / Prayer over the People ---
    add_spacer(doc)
    add_heading2(doc, rules.blessing_label)
    if rules.use_prayer_over_people:
        add_rubric(doc, "The Deacon or a Priest says")
        add_celebrant_line(doc, "", "Bow down before the Lord.")
        add_spacer(doc)
        add_rubric(doc, "The people bow their heads and the Celebrant says "
                   "the following prayer")
        blessing_text = data.get("blessing_text", "")
        if blessing_text:
            _add_body_with_amen(doc, blessing_text)
    else:
        blessing_text = data.get("blessing_text", "")
        if blessing_text:
            p = doc.add_paragraph(style="Body - Celebrant")
            run = p.add_run("Celebrant")
            p.add_run("\t")
            run2 = p.add_run(blessing_text)
        else:
            p = doc.add_paragraph(style="Body - Celebrant")
            run = p.add_run("Celebrant")
            p.add_run("\t")
            run2 = p.add_run(
                "\u2026 and the blessing of God Almighty, "
                f"{CROSS_SYMBOL} the Father, the Son, and the Holy Spirit, "
                "be among you, and remain with you always."
            )

    # --- Closing Hymn ---
    add_spacer(doc)
    add_heading2(doc, "Closing Hymn")
    _add_song_smart(doc, data.get("closing_hymn"))

    # --- Dismissal ---
    add_spacer(doc)
    add_heading2(doc, "Dismissal")
    add_rubric(doc, "The Deacon or a Priest dismisses with these words")
    add_celebrant_line(doc, "", data.get("dismissal_deacon", ""))
    add_people_line(doc, "People", data.get("dismissal_people", ""))

    # --- Postlude ---
    add_spacer(doc)
    add_heading2(doc, "Postlude")
    postlude_songs = data.get("postlude_songs", [])
    for song in postlude_songs:
        _add_song_smart(doc, song)


def _add_prayer_a_or_b(doc: Document, ep_data: dict, data: dict,
                       prayers: dict, prayer_key: str):
    """Add Eucharistic Prayer A or B."""
    key = f"prayer_{prayer_key.lower()}"
    prayer = ep_data[key]

    # Preface opening + proper preface + Sanctus transition
    add_body(doc, ep_data["preface_opening"])
    add_spacer(doc)

    # Proper preface
    preface_text = data.get("proper_preface_text", "")
    if preface_text:
        add_body(doc, preface_text)
        add_spacer(doc)

    add_body(doc, ep_data["sanctus_transition"])
    add_spacer(doc)

    # Sanctus hymn
    sanctus_song = data.get("sanctus_song")
    if sanctus_song:
        _add_song_smart(doc, sanctus_song)
    else:
        # Text version of Sanctus
        for line in prayers["sanctus"]:
            doc.add_paragraph(line, style="Body - Lyrics")

    # Kneeling rubric
    add_rubric(doc, "Please kneel or remain standing. The Celebrant continues")

    # Post-Sanctus
    add_body(doc, prayer["post_sanctus"])
    add_spacer(doc)

    # Institution narrative (Prayer A has institution_1, B does not)
    if "institution_1" in prayer:
        add_body(doc, prayer["institution_1"])
        add_spacer(doc)

    # Words of institution - bread
    _add_institution_words(doc, prayer["institution_bread"])
    add_spacer(doc)

    # Words of institution - cup
    _add_institution_words(doc, prayer["institution_cup"])
    add_spacer(doc)

    # Memorial acclamation
    add_body(doc, prayer["memorial_intro"])
    p = doc.add_paragraph(style="Body")
    for i, line in enumerate(prayer["memorial_acclamation"]):
        if i > 0:
            run = p.add_run()
            run.add_break()
        run = p.add_run(line)
        run.bold = True
    add_spacer(doc)

    # Epiclesis
    add_rubric(doc, "The Celebrant continues")
    add_body(doc, prayer["epiclesis_1"] if "epiclesis_1" in prayer else prayer.get("epiclesis", ""))
    add_spacer(doc)
    add_body(doc, prayer["epiclesis_2"])
    add_spacer(doc)

    # Doxology
    _add_doxology_amen(doc, prayer)


def _add_prayer_c(doc: Document, ep_data: dict, data: dict, prayers: dict):
    """Add Eucharistic Prayer C (responsive format).

    Prayer C has a unique structure different from A/B — it's fully
    responsive with Celebrant/People exchanges throughout.
    """
    # Prayer C text — we output the full responsive text
    # The preface is woven into the prayer itself, not separate
    pc = [
        ("celebrant", "God of all power, Ruler of the Universe, you are worthy of glory and praise."),
        ("people", "Glory to you for ever and ever."),
        ("celebrant", "At your command all things came to be: the vast expanse of interstellar space, galaxies, suns, the planets in their courses, and this fragile earth, our island home."),
        ("people", "By your will they were created and have their being."),
        ("celebrant", "From the primal elements you brought forth the human race, and blessed us with memory, reason, and skill. You made us the rulers of creation. But we turned against you, and betrayed your trust; and we turned against one another."),
        ("people", "Have mercy, Lord, for we are sinners in your sight."),
        ("celebrant", "Again and again, you called us to return. Through prophets and sages you revealed your righteous Law. And in the fullness of time you sent your only Son, born of a woman, to fulfill your Law, to open for us the way of freedom and peace."),
        ("people", "By his blood, he reconciled us. By his wounds, we are healed."),
        ("celebrant", "And therefore we praise you, joining with the heavenly chorus, with prophets, apostles, and martyrs, and with all those in every generation who have looked to you in hope, to proclaim with them your glory, in their unending hymn:"),
    ]

    for role, text in pc:
        if role == "celebrant":
            add_body(doc, text)
        else:
            p = doc.add_paragraph(style="Body")
            run = p.add_run(text)
            run.bold = True
        add_spacer(doc)

    # Sanctus
    sanctus_song = data.get("sanctus_song")
    if sanctus_song:
        _add_song_smart(doc, sanctus_song)
    else:
        for line in prayers["sanctus"]:
            doc.add_paragraph(line, style="Body - Lyrics")

    # Post-Sanctus continuation
    add_rubric(doc, "Please kneel or remain standing. The Celebrant continues")

    pc2 = [
        ("celebrant", "And so, Father, we who have been redeemed by him, and made a new people by water and the Spirit, now bring before you these gifts. Sanctify them by your Holy Spirit to be the Body and Blood of Jesus Christ our Lord."),
    ]
    for role, text in pc2:
        add_body(doc, text)
    add_spacer(doc)

    # Institution - bread
    _add_institution_words(doc,
        "On the night he was betrayed he took bread, said the blessing, "
        "broke the bread, and gave it to his friends, and said, "
        '"Take, eat: This is my Body, which is given for you. '
        'Do this for the remembrance of me."')
    add_spacer(doc)

    # Institution - cup
    _add_institution_words(doc,
        "After supper, he took the cup of wine, gave thanks, and said, "
        '"Drink this, all of you: This is my Blood of the new Covenant, '
        "which is shed for you and for many for the forgiveness of sins. "
        'Whenever you drink it, do this for the remembrance of me."')
    add_spacer(doc)

    # Memorial acclamation
    pc3 = [
        ("celebrant", "Remembering now his work of redemption, and offering to you this sacrifice of thanksgiving,"),
        ("people", "We celebrate his death and resurrection, as we await the day of his coming."),
    ]
    for role, text in pc3:
        if role == "celebrant":
            add_body(doc, text)
        else:
            p = doc.add_paragraph(style="Body")
            run = p.add_run(text)
            run.bold = True
    add_spacer(doc)

    # Epiclesis
    pc4 = [
        ("celebrant", "Lord God of our Fathers; God of Abraham, Isaac, and Jacob; God and Father of our Lord Jesus Christ: Open our eyes to see your hand at work in the world about us. Deliver us from the presumption of coming to this Table for solace only, and not for strength; for pardon only, and not for renewal. Let the grace of this Holy Communion make us one body, one spirit in Christ, that we may worthily serve the world in his name."),
        ("people", "Risen Lord, be known to us in the breaking of the Bread."),
    ]
    for role, text in pc4:
        if role == "celebrant":
            add_body(doc, text)
        else:
            p = doc.add_paragraph(style="Body")
            run = p.add_run(text)
            run.bold = True
    add_spacer(doc)

    # Doxology
    p = doc.add_paragraph(style="Body")
    p.add_run(
        "Accept these prayers and praises, Father, through Jesus Christ our "
        "great High Priest, to whom, with you and the Holy Spirit, your "
        "Church gives honor, glory, and worship, from generation to generation. "
    )
    run = p.add_run("AMEN.")
    run.bold = True


def _add_institution_words(doc: Document, text: str):
    """Add institution narrative with bold-italic formatting for the actual words of institution."""
    # The words of institution (inside quotes) are formatted specially
    p = doc.add_paragraph(style="Body")
    # Find the quote marks
    if '"' in text:
        before_quote = text[:text.index('"')]
        from_quote = text[text.index('"'):]
        p.add_run(before_quote)
        run = p.add_run(from_quote)
        run.bold = True
        run.italic = True
    else:
        p.add_run(text)


def _add_doxology_amen(doc: Document, prayer: dict):
    """Add the prayer doxology with bold AMEN."""
    p = doc.add_paragraph(style="Body")
    p.add_run(prayer["doxology"] + " ")
    run = p.add_run(prayer["doxology_response"])
    run.bold = True


def _add_song_smart(doc: Document, song_data: dict | None):
    """Add a song, choosing two-column layout for 3+ verse songs."""
    if not song_data:
        add_body(doc, "[Song lyrics not found]")
        return

    verses = [s for s in song_data.get("sections", []) if s["type"] == "verse"]
    max_line_len = max(
        (len(line) for s in song_data.get("sections", []) for line in s["lines"]),
        default=0
    )

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
