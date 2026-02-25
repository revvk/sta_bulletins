"""
Main document builder that orchestrates bulletin assembly.

Coordinates data fetching, rules computation, and section assembly
to produce a complete .docx bulletin.
"""

import re
from pathlib import Path
from datetime import date, datetime

from bulletin.config import CHURCH_NAME, GIVING_URL
from bulletin.document.styles import create_document
from bulletin.document.sections.cover import add_cover
from bulletin.document.sections.word_of_god import add_word_of_god
from bulletin.document.sections.holy_communion import add_holy_communion
from bulletin.document.sections.back_page import add_back_page
from bulletin.logic.rules import get_seasonal_rules, get_dismissal_text
from bulletin.data.loader import (
    load_common_prayers, load_pop_forms, load_blessings,
    get_proper_preface_text, get_preface_option_labels,
    load_staff,
)


class BulletinBuilder:
    """Builds a bulletin .docx from all data sources.

    Usage:
        builder = BulletinBuilder(target_date, sheet_data, music_data,
                                  scripture_data, song_lookup_fn)
        builder.resolve_interactive()  # prompts user for choices
        doc = builder.build()
        doc.save("output/bulletin.docx")
    """

    def __init__(self, target_date: date, sheet_data, music_data,
                 scripture_readings: dict, song_lookup_fn,
                 parish_ministries: str):
        """
        Args:
            target_date: The Sunday date.
            sheet_data: BulletinData from google_sheet module.
            music_data: Music9amEntry from music_9am module (or None).
            scripture_readings: Dict of {ref: ScriptureReading} from scripture module.
            song_lookup_fn: Callable(identifier, service) -> song_data dict or None.
            parish_ministries: Formatted ministry string for POP.
        """
        self.target_date = target_date
        self.sheet = sheet_data
        self.music = music_data
        self.scripture = scripture_readings
        self.song_lookup = song_lookup_fn
        self.parish_ministries = parish_ministries

        # Derived data
        self.schedule = sheet_data.schedule
        self.clergy = sheet_data.clergy

        # Compute seasonal rules
        self.rules = get_seasonal_rules(
            title=self.schedule.title,
            color=self.schedule.color,
            notes=self.schedule.notes or "",
            pop_form=self.schedule.pop_form or "",
        )

        # Data that may need user input
        self.proper_preface_text = ""
        self.advent_wreath_verse = None
        self.penitential_sentence = None
        self.blessing_text = ""
        self.eucharistic_prayer = "A"
        self.post_communion_prayer = "a"
        self._missing_songs = []

    def resolve_all(self, prompt_fn=None):
        """Resolve all data that might need user input.

        Args:
            prompt_fn: Callable(question, options) -> answer.
                       If None, makes best-guess choices silently.
        """
        self._resolve_eucharistic_prayer()
        self._resolve_proper_preface(prompt_fn)
        self._resolve_penitential_sentence(prompt_fn)
        self._resolve_advent_wreath(prompt_fn)
        self._resolve_blessing()
        self._resolve_songs(prompt_fn)

    def build(self) -> "Document":
        """Build and return the complete document."""
        doc = create_document()

        # Format the date nicely
        date_str = self.target_date.strftime("%B %-d, %Y")
        service_time = "9 am"

        # Cover page
        add_cover(doc, date_str, service_time,
                  self.schedule.title,
                  giving_url=GIVING_URL)

        # Word of God
        wog_data = self._prepare_word_of_god_data()
        add_word_of_god(doc, self.rules, wog_data)

        # Holy Communion
        hc_data = self._prepare_holy_communion_data()
        add_holy_communion(doc, self.rules, hc_data)

        # Back page
        add_back_page(doc)

        return doc

    # ------------------------------------------------------------------
    # Resolution methods
    # ------------------------------------------------------------------

    def _resolve_eucharistic_prayer(self):
        """Determine which Eucharistic Prayer to use from the sheet."""
        ep = self.schedule.eucharistic_prayer or "A"
        self.eucharistic_prayer = ep.strip().upper()

    def _resolve_proper_preface(self, prompt_fn):
        """Resolve the proper preface text."""
        key = self.rules.proper_preface_key
        options = self.rules.proper_preface_options

        if not options:
            # Single preface for this season
            self.proper_preface_text = get_proper_preface_text(key)
        elif prompt_fn and self.rules.prompt_preface:
            # Multiple options — ask the user
            labels = get_preface_option_labels(key)
            question = f"Proper Preface for '{self.schedule.title}':"
            option_strs = [f"  {i+1}. {label}" for i, (_, label) in enumerate(labels)]
            answer = prompt_fn(question, [label for _, label in labels])
            # Match answer to option key
            for opt_key, label in labels:
                if label == answer or answer == opt_key:
                    self.proper_preface_text = get_proper_preface_text(key, opt_key)
                    return
            # Default to first option
            self.proper_preface_text = get_proper_preface_text(key, options[0])
        else:
            # No prompt function — default to first option
            self.proper_preface_text = get_proper_preface_text(key, options[0])

    def _resolve_penitential_sentence(self, prompt_fn):
        """For Lent 2-5, pick a Penitential Order scripture sentence."""
        if not self.rules.use_penitential_order or self.rules.use_decalogue:
            return

        prayers = load_common_prayers()
        sentences = prayers.get("penitential_sentences", [])

        if prompt_fn and sentences:
            question = "Penitential Order scripture sentence:"
            options = [f'{s["reference"]}: {s["text"][:60]}...' for s in sentences]
            answer = prompt_fn(question, options)
            for s in sentences:
                if s["reference"] in answer:
                    self.penitential_sentence = s["text"]
                    return

        # Default to first sentence
        if sentences:
            self.penitential_sentence = sentences[0]["text"]

    def _resolve_advent_wreath(self, prompt_fn):
        """For Advent, resolve the O Antiphon verse for the wreath lighting."""
        if not self.rules.is_advent:
            return

        if prompt_fn:
            answer = prompt_fn(
                "Advent wreath: Enter the O Antiphon verse lines "
                "(one per line, or leave blank for verse 1 only):",
                []
            )
            if answer:
                self.advent_wreath_verse = answer.strip().split("\n")

    def _resolve_blessing(self):
        """Look up the seasonal blessing from BOS data."""
        blessings = load_blessings()
        season = self.rules.season

        # Map season to blessing key
        season_map = {
            "advent": "advent",
            "christmas": "christmas",
            "epiphany": "epiphany",
            "lent": "lent",
            "easter": "easter",
            "pentecost_day": "pentecost",
            "ordinary": None,
        }

        blessing_key = season_map.get(season)
        if blessing_key and blessing_key in blessings:
            blessing = blessings[blessing_key]
            if isinstance(blessing, dict):
                self.blessing_text = blessing.get("text", "")
            else:
                self.blessing_text = str(blessing)

    def _resolve_songs(self, prompt_fn):
        """Look up all song lyrics from music data."""
        self._missing_songs = []

        if self.music:
            # Map music slots to bulletin slots
            for slot in self.music.slots:
                song = self.song_lookup(slot.song_title, "9 am")
                if not song and slot.song_title:
                    self._missing_songs.append(
                        f"{slot.service_part}: {slot.song_title}")

        if self._missing_songs and prompt_fn:
            prompt_fn(
                f"Warning: Could not find lyrics for:\n" +
                "\n".join(f"  - {s}" for s in self._missing_songs),
                ["Continue anyway"]
            )

    # ------------------------------------------------------------------
    # Data preparation for section modules
    # ------------------------------------------------------------------

    def _prepare_word_of_god_data(self) -> dict:
        """Prepare the data dict for add_word_of_god."""
        # Look up songs from music data
        processional = self._lookup_slot("Processional")
        song_of_praise = self._lookup_slot("Song of Praise")
        sequence = self._lookup_slot("Sequence")

        # Scripture readings
        # Note: fetch_readings returns {label: ScriptureReading} using
        # "reading" and "gospel" as keys (set up in generate.py)
        reading_1_ref = self.schedule.reading or ""
        reading_1 = self.scripture.get("reading")

        psalm_raw = self.schedule.psalm or ""
        # Clean: sheet may contain "Psalm 121\nunison" — extract just the reference
        psalm_ref = re.split(r"[\n\r]+", psalm_raw)[0].strip()
        psalm_ref = re.sub(
            r"\s+(responsively|unison|in unison|antiphonally).*$",
            "", psalm_ref, flags=re.IGNORECASE,
        ).strip()
        psalm_text: list[str] = []
        if psalm_ref:
            try:
                from bulletin.sources.psalms import get_psalm
                psalm_selection = get_psalm(psalm_ref)
                psalm_text = psalm_selection.to_lines()
            except (ValueError, Exception) as e:
                print(f"  Warning: Could not look up psalm: {e}")
                psalm_text = []

        gospel_ref = self.schedule.gospel or ""
        gospel = self.scripture.get("gospel")

        # Extract gospel book name from reference
        gospel_book = gospel_ref.split()[0] if gospel_ref else ""

        # Preacher from clergy rota
        preacher = ""
        if self.clergy:
            preacher = self.clergy.preacher_9am or ""

        # POP elements
        pop_elements = self._prepare_pop_elements()

        return {
            "processional": processional,
            "song_of_praise": song_of_praise,
            "sequence_hymn": sequence,
            "collect_text": "[Collect of the Day]",  # Collect is read live, not printed
            "reading_1_ref": reading_1_ref,
            "reading_1_text": reading_1 or reading_1_ref,
            "psalm_ref": psalm_ref,
            "psalm_rubric": self._get_psalm_rubric(),
            "psalm_text": psalm_text or [],
            "gospel_ref": gospel_ref,
            "gospel_book": gospel_book,
            "gospel_text": gospel or gospel_ref,
            "preacher": preacher,
            "pop_elements": pop_elements,
            "pop_concluding_rubric": "The Celebrant concludes with a suitable Collect.",
            "advent_wreath_verse": self.advent_wreath_verse,
            "advent_hymnal_ref": "#56 (Hymnal 1982)",
            "penitential_sentence": self.penitential_sentence,
        }

    def _prepare_holy_communion_data(self) -> dict:
        """Prepare the data dict for add_holy_communion."""
        offertory = self._lookup_slot("Offertory")
        sanctus = self._lookup_slot("Sanctus")
        closing = self._lookup_slot("Recessional")
        fraction = self._lookup_slot("Fraction") if self.rules.use_fraction_anthem else None

        # Communion songs
        comm_songs = []
        for slot_name in ["Communion 1", "Communion 2", "Communion 3"]:
            song = self._lookup_slot(slot_name)
            if song:
                comm_songs.append(song)

        # Postlude: reuse various song snippets (this is a medley typically)
        postlude_songs = []

        # Dismissal text
        dismissal_num = self.schedule.dismissal or "3"
        deacon_text, people_text = get_dismissal_text(
            dismissal_num, self.rules.dismissal_has_alleluia)

        return {
            "offertory_song": offertory,
            "sanctus_song": sanctus,
            "communion_songs": comm_songs,
            "closing_hymn": closing,
            "postlude_songs": postlude_songs,
            "eucharistic_prayer": self.eucharistic_prayer,
            "proper_preface_text": self.proper_preface_text,
            "fraction_song": fraction,
            "blessing_text": self.blessing_text,
            "dismissal_deacon": deacon_text,
            "dismissal_people": people_text,
            "post_communion_prayer": self.post_communion_prayer,
            "include_lev": True,
        }

    def _lookup_slot(self, slot_name: str) -> dict | None:
        """Look up a song from the music data for a given service slot."""
        if not self.music:
            return None

        for slot in self.music.slots:
            if slot.service_part and slot.service_part.lower() == slot_name.lower():
                if slot.song_title:
                    return self.song_lookup(slot.song_title, "9 am")
        return None

    def _get_psalm_rubric(self) -> str:
        """Determine the psalm rubric from the sheet's psalm field."""
        psalm_field = (self.schedule.psalm or "").lower()
        if "responsiv" in psalm_field:
            return "Read responsively by whole verse."
        if "antiphon" in psalm_field:
            return ""
        # "unison" or default
        return "Read in unison."

    def _prepare_pop_elements(self) -> list[dict]:
        """Build POP elements from the appropriate form."""
        pop_forms = load_pop_forms()
        staff = load_staff()
        liturgical_names = staff.get("liturgical_names", {})

        # Determine which form to use
        form_key = self._get_pop_form_key()
        form = pop_forms.get(form_key)

        if not form:
            return [{"type": "leader",
                     "text": f"[Prayers of the People form '{form_key}' not found]"}]

        elements = form.get("elements", [])

        # Substitute placeholders
        subs = {
            "{presiding_bishop}": liturgical_names.get("presiding_bishop", "N."),
            "{bishop}": liturgical_names.get("bishop", "N."),
            "{bishop_coadjutor}": liturgical_names.get("bishop_coadjutor", ""),
            "{clergy_names}": ", ".join(liturgical_names.get("clergy_first_names", [])),
            "{president}": liturgical_names.get("president", "N."),
            "{governor}": liturgical_names.get("governor", "N."),
            "{mayor}": liturgical_names.get("mayor", "N."),
            "{parish_ministries}": self.parish_ministries,
            "{departed}": "",  # Left blank; filled in manually
        }

        result = []
        for elem in elements:
            new_elem = dict(elem)
            for field in ("text", "leader_text", "people_text"):
                if field in new_elem:
                    for placeholder, value in subs.items():
                        new_elem[field] = new_elem[field].replace(placeholder, value)
            result.append(new_elem)

        return result

    def _get_pop_form_key(self) -> str:
        """Map the sheet's POP form designation to a YAML key."""
        form = (self.schedule.pop_form or "").strip()

        # Check for BCP Form VI w/ confession (handled in YAML under form_VI)
        if self.rules.pop_use_bcp_form_vi:
            return "form_VI"

        # Check for Advent forms
        if self.rules.is_advent:
            title_lower = self.schedule.title.lower()
            if "first" in title_lower or "1" in title_lower.replace(" ", ""):
                return "advent_I"
            elif "second" in title_lower or "2" in title_lower.replace(" ", ""):
                return "advent_II"
            elif "third" in title_lower or "3" in title_lower.replace(" ", ""):
                return "advent_III"
            elif "fourth" in title_lower or "4" in title_lower.replace(" ", ""):
                return "advent_IV"

        # Standard forms: "I", "II", "III", "IV", "V", "VI"
        roman_map = {
            "I": "form_I", "II": "form_II", "III": "form_III",
            "IV": "form_IV", "V": "form_V", "VI": "form_VI",
            "1": "form_I", "2": "form_II", "3": "form_III",
            "4": "form_IV", "5": "form_V", "6": "form_VI",
        }

        # Extract just the Roman numeral from the form string
        form_clean = form.split("(")[0].strip()
        return roman_map.get(form_clean, "form_I")
