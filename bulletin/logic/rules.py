"""
Liturgical rules engine for determining seasonal variations in the service.

This module encodes the Episcopal/Anglican liturgical rules that determine:
  - Whether to use the Penitential Order (Lent) or standard Word of God
  - Which Opening Acclamation to use
  - Whether to include "Alleluia" in Breaking of the Bread and Dismissal
  - Which form of Blessing vs. Prayer over the People
  - Collect for Purity inclusion
  - Position of Confession (before or after Prayers of the People)
  - POP form selection (custom vs. BCP Form VI w/ confession)
  - Advent Wreath lighting
  - Proper Preface selection (BCP pp.377-382)

All decisions are derived from the liturgical title, color, and notes
fields in the Google Sheet.
"""

from dataclasses import dataclass, field
from datetime import date
from typing import Optional


@dataclass
class SeasonalRules:
    """All seasonal rules for a given service."""

    # Structure
    use_penitential_order: bool     # Lent: Penitential Order replaces Word of God opening
    use_decalogue: bool             # Lent 1 only: long form with Decalogue (BCP p.350)
    confession_before_word: bool    # Lent (Penitential Order): Confession before Word of God
    no_confession_after_pop: bool   # Lent (Penitential Order): no Confession after POP
    include_collect_for_purity: bool  # Standard order: yes; Penitential Order: no

    # Opening Acclamation
    acclamation_celebrant: str
    acclamation_people: str

    # Song of Praise / Kyrie
    song_of_praise_label: str       # "Kyrie", "Song of Praise", or custom
    is_advent: bool                 # Advent: Wreath lighting + O Come Emmanuel

    # Breaking of the Bread
    use_fraction_anthem: bool       # Lent: sung fraction anthem / Agnus Dei
    fraction_celebrant: str
    fraction_people: str

    # Post-communion
    use_prayer_over_people: bool    # Lent: Prayer over the People replaces Blessing
    blessing_label: str             # "Blessing" or "Prayer over the People"

    # Dismissal -- uses the sheet's dismissal number directly
    dismissal_has_alleluia: bool    # Easter season: double Alleluia added

    # POP (Prayers of the People)
    pop_use_bcp_form_vi: bool       # True when notes say "(w/ confession)" on Form VI
    pop_has_confession: bool        # True when POP form includes confession (Form VI w/ confession)

    # Liturgical season name (for informational use)
    season: str

    # Proper Preface (BCP pp.377-382)
    # proper_preface_key: YAML key in proper_prefaces.yaml (e.g., "advent", "easter")
    # When "lords_day" or "lent", there are multiple options; prompt_preface is True.
    proper_preface_key: str = ""
    proper_preface_options: list[str] = field(default_factory=list)
    prompt_preface: bool = False    # True when user must choose among options


# ---------------------------------------------------------------------------
# Opening Acclamations
# ---------------------------------------------------------------------------

_ACCLAMATION_STANDARD = (
    "Blessed be God: {cross} Father, Son, and Holy Spirit.",
    "And blessed be his kingdom, now and for ever. Amen.",
)

_ACCLAMATION_LENT = (
    "Bless the Lord who forgives all our sins.",
    "His mercy endures forever.",
)

_ACCLAMATION_EASTER = (
    "Alleluia. Christ is risen.",
    "The Lord is risen indeed. Alleluia.",
)

# ---------------------------------------------------------------------------
# Breaking of the Bread
# ---------------------------------------------------------------------------

_FRACTION_ALLELUIA = (
    "Alleluia. Christ our Passover is sacrificed for us;",
    "Therefore let us keep the feast. Alleluia.",
)

_FRACTION_NO_ALLELUIA = (
    "Christ our Passover is sacrificed for us;",
    "Therefore let us keep the feast.",
)

# ---------------------------------------------------------------------------
# Dismissal texts (BCP p.366, numbered 1-4 in order)
# Sheet's Dismissal column maps directly: "1" -> first, "2" -> second, etc.
# ---------------------------------------------------------------------------

DISMISSALS = {
    "1": ("Let us go forth in the name of Christ.",
          "Thanks be to God."),
    "2": ("Go in peace to love and serve the Lord.",
          "Thanks be to God."),
    "3": ("Let us go forth into the world, rejoicing in the power of the Spirit.",
          "Thanks be to God."),
    "4": ("Let us bless the Lord.",
          "Thanks be to God."),
}


def get_dismissal_text(dismissal_num: str, has_alleluia: bool) -> tuple[str, str]:
    """Get the dismissal text for a given number, adding Alleluia if Easter season.

    Args:
        dismissal_num: "1", "2", "3", or "4" from the sheet
        has_alleluia: True during Easter season

    Returns:
        (deacon_text, people_response) tuple
    """
    base = DISMISSALS.get(dismissal_num, DISMISSALS["3"])
    if has_alleluia:
        return (
            base[0].rstrip(".") + ". Alleluia, alleluia.",
            base[1].rstrip(".") + ". Alleluia, alleluia.",
        )
    return base


# ---------------------------------------------------------------------------
# Season detection
# ---------------------------------------------------------------------------

def _detect_season(title: str, color: str, notes: str) -> str:
    """Detect the liturgical season from the title, color, and notes."""
    title_lower = title.lower()

    if "advent" in title_lower:
        return "advent"
    if "christmas" in title_lower or "christmastide" in title_lower:
        return "christmas"
    if "epiphany" in title_lower:
        return "epiphany"
    if "ash wednesday" in title_lower:
        return "lent"
    if "lent" in title_lower:
        return "lent"
    if "palm sunday" in title_lower or "passion" in title_lower:
        return "lent"  # Holy Week is liturgically part of Lent
    if "maundy" in title_lower or "good friday" in title_lower:
        return "lent"
    if "holy saturday" in title_lower:
        return "lent"
    if "easter" in title_lower:
        return "easter"
    if "ascension" in title_lower:
        return "easter"  # Ascension is in the Easter season
    if "pentecost" in title_lower and "after" not in title_lower:
        return "pentecost_day"
    if "trinity" in title_lower:
        return "ordinary"
    if "proper" in title_lower:
        return "ordinary"
    if "pentecost" in title_lower:
        return "ordinary"

    # Fall back to color
    if color.lower() in ("violet", "purple"):
        return "lent"  # Could be Advent but title check handles that first
    if color.lower() == "red":
        return "pentecost_day"

    return "ordinary"


def _is_in_easter_season(season: str) -> bool:
    """Easter Day through the Day of Pentecost."""
    return season in ("easter", "pentecost_day")


def _is_lent(season: str) -> bool:
    """Lent (including Holy Week)."""
    return season == "lent"


def _is_lent_1(title: str) -> bool:
    """Check if this is the First Sunday in Lent (uses Decalogue)."""
    title_lower = title.lower()
    return ("first sunday" in title_lower and "lent" in title_lower) or \
           "lent 1" in title_lower.replace(" ", "")


def _pop_has_confession(pop_form: str, notes: str) -> bool:
    """Check if the POP form includes a built-in confession.

    This happens when Form VI is used "(w/ confession)" -- the BCP form
    (pp. 392-393) has an optional confession section.
    """
    combined = f"{pop_form} {notes}".lower()
    return "w/ confession" in combined or "with confession" in combined


# ---------------------------------------------------------------------------
# Proper Preface selection (BCP pp.377-382)
# ---------------------------------------------------------------------------

def _is_holy_week(title: str) -> bool:
    """Check if this is Holy Week (Palm Sunday through Holy Saturday)."""
    t = title.lower()
    return any(kw in t for kw in ("palm sunday", "passion", "maundy",
                                   "good friday", "holy saturday",
                                   "holy week"))


def _is_ascension(title: str) -> bool:
    """Check if this is Ascension Day."""
    return "ascension" in title.lower()


def _is_trinity_sunday(title: str) -> bool:
    """Check if this is Trinity Sunday."""
    return "trinity" in title.lower()


def _get_proper_preface_key(title: str, season: str) -> tuple[str, list[str], bool]:
    """Determine the proper preface key, options list, and whether to prompt.

    Returns:
        (key, options, prompt) where:
          - key: YAML key in proper_prefaces.yaml
          - options: list of sub-keys when multiple choices exist
          - prompt: True if the user should be prompted to choose
    """
    # Special occasions override the season
    if _is_trinity_sunday(title):
        return ("trinity", [], False)
    if _is_ascension(title):
        return ("ascension", [], False)
    if _is_holy_week(title):
        return ("holy_week", [], False)

    # Season-based prefaces
    if season == "advent":
        return ("advent", [], False)
    if season == "christmas":
        return ("incarnation", [], False)
    if season == "epiphany":
        return ("epiphany", [], False)
    if season == "lent":
        return ("lent", ["option_1", "option_2"], True)
    if season == "easter":
        return ("easter", [], False)
    if season == "pentecost_day":
        return ("pentecost", [], False)

    # Ordinary Time Sundays -> Lord's Day preface (3 options)
    return ("lords_day", ["of_god_the_father", "of_god_the_son",
                          "of_god_the_holy_spirit"], True)


# ---------------------------------------------------------------------------
# Main rule generation
# ---------------------------------------------------------------------------

def get_seasonal_rules(title: str, color: str, notes: str,
                       pop_form: str = "") -> SeasonalRules:
    """Determine all seasonal rules for a service.

    Args:
        title: Liturgical day title (e.g., "Third Sunday in Lent")
        color: Liturgical color (e.g., "Violet", "White", "Green")
        notes: Notes from the Google Sheet
        pop_form: Prayers of the People form designation (e.g., "I", "VI (w/ confession)")
    """
    season = _detect_season(title, color, notes)
    is_lent = _is_lent(season)
    is_easter = _is_in_easter_season(season)
    is_advent = season == "advent"

    # Penitential Order: all Sundays in Lent
    use_penitential = is_lent
    use_decalogue = is_lent and _is_lent_1(title)

    # Opening Acclamation
    if use_penitential:
        acc_cel, acc_ppl = _ACCLAMATION_LENT
    elif is_easter:
        acc_cel, acc_ppl = _ACCLAMATION_EASTER
    else:
        acc_cel, acc_ppl = _ACCLAMATION_STANDARD

    # Song of Praise / Kyrie label
    if is_lent:
        sop_label = "Kyrie"
    elif is_advent:
        sop_label = "Song of Praise and Lighting of the Advent Wreath"
    else:
        sop_label = "Song of Praise"

    # Breaking of the Bread
    if is_lent:
        use_fraction = True
        frac_cel = ""
        frac_ppl = ""
    elif is_easter:
        use_fraction = False
        frac_cel, frac_ppl = _FRACTION_ALLELUIA
    else:
        use_fraction = False
        frac_cel, frac_ppl = _FRACTION_NO_ALLELUIA

    # Blessing vs Prayer over the People
    if is_lent:
        use_pop_prayer = True
        blessing_label = "Prayer over the People"
    else:
        use_pop_prayer = False
        blessing_label = "Blessing"

    # Dismissal: sheet column governs the number; we just set the Alleluia flag
    dismissal_alleluia = is_easter

    # POP (Prayers of the People) special handling
    pop_confession = _pop_has_confession(pop_form, notes)
    pop_use_bcp_vi = pop_confession  # BCP Form VI only when "(w/ confession)"

    # Proper Preface
    preface_key, preface_options, prompt_preface = _get_proper_preface_key(
        title, season)

    return SeasonalRules(
        use_penitential_order=use_penitential,
        use_decalogue=use_decalogue,
        confession_before_word=use_penitential,
        no_confession_after_pop=use_penitential or pop_confession,
        include_collect_for_purity=not use_penitential,
        acclamation_celebrant=acc_cel,
        acclamation_people=acc_ppl,
        song_of_praise_label=sop_label,
        is_advent=is_advent,
        use_fraction_anthem=use_fraction,
        fraction_celebrant=frac_cel,
        fraction_people=frac_ppl,
        use_prayer_over_people=use_pop_prayer,
        blessing_label=blessing_label,
        dismissal_has_alleluia=dismissal_alleluia,
        pop_use_bcp_form_vi=pop_use_bcp_vi,
        pop_has_confession=pop_confession,
        proper_preface_key=preface_key,
        proper_preface_options=preface_options,
        prompt_preface=prompt_preface,
        season=season,
    )
