"""
Loader and typed access for per-service funeral YAML files
(`bulletin/data/funerals/services/<slug>.yaml`).

Each YAML captures one funeral or memorial service end-to-end (rite,
HC × Commendation × Committal flags, deceased details, participants,
readings, music, liturgical choices). See
docs/plans/funeral_bulletins.md for the full schema and
bulletin/data/funerals/services/2026-01-31-cox.yaml for the reference.

Public surface
==============

``load_service(slug_or_path)``
    Locate, parse, validate, and return a ``FuneralData`` dataclass.

``FuneralData``
    Frozen dataclass exposing the parsed YAML plus a few derived
    convenience attributes:

    - ``cover_subtitle``     —  "The Burial of the Dead", "Memorial
                                 Service", etc. (overridable in YAML).
    - ``life_dates``         —  "November 30, 1932 – January 8, 2026"
                                 (en-dash). Falls back to just the death
                                 date if ``born`` is absent.
    - ``service_date_long``  —  "January 31, 2026"
    - ``substitutions``      —  Dict of placeholder → text for the
                                 pronoun/name substitutions used by the
                                 builder (``{N}``, ``{he}``, ``{him}``,
                                 ``{his}``, ``{brother}``).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any

import yaml


# ---------------------------------------------------------------------------
# Where service YAMLs live
# ---------------------------------------------------------------------------

_SERVICES_DIR = (
    Path(__file__).resolve().parent.parent
    / "data" / "funerals" / "services"
)


# ---------------------------------------------------------------------------
# Subtitle defaults derived from service_kind. Per-service `cover_subtitle:`
# (if present) overrides this.
# ---------------------------------------------------------------------------

_DEFAULT_SUBTITLES = {
    "burial":    "The Burial of the Dead",
    "memorial":  "Memorial Service",
    "committal": "Graveside Service",
}


# ---------------------------------------------------------------------------
# Pronoun substitution tables. Keyed by deceased.pronoun.
# ---------------------------------------------------------------------------

_PRONOUNS = {
    "he":   {"{he}": "he",   "{he_cap}": "He",   "{him}": "him",  "{his}": "his",   "{brother}": "brother"},
    "she":  {"{he}": "she",  "{he_cap}": "She",  "{him}": "her",  "{his}": "her",   "{brother}": "sister"},
    "they": {"{he}": "they", "{he_cap}": "They", "{him}": "them", "{his}": "their", "{brother}": "sibling"},
}


# ---------------------------------------------------------------------------
# FuneralData dataclass
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class FuneralData:
    """Parsed per-service funeral YAML, with a few derived fields."""

    # ----- Raw structural fields (1-to-1 with the YAML) ----------------

    service_kind:         str            # "burial" | "memorial" | "committal"
    rite:                 str            # "I" | "II"
    holy_eucharist:       dict           # {enabled: bool, prayer: str|None}
    include_commendation: bool
    include_committal:    bool

    deceased:     dict
    service:      dict
    participants: dict
    readings:     dict
    music:        dict
    collects:     dict
    prayers_for_the_departed: dict
    special_prayers: list                # [{key, reader, reader_role}, ...]
    reception:    dict
    interment_notice: str | None
    cover_subtitle:   str | None         # explicit YAML override, may be None

    # The original parsed dict — handy for debugging / future fields.
    _raw: dict = field(repr=False)

    # ----- Convenience accessors ---------------------------------------

    @property
    def slug(self) -> str:
        """e.g. '2026-01-31-cox' — used for output filenames."""
        d  = self._fmt_date_iso(self.service["date"])
        nm = self.deceased["full_name"].split()[-1].lower()
        return f"{d}-{nm}"

    @property
    def service_date_long(self) -> str:
        """e.g. 'January 31, 2026'."""
        return self._fmt_date_long(self.service["date"])

    @property
    def life_dates(self) -> str:
        """Formatted birth–death range, en-dash separated. Falls back to
        just the death date when ``born`` is unset."""
        born = self.deceased.get("born")
        died = self.deceased.get("died")
        if born and died:
            return f"{self._fmt_date_long(born)} – {self._fmt_date_long(died)}"
        if died:
            return self._fmt_date_long(died)
        return ""

    @property
    def cover_subtitle_resolved(self) -> str:
        """Per-service override wins; otherwise default by service_kind."""
        if self.cover_subtitle:
            return self.cover_subtitle
        return _DEFAULT_SUBTITLES.get(self.service_kind, "Funeral Service")

    @property
    def substitutions(self) -> dict[str, str]:
        """Placeholder → text dict for {N} and pronoun substitutions.

        ``{NN}`` (the bereaved family-member names, used only in the
        for-those-who-mourn collect) is left out here; the builder
        fills it from the per-service YAML's bereaved-names field.
        """
        pronoun = self.deceased.get("pronoun", "they").lower()
        sub = dict(_PRONOUNS.get(pronoun, _PRONOUNS["they"]))
        sub["{N}"] = self.deceased.get("preferred_name") or self.deceased.get("full_name", "")
        return sub

    # ----- Date formatting helpers -------------------------------------

    @staticmethod
    def _fmt_date_long(d) -> str:
        """Accept date / datetime / 'YYYY-MM-DD' string. Return e.g.
        'January 31, 2026' (no leading zero on day)."""
        if isinstance(d, str):
            d = datetime.strptime(d, "%Y-%m-%d").date()
        if isinstance(d, datetime):
            d = d.date()
        return d.strftime("%B %-d, %Y")

    @staticmethod
    def _fmt_date_iso(d) -> str:
        """ISO 'YYYY-MM-DD'."""
        if isinstance(d, str):
            return d
        if isinstance(d, datetime):
            d = d.date()
        return d.isoformat()


# ---------------------------------------------------------------------------
# Loading + validation
# ---------------------------------------------------------------------------

def load_service(slug_or_path: str | Path) -> FuneralData:
    """Locate, parse, validate, and return a per-service ``FuneralData``.

    ``slug_or_path`` may be:
      - A bare slug like ``"2026-01-31-cox"`` (resolved against
        ``bulletin/data/funerals/services/<slug>.yaml``).
      - A path (relative or absolute) to a YAML file.
    """
    path = _resolve_path(slug_or_path)
    if not path.exists():
        raise FileNotFoundError(
            f"Funeral service YAML not found: {path}\n"
            f"  Expected at: {_SERVICES_DIR / (str(slug_or_path) + '.yaml')}\n"
            f"  Or pass an explicit path."
        )
    with open(path, encoding="utf-8") as f:
        raw = yaml.safe_load(f)

    _validate(raw, path)
    return _from_raw(raw)


def _resolve_path(slug_or_path: str | Path) -> Path:
    """Resolve a bare slug to a path under the services directory; pass
    explicit paths through unchanged."""
    p = Path(slug_or_path)
    if p.suffix == ".yaml" or p.is_absolute() or p.parent != Path("."):
        return p.resolve() if not p.is_absolute() else p
    return _SERVICES_DIR / f"{slug_or_path}.yaml"


def _validate(raw: dict, path: Path) -> None:
    """Quick schema check — surface obvious mistakes before they become
    confusing render errors. Not exhaustive; the builder still has to
    handle missing optional fields gracefully."""
    required = [
        "service_kind", "rite",
        "holy_eucharist", "include_commendation", "include_committal",
        "deceased", "service",
    ]
    missing = [k for k in required if k not in raw]
    if missing:
        raise ValueError(
            f"{path.name}: missing required top-level fields: {missing}"
        )

    if raw["service_kind"] not in {"burial", "memorial", "committal"}:
        raise ValueError(
            f"{path.name}: service_kind must be one of "
            "'burial', 'memorial', 'committal'; got {raw['service_kind']!r}"
        )

    if raw["rite"] not in {"I", "II"}:
        raise ValueError(
            f"{path.name}: rite must be 'I' or 'II'; got {raw['rite']!r}"
        )

    he = raw["holy_eucharist"]
    if he.get("enabled"):
        prayer = he.get("prayer")
        valid_for_rite = {
            "I":  {"I", "II"},
            "II": {"A", "B", "C", "D"},
        }
        if prayer not in valid_for_rite[raw["rite"]]:
            raise ValueError(
                f"{path.name}: holy_eucharist.prayer={prayer!r} is not "
                f"valid for rite {raw['rite']!r}. "
                f"Expected one of {sorted(valid_for_rite[raw['rite']])}."
            )

    # The matrix is permissive but a memorial without remains can't have
    # a committal — flag that as a likely error.
    if raw["service_kind"] == "memorial" and raw.get("include_committal"):
        raise ValueError(
            f"{path.name}: service_kind='memorial' with "
            "include_committal=true. A memorial service has no remains "
            "to commit; double-check the matrix in "
            "docs/plans/funeral_bulletins.md."
        )


def _from_raw(raw: dict) -> FuneralData:
    """Build a FuneralData from the raw parsed dict, normalizing some
    optional fields."""
    return FuneralData(
        service_kind         = raw["service_kind"],
        rite                 = raw["rite"],
        holy_eucharist       = raw["holy_eucharist"],
        include_commendation = bool(raw["include_commendation"]),
        include_committal    = bool(raw["include_committal"]),
        deceased             = raw["deceased"],
        service              = raw["service"],
        participants         = raw.get("participants", {}),
        readings             = raw.get("readings", {}),
        music                = raw.get("music", {}),
        collects             = raw.get("collects", {}),
        prayers_for_the_departed = raw.get("prayers_for_the_departed", {}),
        special_prayers      = raw.get("special_prayers", []) or [],
        reception            = raw.get("reception", {"shown": False, "text": None}),
        interment_notice     = raw.get("interment_notice"),
        cover_subtitle       = raw.get("cover_subtitle"),
        _raw                 = raw,
    )
