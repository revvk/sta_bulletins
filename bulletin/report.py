"""
Structured run report for bulletin generation.

Replaces the historical pattern of scattered ``print("Warning: …")`` calls
with a single object that accumulates issues and renders them either to
the console (CLI) or to JSON (web UI).

Design goals:
- **Backwards compatible.** Every consumer that doesn't pass a report just
  keeps printing warnings to stdout exactly as before. Threading a
  ``RunReport`` in is purely additive.
- **One module, no other dependencies.** This file imports nothing from
  the rest of ``bulletin/`` so any source can import ``RunReport`` without
  worrying about circular imports.
- **Three severity buckets** so the post-generation TODO list reads like
  a punch list rather than a wall of text:

    * ``blocker`` — bulletin is wrong/incomplete; user must fix before
      printing (e.g. scripture fetch failed, POP form unknown)
    * ``warning`` — bulletin generated with a fallback/placeholder; user
      should review (e.g. missing song lyrics, ministry rotation
      missing)
    * ``manual``  — bulletin is fine but user has a hand-edit to make in
      the .docx (e.g. paste lyrics into a placeholder, double-check a
      preacher name)

The categories are free-form short strings (``song``, ``scripture``,
``ministry``, ``aac``, ``manual_paste``, …) — keeping them strings means
sources can introduce new categories without coordinating with this file.
"""

from __future__ import annotations

from dataclasses import dataclass, field, asdict
from typing import Literal, Optional


Severity = Literal["blocker", "warning", "manual"]


@dataclass
class TodoItem:
    """One actionable line on the post-generation report."""

    severity: Severity
    category: str
    message: str
    fix_hint: Optional[str] = None
    link: Optional[str] = None  # e.g. "/songs/new?title=…"

    def to_dict(self) -> dict:
        return asdict(self)


# Console glyphs per severity. ASCII fallbacks are used if stdout can't
# encode emoji (rare on macOS; common on Windows shells).
_GLYPHS = {
    "blocker": ("🛑", "[!!]"),
    "warning": ("⚠️", "[ ! ]"),
    "manual":  ("📝", "[ . ]"),
}

_SEVERITY_LABELS = {
    "blocker": "Blockers",
    "warning": "Warnings",
    "manual":  "Manual TODOs",
}

# Display order for grouping
_SEVERITY_ORDER: tuple[Severity, ...] = ("blocker", "warning", "manual")


@dataclass
class RunReport:
    """Accumulates :class:`TodoItem` entries during a generation run.

    Pass an instance to ``BulletinBuilder`` (or any source-layer fetcher
    that accepts a ``report=`` kwarg). At the end of generation, call
    :meth:`print_console` for the CLI or :meth:`to_dict` for the web UI.
    """

    items: list[TodoItem] = field(default_factory=list)

    # ------------------------------------------------------------------ add

    def add(
        self,
        *,
        severity: Severity,
        category: str,
        message: str,
        fix_hint: Optional[str] = None,
        link: Optional[str] = None,
    ) -> None:
        """Record a new item. Keyword-only to make call sites self-documenting."""
        self.items.append(TodoItem(
            severity=severity,
            category=category,
            message=message,
            fix_hint=fix_hint,
            link=link,
        ))

    # Convenience shorthands — most sources will want these instead of
    # spelling out severity= every time.

    def blocker(self, category: str, message: str,
                fix_hint: Optional[str] = None,
                link: Optional[str] = None) -> None:
        self.add(severity="blocker", category=category, message=message,
                 fix_hint=fix_hint, link=link)

    def warning(self, category: str, message: str,
                fix_hint: Optional[str] = None,
                link: Optional[str] = None) -> None:
        self.add(severity="warning", category=category, message=message,
                 fix_hint=fix_hint, link=link)

    def manual(self, category: str, message: str,
               fix_hint: Optional[str] = None,
               link: Optional[str] = None) -> None:
        self.add(severity="manual", category=category, message=message,
                 fix_hint=fix_hint, link=link)

    # ----------------------------------------------------------------- query

    def __len__(self) -> int:
        return len(self.items)

    def __bool__(self) -> bool:
        return bool(self.items)

    def by_severity(self, severity: Severity) -> list[TodoItem]:
        return [it for it in self.items if it.severity == severity]

    def has_blockers(self) -> bool:
        return any(it.severity == "blocker" for it in self.items)

    # --------------------------------------------------------------- output

    def to_dict(self) -> dict:
        """JSON-serializable form, used by the web UI."""
        return {
            "items": [it.to_dict() for it in self.items],
            "counts": {
                sev: len(self.by_severity(sev))
                for sev in _SEVERITY_ORDER
            },
        }

    def print_console(self, *, use_emoji: bool = True) -> None:
        """Render the report to stdout, grouped by severity.

        Designed to be called once at the end of ``generate.py``.
        Prints nothing if the report is empty, so no-issues runs stay
        quiet.
        """
        if not self.items:
            return

        print()
        print("  " + "─" * 56)
        print("  Generation Report")
        print("  " + "─" * 56)

        for sev in _SEVERITY_ORDER:
            bucket = self.by_severity(sev)
            if not bucket:
                continue
            glyph = _GLYPHS[sev][0 if use_emoji else 1]
            label = _SEVERITY_LABELS[sev]
            print(f"\n  {glyph} {label} ({len(bucket)})")
            for item in bucket:
                # Two-space indent + bullet, matching the existing
                # generate.py output style.
                print(f"     • {item.message}")
                if item.fix_hint:
                    print(f"         → {item.fix_hint}")
