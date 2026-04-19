"""
Golden-file tests for ``web.song_parser``.

These tests are deliberately structural: they assert section types,
section counts, and selected line content rather than full equality
against a frozen dict. That keeps the suite robust against trivial
formatting tweaks (a stray space, an extra blank line) while still
catching real regressions in the markdown / paste parsers.

The canonical fixture is ``source_documents/Alleluia, alleluia,
alleluia, The strife is o'er.md`` — the exact format Andrew uses when
hand-typing a new hymn.
"""

from __future__ import annotations

import unittest
from pathlib import Path

from web.song_parser import parse_markdown, parse_paste


REPO_ROOT = Path(__file__).resolve().parents[2]
ALLELUIA_MD = (
    REPO_ROOT
    / "source_documents"
    / "Alleluia, alleluia, alleluia, The strife is o\u2019er.md"
)


class TestMarkdownParser(unittest.TestCase):
    """Parse the Alleluia/strife-is-o'er fixture and verify the shape."""

    @classmethod
    def setUpClass(cls) -> None:
        cls.md = ALLELUIA_MD.read_text(encoding="utf-8")
        cls.result = parse_markdown(cls.md)

    def test_title_is_extracted_and_unbolded(self) -> None:
        # The fixture has `# **Alleluia, ...**` — the bold markers must
        # be stripped from the title so it round-trips into songs.yaml
        # the same way the existing entries are stored.
        self.assertEqual(
            self.result["title"],
            "Alleluia, alleluia, alleluia, The strife is o\u2019er",
        )

    def test_hymnal_reference_is_split(self) -> None:
        self.assertEqual(self.result["hymnal_number"], "208")
        self.assertEqual(self.result["hymnal_name"], "Hymnal 1982")

    def test_section_types_match_spec(self) -> None:
        # Spec from the plan: one antiphon (chorus) + 5 verses + one
        # trailing antiphon (chorus).
        types = [s["type"] for s in self.result["sections"]]
        self.assertEqual(types, ["chorus", "verse", "verse", "verse", "verse", "verse", "chorus"])

    def test_antiphon_lines_are_stripped_to_lyrics(self) -> None:
        # The italic label line ("*Antiphon (at the beginning)*") must
        # NOT appear in the lyric lines — only the actual lyric does.
        first_chorus = self.result["sections"][0]
        self.assertEqual(first_chorus["lines"], ["Alleluia, alleluia, alleluia!"])

    def test_verses_have_four_lines_each(self) -> None:
        # Verses 1-5 in the fixture are all four lines (three lyric
        # lines + one Alleluia refrain).
        for i in range(1, 6):
            section = self.result["sections"][i]
            self.assertEqual(section["type"], "verse")
            self.assertEqual(
                len(section["lines"]),
                4,
                f"Verse {i} should have four lines, got {section['lines']!r}",
            )

    def test_first_verse_first_line(self) -> None:
        self.assertEqual(
            self.result["sections"][1]["lines"][0],
            "The strife is o'er, the battle done;",
        )

    def test_markdown_escapes_are_unescaped(self) -> None:
        # The fixture has `Alleluia\!` with an escaped `!`. After
        # normalization, the bang should be there but the backslash
        # should NOT.
        for section in self.result["sections"]:
            for line in section["lines"]:
                self.assertNotIn(r"\!", line, f"Stray escape in: {line!r}")
                self.assertNotIn(r"\#", line, f"Stray escape in: {line!r}")

    def test_trailing_antiphon_present(self) -> None:
        # The closing chorus mirrors the opening one.
        last = self.result["sections"][-1]
        self.assertEqual(last["type"], "chorus")
        self.assertEqual(last["lines"], ["Alleluia, alleluia, alleluia!"])


class TestPasteParser(unittest.TestCase):
    """Verify the paste-text path matches the songs.yaml hand-format."""

    def test_simple_two_verse_paste(self) -> None:
        paste = (
            "Love's redeeming work is done;\n"
            "fought the fight, the battle won.\n"
            "\n"
            "Lives again our glorious King;\n"
            "where, O death, is now thy sting?\n"
        )
        result = parse_paste(
            paste,
            title="Love's redeeming work is done",
            hymnal_number="188",
        )
        self.assertEqual(result["title"], "Love's redeeming work is done")
        self.assertEqual(result["hymnal_number"], "188")
        self.assertEqual(result["hymnal_name"], "Hymnal 1982")
        types = [s["type"] for s in result["sections"]]
        self.assertEqual(types, ["verse", "verse"])
        self.assertEqual(
            result["sections"][0]["lines"],
            [
                "Love's redeeming work is done;",
                "fought the fight, the battle won.",
            ],
        )

    def test_chorus_label_promotes_block_to_chorus(self) -> None:
        paste = (
            "Verse one line one\n"
            "Verse one line two\n"
            "\n"
            "Chorus:\n"
            "All hail the power\n"
            "of Jesus' name\n"
            "\n"
            "Verse two line one\n"
            "Verse two line two\n"
        )
        result = parse_paste(paste, title="Test")
        types = [s["type"] for s in result["sections"]]
        self.assertEqual(types, ["verse", "chorus", "verse"])
        self.assertEqual(
            result["sections"][1]["lines"],
            ["All hail the power", "of Jesus' name"],
        )

    def test_refrain_keyword_is_alias_for_chorus(self) -> None:
        paste = "Refrain:\nAlleluia\n"
        result = parse_paste(paste, title="x")
        self.assertEqual(result["sections"][0]["type"], "chorus")

    def test_tag_keyword_emits_tag_section(self) -> None:
        paste = "Tag:\nForever and ever, Amen\n"
        result = parse_paste(paste, title="x")
        self.assertEqual(result["sections"][0]["type"], "tag")

    def test_chorus_word_in_lyric_is_not_a_label(self) -> None:
        # Guards against false positives like "Chorus of angels, raise..."
        paste = "Chorus of angels, raise your voice in song this day,\nand sing to him a hymn of praise.\n"
        result = parse_paste(paste, title="x")
        # Should be a single verse, NOT a chorus with truncated lyrics.
        self.assertEqual(result["sections"][0]["type"], "verse")
        self.assertEqual(len(result["sections"][0]["lines"]), 2)

    def test_no_hymnal_number_omits_hymnal_keys(self) -> None:
        result = parse_paste("verse line\n", title="x")
        self.assertNotIn("hymnal_number", result)
        self.assertNotIn("hymnal_name", result)

    def test_services_passthrough(self) -> None:
        result = parse_paste("verse line\n", title="x", services="9am")
        self.assertEqual(result["services"], "9am")

    def test_dict_key_order_matches_songs_yaml(self) -> None:
        # songs.yaml entries follow the order:
        #   title, services?, hymnal_number?, hymnal_name?, sections
        result = parse_paste(
            "verse\n",
            title="x",
            hymnal_number="100",
            services="9am",
        )
        self.assertEqual(
            list(result.keys()),
            ["title", "services", "hymnal_number", "hymnal_name", "sections"],
        )


class TestSharedCleanup(unittest.TestCase):
    """Cross-cutting line-cleanup behavior."""

    def test_blank_lines_collapse_runs(self) -> None:
        paste = "verse one\n\n\n\n\nverse two\n"
        result = parse_paste(paste, title="x")
        self.assertEqual(len(result["sections"]), 2)

    def test_crlf_line_endings_normalized(self) -> None:
        paste = "verse one\r\n\r\nverse two\r\n"
        result = parse_paste(paste, title="x")
        self.assertEqual(len(result["sections"]), 2)


if __name__ == "__main__":
    unittest.main()
