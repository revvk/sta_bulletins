"""Confirm that ``clear_all_caches()`` actually flushes the loader's
``@lru_cache`` so the next read sees a freshly-written YAML.

This protects the web app's save endpoints: long-running uvicorn
processes that edit data files in place would otherwise serve stale
data until the user restarted the server.

Run via::

    python3.11 -m pytest bulletin/tests/test_loader_cache.py -v
"""

from __future__ import annotations

import textwrap
from pathlib import Path

import pytest


def test_clear_all_caches_invalidates_lru(tmp_path: Path, monkeypatch):
    """Write a YAML, read it via the loader, mutate the file, and
    confirm clear_all_caches() makes the next read see the new value.

    Uses ``monkeypatch`` to point ``loader._DATA_DIR`` at a tmpdir so we
    don't touch real bulletin data.
    """
    from bulletin.data import loader

    fake_yaml = tmp_path / "fake.yaml"
    fake_yaml.write_text("greeting: hello\n", encoding="utf-8")

    monkeypatch.setattr(loader, "_DATA_DIR", tmp_path)
    # The cache may already hold entries from earlier tests/imports —
    # start from a known-empty state.
    loader.clear_all_caches()

    first = loader._load_yaml("fake.yaml")
    assert first == {"greeting": "hello"}

    # Mutate the file underneath the cache. Without clear_all_caches()
    # the loader would still hand back "hello".
    fake_yaml.write_text("greeting: world\n", encoding="utf-8")
    stale = loader._load_yaml("fake.yaml")
    assert stale == {"greeting": "hello"}, (
        "cache hit expected before clear_all_caches() — if this fails "
        "then the @lru_cache decorator is no longer in place"
    )

    loader.clear_all_caches()
    fresh = loader._load_yaml("fake.yaml")
    assert fresh == {"greeting": "world"}, (
        "clear_all_caches() did not invalidate the loader cache — the "
        "web app would serve stale YAML after a save"
    )


def test_clear_all_caches_is_idempotent():
    """Calling clear_all_caches() on an already-empty cache must not
    raise. The web app may call it from several save endpoints in
    quick succession."""
    from bulletin.data import loader

    loader.clear_all_caches()
    loader.clear_all_caches()  # should be a no-op
