"""
FastAPI app for the local bulletin web UI.

This module owns the HTTP layer only. All bulletin-domain logic lives
under ``bulletin.*`` and is imported here as if it were a third-party
library — nothing in ``bulletin/`` imports from ``web/``.

Routes (v1):

  GET  /                  generate.html (placeholder for v1)
  GET  /songs             searchable list of songs
  GET  /songs/new         add-song form (paste OR markdown upload)
  POST /songs/preview     parse + render a preview without saving
  POST /songs/save        parse + write to songs.yaml / hidden_springs_songs.yaml

Both libraries (the main 8/9/11 am ``songs.yaml`` and
``hidden_springs_songs.yaml``) share the same templates; the
``library`` query param / form field selects which file to read or
write.

The save path uses ``ruamel.yaml`` in round-trip mode so existing
hand-written entries keep their comments, key order, and quote style.
"""

from __future__ import annotations

import io
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, Form, Query, Request, UploadFile, File
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from ruamel.yaml import YAML

from web.song_parser import parse_markdown, parse_paste


# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------

WEB_DIR = Path(__file__).resolve().parent
REPO_ROOT = WEB_DIR.parent
HYMNS_DIR = REPO_ROOT / "bulletin" / "data" / "hymns"

LIBRARY_FILES: dict[str, Path] = {
    "main":           HYMNS_DIR / "songs.yaml",
    "hidden_springs": HYMNS_DIR / "hidden_springs_songs.yaml",
}
LIBRARY_LABELS: dict[str, str] = {
    "main":           "bulletin/data/hymns/songs.yaml",
    "hidden_springs": "bulletin/data/hymns/hidden_springs_songs.yaml",
}


# ---------------------------------------------------------------------------
# YAML loader (round-trip preserves comments / key order / quote style)
# ---------------------------------------------------------------------------

_yaml = YAML(typ="rt")
_yaml.width = 4096                  # don't wrap long lyric lines
_yaml.indent(mapping=2, sequence=2, offset=0)
_yaml.preserve_quotes = True


def _load_library(library: str) -> list[dict]:
    path = LIBRARY_FILES[library]
    if not path.exists():
        return []
    with open(path, encoding="utf-8") as f:
        data = _yaml.load(f)
    return list(data) if data else []


def _save_library(library: str, songs: list[dict]) -> None:
    path = LIBRARY_FILES[library]
    with open(path, "w", encoding="utf-8") as f:
        _yaml.dump(songs, f)


def _yaml_dump_one(song: dict) -> str:
    """Serialize a single song dict to a YAML snippet for preview."""
    buf = io.StringIO()
    # Wrap in a list so the snippet shows the leading "- " marker the
    # user will see in the file.
    _yaml.dump([song], buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# App + templates
# ---------------------------------------------------------------------------

app = FastAPI(title="St. Andrew's Bulletin Generator")

app.mount(
    "/static",
    StaticFiles(directory=str(WEB_DIR / "static")),
    name="static",
)

templates = Jinja2Templates(directory=str(WEB_DIR / "templates"))


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.get("/", response_class=HTMLResponse, name="home")
def home(request: Request) -> HTMLResponse:
    return templates.TemplateResponse(
        request, "generate.html",
        {"active": "home"},
    )


@app.get("/songs", response_class=HTMLResponse, name="songs_list")
def songs_list(
    request: Request,
    library: str = Query("main"),
) -> HTMLResponse:
    if library not in LIBRARY_FILES:
        library = "main"
    songs = _load_library(library)
    # ruamel.yaml returns CommentedMap / CommentedSeq objects; the
    # templates treat them as plain dicts/lists, which works.
    return templates.TemplateResponse(
        request, "songs/list.html",
        {
            "active": "songs",
            "songs": songs,
            "library": library,
            "library_label": LIBRARY_LABELS[library],
        },
    )


@app.get("/songs/new", response_class=HTMLResponse, name="song_new")
def song_new(
    request: Request,
    library: str = Query("main"),
    title: Optional[str] = Query(None),
) -> HTMLResponse:
    if library not in LIBRARY_FILES:
        library = "main"
    return templates.TemplateResponse(
        request, "songs/new.html",
        {
            "active": "songs",
            "library": library,
            "library_label": LIBRARY_LABELS[library],
            "prefill_title": title or "",
            "preview": None,
        },
    )


# --------------------------------------------------------------------------
# Parse helpers shared by /songs/preview and /songs/save
# --------------------------------------------------------------------------

async def _parse_submission(
    *,
    source: str,
    title: str,
    services: str,
    hymnal_number: str,
    hymnal_name: str,
    lyrics: str,
    md_file: Optional[UploadFile],
) -> tuple[dict, Optional[str]]:
    """Run the song parser on the submitted form. Returns (song, error)."""
    services = services.strip() or None
    if source == "markdown":
        if md_file is None or md_file.filename == "":
            return {}, "Please choose a markdown file to upload."
        raw = (await md_file.read()).decode("utf-8", errors="replace")
        song = parse_markdown(raw, services=services)
        if not song.get("title"):
            return song, (
                "Could not find a title in the markdown file. The first "
                "line should be a heading like `# **Title of the Song**`."
            )
        return song, None

    # Paste path
    if not title.strip():
        return {}, "Please give the song a title."
    if not lyrics.strip():
        return {}, "Please paste the lyrics."
    song = parse_paste(
        lyrics,
        title=title.strip(),
        hymnal_number=(hymnal_number.strip() or None),
        hymnal_name=(hymnal_name.strip() or None),
        services=services,
    )
    return song, None


@app.post("/songs/preview", response_class=HTMLResponse, name="song_preview")
async def song_preview(
    request: Request,
    library: str = Form("main"),
    source: str = Form("paste"),
    title: str = Form(""),
    services: str = Form(""),
    hymnal_number: str = Form(""),
    hymnal_name: str = Form(""),
    lyrics: str = Form(""),
    md_file: Optional[UploadFile] = File(None),
) -> HTMLResponse:
    if library not in LIBRARY_FILES:
        library = "main"

    song, error = await _parse_submission(
        source=source,
        title=title, services=services,
        hymnal_number=hymnal_number, hymnal_name=hymnal_name,
        lyrics=lyrics, md_file=md_file,
    )

    return templates.TemplateResponse(
        request, "songs/new.html",
        {
            "active": "songs",
            "library": library,
            "library_label": LIBRARY_LABELS[library],
            "prefill_title": song.get("title", ""),
            "preview": song if song else None,
            "preview_yaml": _yaml_dump_one(song) if song else "",
            "error": error,
        },
    )


@app.post("/songs/save", name="song_save")
async def song_save(
    request: Request,
    library: str = Form("main"),
    source: str = Form("paste"),
    title: str = Form(""),
    services: str = Form(""),
    hymnal_number: str = Form(""),
    hymnal_name: str = Form(""),
    lyrics: str = Form(""),
    md_file: Optional[UploadFile] = File(None),
):
    if library not in LIBRARY_FILES:
        library = "main"

    song, error = await _parse_submission(
        source=source,
        title=title, services=services,
        hymnal_number=hymnal_number, hymnal_name=hymnal_name,
        lyrics=lyrics, md_file=md_file,
    )

    if error or not song:
        # Re-render the form with the error and the preview so the user
        # can see what went wrong.
        return templates.TemplateResponse(
            request, "songs/new.html",
            {
                "active": "songs",
                "library": library,
                "library_label": LIBRARY_LABELS[library],
                "prefill_title": song.get("title", "") if song else title,
                "preview": song if song else None,
                "preview_yaml": _yaml_dump_one(song) if song else "",
                "error": error or "Could not parse the submission.",
            },
        )

    # Append to the chosen library and write back.
    songs = _load_library(library)
    songs.append(song)
    _save_library(library, songs)

    # Invalidate the in-process cache so the bulletin builder picks up
    # the new song the next time it generates.
    try:
        from bulletin.sources.songs import clear_cache
        clear_cache()
    except Exception:  # pragma: no cover — defensive
        pass

    url = request.url_for("songs_list").include_query_params(library=library)
    return RedirectResponse(url=str(url), status_code=303)
