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

import html
import io
import re
import uuid
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, Form, Query, Request, UploadFile, File
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from markupsafe import Markup

from ruamel.yaml import YAML

from bulletin.data.loader import load_pop_forms
from bulletin.report import RunReport
from bulletin.runner import RunAborted, RunOptions, run_generation
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


# Jinja filter: wrap {placeholder} tokens in the POP form text with
# <code> so they stand out visually. Anything outside the braces is
# escaped as normal; the replacement itself is also HTML-safe.
_PLACEHOLDER_RE = re.compile(r"\{[a-zA-Z0-9_]+\}")


def _highlight_placeholders(text: str) -> Markup:
    if text is None:
        return Markup("")
    escaped = html.escape(str(text))
    out = _PLACEHOLDER_RE.sub(
        lambda m: f"<code>{m.group(0)}</code>", escaped)
    return Markup(out)


templates.env.filters["highlight_placeholders"] = _highlight_placeholders


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

# --------------------------------------------------------------------------
# In-memory store of recent runs. Single-user app, so a dict keyed by a
# uuid is fine — no persistence needed across restarts.
# --------------------------------------------------------------------------

_RECENT_RUNS: "dict[str, dict]" = {}
_RECENT_RUN_ORDER: list[str] = []
_MAX_RECENT_RUNS = 20

_SERVICE_LABELS = {
    "all":            "All Sunday services",
    "8 am":           "8 am",
    "9 am":           "9 am",
    "11 am":          "11 am",
    "sunrise":        "Easter sunrise",
    "hidden_springs": "Hidden Springs",
    "7 pm":           "Special weekday (7 pm)",
}


def _next_sunday(today: Optional[date] = None) -> date:
    today = today or date.today()
    # weekday(): Mon=0, …, Sun=6
    days_ahead = (6 - today.weekday()) % 7
    if days_ahead == 0:
        days_ahead = 7
    return today + timedelta(days=days_ahead)


def _store_run(run: dict) -> str:
    run_id = uuid.uuid4().hex[:12]
    run["run_id"] = run_id
    _RECENT_RUNS[run_id] = run
    _RECENT_RUN_ORDER.insert(0, run_id)
    while len(_RECENT_RUN_ORDER) > _MAX_RECENT_RUNS:
        old = _RECENT_RUN_ORDER.pop()
        _RECENT_RUNS.pop(old, None)
    return run_id


def _recent_runs() -> list[dict]:
    return [_RECENT_RUNS[r] for r in _RECENT_RUN_ORDER if r in _RECENT_RUNS]


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.get("/", response_class=HTMLResponse, name="home")
def home(request: Request, error: Optional[str] = Query(None)) -> HTMLResponse:
    return templates.TemplateResponse(
        request, "generate.html",
        {
            "active": "home",
            "default_date": _next_sunday().isoformat(),
            "recent_runs": _recent_runs(),
            "error": error,
        },
    )


@app.post("/run", name="run_bulletin")
def run_bulletin(
    request: Request,
    target_date: str = Form(...),
    service: str = Form("all"),
    reading_sheets: Optional[str] = Form(None),
    force_fetch: Optional[str] = Form(None),
):
    try:
        parsed_date = datetime.strptime(target_date, "%Y-%m-%d").date()
    except ValueError:
        url = request.url_for("home").include_query_params(
            error=f"Invalid date '{target_date}'. Use YYYY-MM-DD.")
        return RedirectResponse(url=str(url), status_code=303)

    options = RunOptions(
        target_date=parsed_date,
        service=service,
        output_dir=REPO_ROOT / "output",
        reading_sheets=bool(reading_sheets),
        force_fetch=bool(force_fetch),
    )

    # Capture all progress lines into a buffer so the report page can
    # show them (collapsed by default). Web runs are non-interactive —
    # the CLI's prompt_choice() is replaced with None, so any
    # disambiguation falls through to the song-list defaults.
    console_lines: list[str] = []
    def progress(line: str) -> None:
        console_lines.append(str(line))

    report = RunReport()
    try:
        result = run_generation(
            options,
            prompt_fn=None,
            progress_fn=progress,
            report=report,
        )
    except RunAborted as e:
        url = request.url_for("home").include_query_params(error=str(e))
        return RedirectResponse(url=str(url), status_code=303)
    except Exception as e:  # pragma: no cover — surface unexpected errors
        url = request.url_for("home").include_query_params(
            error=f"Unexpected error: {e}")
        return RedirectResponse(url=str(url), status_code=303)

    run_id = _store_run({
        "target_date":      result.target_date,
        "service":          service,
        "service_label":    _SERVICE_LABELS.get(service, service),
        "bulletins":        result.bulletins,
        "reading_sheets":   result.reading_sheets,
        "report":           result.report,
        "console":          "\n".join(console_lines),
    })

    url = request.url_for("run_report", run_id=run_id)
    return RedirectResponse(url=str(url), status_code=303)


@app.get("/report/{run_id}", response_class=HTMLResponse, name="run_report")
def run_report(request: Request, run_id: str) -> HTMLResponse:
    run = _RECENT_RUNS.get(run_id)
    if run is None:
        url = request.url_for("home").include_query_params(
            error="That report has expired. Generate again to see fresh output.")
        return RedirectResponse(url=str(url), status_code=303)
    return templates.TemplateResponse(
        request, "report.html",
        {
            "active": "home",
            "run": run,
            "output_dir": str(REPO_ROOT / "output"),
        },
    )


def _has_lyrics(song: dict) -> bool:
    """True if the song carries any lines we could render.

    Some entries (especially in ``hidden_springs_songs.yaml``) are
    title-only stubs for instrumentals or songs still missing lyrics.
    A song counts as having lyrics as long as at least one section has
    at least one non-blank line.
    """
    for sec in (song.get("sections") or []):
        for line in (sec.get("lines") or []):
            if str(line).strip():
                return True
    return False


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
            "has_lyrics": _has_lyrics,
        },
    )


@app.get("/songs/view", response_class=HTMLResponse, name="song_view")
def song_view(
    request: Request,
    library: str = Query("main"),
    i: int = Query(..., ge=0),
) -> HTMLResponse:
    """Render a single song's lyrics, addressed by library + list index.

    Using the file position as the key keeps the URL stable for as long
    as no one edits the YAML out from under the page. Duplicate titles
    (e.g. 10,000 Reasons at 9 am and 11 am) are disambiguated for free.
    """
    if library not in LIBRARY_FILES:
        library = "main"
    songs = _load_library(library)
    if i >= len(songs):
        url = request.url_for("songs_list").include_query_params(library=library)
        return RedirectResponse(url=str(url), status_code=303)
    return templates.TemplateResponse(
        request, "songs/view.html",
        {
            "active": "songs",
            "library": library,
            "library_label": LIBRARY_LABELS[library],
            "song": songs[i],
            "index": i,
            "has_lyrics": _has_lyrics(songs[i]),
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


# ---------------------------------------------------------------------------
# Prayers of the People (POP forms)
# ---------------------------------------------------------------------------
#
# The pop_forms.yaml catalog stores one entry per form variant. We group
# them for display so that (for example) Form II and Form II (Immigration
# Focus) sit next to each other. The grouping also tells the user what
# text to drop into the Google Sheet's POP column to pick a specific
# variant — that matters because the CLI used to resolve ambiguity with
# an interactive prompt that the web runner can't answer.

# Every field the page needs — group membership, display label, and the
# exact sheet values that pick each form — is read out of pop_forms.yaml.
# Adding a new form is purely a YAML edit; no code changes required.
#
# YAML shape per entry:
#     group:         short id shared by every variant in a family
#                    (e.g. "form_II", "advent", "easter", "special")
#     group_label:   human label for the panel header. Only the "base"
#                    entry in a group needs to set this; variants inherit.
#     sheet_values:  list of strings the planner can type in the POP
#                    column to select this exact form. Matching is
#                    case-insensitive.


def _planner_hint(form: dict) -> dict:
    """Return info describing how to select this form from the planner.

    Shape::

        {"sheet_value": "II (immigration)", "note": None}
        # or, if sheet_values is empty:
        {"sheet_value": None, "note": "…explanation…"}
    """
    values = [str(v) for v in (form.get("sheet_values") or [])]
    if values:
        note = None
        if len(values) > 1:
            aliases = ", ".join(f'"{v}"' for v in values[1:])
            note = f"Also accepts: {aliases}"
        return {"sheet_value": values[0], "note": note}
    return {
        "sheet_value": None,
        "note": (
            "No POP-column selector is configured for this form. Add a "
            "`sheet_values:` list to its entry in pop_forms.yaml."
        ),
    }


def _prayer_catalog() -> list[dict]:
    """Build the grouped catalog of POP forms for the list page.

    Entries are grouped by their YAML ``group`` field. The first entry
    that declares ``group_label`` supplies the panel header. Display
    order preserves the order of appearance in pop_forms.yaml — so the
    YAML itself is the knob for reordering the page.
    """
    forms = load_pop_forms() or {}
    groups: dict[str, dict] = {}

    for key, form in forms.items():
        gid = form.get("group") or key
        if gid not in groups:
            groups[gid] = {
                "id":      gid,
                "label":   form.get("group_label") or gid.replace("_", " ").title(),
                "entries": [],
            }
        elif form.get("group_label") and not groups[gid].get("_label_set"):
            groups[gid]["label"] = form["group_label"]
            groups[gid]["_label_set"] = True

        groups[gid]["entries"].append({
            "key":        key,
            "title":      form.get("title", key),
            "hint":       _planner_hint(form),
            "element_ct": len(form.get("elements", []) or []),
        })

    # Drop the internal bookkeeping flag before handing groups to the view.
    for g in groups.values():
        g.pop("_label_set", None)

    return list(groups.values())


@app.get("/prayers", response_class=HTMLResponse, name="prayers_list")
def prayers_list(request: Request) -> HTMLResponse:
    return templates.TemplateResponse(
        request, "prayers/list.html",
        {
            "active": "prayers",
            "groups": _prayer_catalog(),
        },
    )


@app.get("/prayers/view", response_class=HTMLResponse, name="prayer_view")
def prayer_view(request: Request, key: str = Query(...)) -> HTMLResponse:
    forms = load_pop_forms() or {}
    form = forms.get(key)
    if form is None:
        url = request.url_for("prayers_list")
        return RedirectResponse(url=str(url), status_code=303)
    return templates.TemplateResponse(
        request, "prayers/view.html",
        {
            "active": "prayers",
            "key":    key,
            "form":   form,
            "hint":   _planner_hint(form),
        },
    )
