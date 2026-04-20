# Installing the St. Andrew's Bulletin Generator

This guide is written for the **technical helper** standing the tool up
on a Mac for a priest. Once steps 1–4 are done, the priest only ever
double-clicks one of two icons on their Desktop.

The whole install takes about ten minutes on a clean Mac, most of it
unattended (Homebrew, Python, and `pip install` doing their thing).

---

## What you'll end up with

- A copy of the project at `~/Bulletin/` (or wherever you choose).
- A self-contained Python environment at `~/Bulletin/.venv/` — nothing
  installed system-wide except Python itself.
- Two icons on the priest's Desktop:
  - **Bulletin.command** — opens the web UI in their browser.
  - **Update.command** — fetches the latest code and reinstalls.

---

## Prerequisites

You need:

- macOS 12 (Monterey) or newer.
- Admin rights to install Homebrew + Python (one time only).
- Network access to GitHub, PyPI, and `oremus.org` (for scripture).
- Access to the parish's Google Sheets (the sheet IDs are baked into
  `bulletin/config.py` for now).

---

## Step 1 — Install Homebrew (if not already present)

Open Terminal (`Applications → Utilities → Terminal`) and paste:

```sh
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
```

Follow Homebrew's prompts — it will ask for the user's password and may
print a "Next steps" block telling you to add `brew` to the PATH. Run
exactly the commands it tells you to.

Verify with:

```sh
brew --version
```

## Step 2 — Install Python 3.11+ and Git

```sh
brew install python@3.11 git
```

Verify:

```sh
python3.11 --version    # → Python 3.11.x
git --version           # → git version 2.x.x
```

## Step 3 — Clone the repo

Pick a stable home for it. The default below is `~/Bulletin/`; anywhere
the priest won't accidentally move it works.

```sh
cd ~
git clone <your-repo-url> Bulletin
cd Bulletin
```

If you already have the project as a folder (no `.git`), copy it in and
skip the `git clone` step. (`Update.command` won't work without a git
remote — that's fine, the priest will get updates from you instead.)

## Step 4 — Run the installer

From the `Bulletin/` directory:

```sh
./install.command
```

Or just **double-click `install.command` in Finder.** Either way it:

1. Creates `.venv/` (a private Python environment).
2. `pip install -e .` (installs all dependencies in editable mode).
3. Drops `Bulletin.command` and `Update.command` on the Desktop with
   the absolute path to this checkout baked in.

Output ends with:

```
==> Install complete.

    Double-click ~/Desktop/Bulletin.command to launch the web UI.
```

## Step 5 — Launch and verify

Double-click **Bulletin.command** on the Desktop. Two things happen:

- A Terminal window opens and shows uvicorn starting up.
- The default browser opens to `http://localhost:8765/`.

Pick a Sunday, click **Generate bulletins**, and wait for the report
page. The generated `.docx` files appear under `~/Bulletin/output/`.

To stop the server, close the Terminal window or press `Control-C` in
it. Closing the browser tab on its own does not stop it.

---

## Updating later

When new features ship:

1. The technical helper does `git push` to the repo as usual.
2. The priest double-clicks **Update.command** on their Desktop.
3. Update.command runs `git pull` + `./install.command` for them and
   prints "Update complete." when done.

If the install layout itself changes (new dependencies, etc.) the
helper may need to re-run `./install.command` once after the pull —
that's the only manual step.

---

## Files this lays down

| Path | Purpose |
|---|---|
| `~/Bulletin/` | the repo |
| `~/Bulletin/.venv/` | private Python env (never check in; covered by `.gitignore` if you ever set one up) |
| `~/Bulletin/output/` | generated `.docx` files (gitignored) |
| `~/Desktop/Bulletin.command` | starts the web UI |
| `~/Desktop/Update.command` | `git pull` + reinstall |

Nothing is installed outside `~/Bulletin/`, `~/Desktop/`, and Homebrew's
own directories.

---

## Troubleshooting

### "Bulletin.command" can't be opened — it's from an unidentified developer

This is Gatekeeper. Right-click `Bulletin.command` → **Open** → confirm.
After the first time, double-click works normally. Same workflow for
`Update.command`.

### The browser opens but says "Safari can't connect"

The server hasn't started yet. Wait a few seconds and refresh. If it
still doesn't load, check the Terminal window — there will usually be
a Python traceback that explains why.

### Port 8765 is in use

Edit `~/Desktop/Bulletin.command` to add a different port:

```sh
cd "$HERE" && exec ./.venv/bin/bulletin-ui --port 8766
```

Or set `BULLETIN_UI_PORT=8766` in the priest's shell profile.

### "command not found: brew" after Homebrew install

Homebrew prints exact lines to add to `~/.zprofile`. Run those, then
open a new Terminal window so they take effect.

### `pip install` fails with a build error for `python-docx` or `lxml`

`brew install libxml2 libxslt` and re-run `./install.command`.

### Google Sheets data isn't loading

Check `bulletin/config.py` — the sheet IDs are hard-coded there. The
sheets need to be **publicly readable** (anyone with the link). If the
priest has changed the share settings, fetches will silently return
empty data.

---

## Removing the install

To wipe everything:

```sh
rm -rf ~/Bulletin
rm -f ~/Desktop/Bulletin.command ~/Desktop/Update.command
brew uninstall python@3.11    # only if no other tools use it
```
