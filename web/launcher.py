"""
Console-script entry point for the local bulletin web UI.

`pip install -e .` registers this as the ``bulletin-ui`` command, which
is what the Desktop ``Bulletin.command`` shim runs. It does three
things:

  1. Starts a uvicorn worker bound to ``127.0.0.1:8765`` (loopback only
     — never exposed to the LAN unless the operator explicitly passes
     ``--host 0.0.0.0``).
  2. Opens the user's default browser to the same URL after a brief
     pause to give the server a chance to come up.
  3. Hands control to uvicorn so the process stays alive until the user
     closes the Terminal window or hits Ctrl-C.

This is *not* meant to be a production server. Single user, local
machine, no auth — same trust boundary as running ``generate.py``
from a shell.
"""

from __future__ import annotations

import argparse
import os
import threading
import time
import webbrowser


DEFAULT_HOST = "127.0.0.1"
DEFAULT_PORT = 8765


def _open_browser_when_ready(url: str, delay: float = 0.8) -> None:
    """Open the default browser after a short delay so the server is up."""
    def opener() -> None:
        time.sleep(delay)
        webbrowser.open(url)
    t = threading.Thread(target=opener, daemon=True)
    t.start()


def main() -> None:
    parser = argparse.ArgumentParser(
        prog="bulletin-ui",
        description="Launch the St. Andrew's bulletin web UI on this Mac.",
    )
    parser.add_argument("--host", default=DEFAULT_HOST,
                        help=f"Bind address (default: {DEFAULT_HOST}). "
                             "Use 0.0.0.0 to expose on the LAN — only do "
                             "that on a trusted network.")
    parser.add_argument("--port", type=int, default=DEFAULT_PORT,
                        help=f"Port to listen on (default: {DEFAULT_PORT})")
    parser.add_argument("--no-browser", action="store_true",
                        help="Don't auto-open the browser. Useful when "
                             "running the server in the background.")
    parser.add_argument("--reload", action="store_true",
                        help="Reload on code change (development only).")
    args = parser.parse_args()

    # Honor BULLETIN_UI_PORT / BULLETIN_UI_HOST env overrides if set —
    # makes it easy to switch ports without re-editing the .command file.
    host = os.environ.get("BULLETIN_UI_HOST", args.host)
    port = int(os.environ.get("BULLETIN_UI_PORT", args.port))

    url = f"http://{host}:{port}/"
    print()
    print(f"  St. Andrew's Bulletin Generator")
    print(f"  Open in your browser: {url}")
    print(f"  Press Control-C to quit.")
    print()

    if not args.no_browser:
        _open_browser_when_ready(url)

    # Imported here (not at module top) so `bulletin-ui --help` is fast
    # and doesn't blow up if uvicorn isn't installed yet.
    import uvicorn
    uvicorn.run(
        "web.app:app",
        host=host,
        port=port,
        reload=args.reload,
        log_level="info",
    )


if __name__ == "__main__":
    main()
