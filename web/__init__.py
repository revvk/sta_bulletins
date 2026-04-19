"""
St. Andrew's bulletin generator — local web UI.

This package wraps the existing ``bulletin.*`` CLI tool with a local
web interface (FastAPI + HTMX) so non-technical users can generate
bulletins, manage the song library, and review post-run TODO items
without touching YAML or the terminal.

The web layer is a one-way consumer of ``bulletin.*``: nothing inside
``bulletin/`` knows the web app exists. This keeps future liturgical
work (new builders, new BCP texts) decoupled from UI concerns.
"""
