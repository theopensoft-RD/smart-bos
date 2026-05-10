"""
app/routes — Flask Blueprints split out from comply_verify_gui.py.

Each blueprint owns a slice of /api/* that doesn't need the legacy
in-memory globals (ROWS, PDF_INDEX, etc.) — it can talk to the DB
directly and reach paths via ``current_app.config``.

Routes that DO touch xlsx/PDF mutation state still live in the main
module (e.g. ``/api/row/apply_catalog`` writes Col D + reloads
ROWS). Move those one at a time as the gui is refactored.

Usage from comply_verify_gui.py::

    from app.routes import register_all
    register_all(app, root=ROOT, output_root=OUTPUT)
"""

from __future__ import annotations

from pathlib import Path

from flask import Flask


def register_all(app: Flask, *, root: Path, output_root: Path) -> None:
    """Wire blueprints into ``app``. Call once at boot."""
    # Stash shared paths so blueprint handlers can read them via
    # current_app.config without import gymnastics.
    app.config["COMPLY_ROOT"] = root
    app.config["COMPLY_OUTPUT"] = output_root

    from .catalog_api import bp as catalog_bp
    from .continuity_api import bp as continuity_bp
    from .export_api import bp as export_bp

    app.register_blueprint(catalog_bp)
    app.register_blueprint(export_bp)
    app.register_blueprint(continuity_bp)
