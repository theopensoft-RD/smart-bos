"""
app — application package for the Comply Verify Tool.

Submodules:
    core      — OOP domain model (Row, CatalogPDF, Project)
    database  — SQLite store (rows + audit + FTS + learning)
    learning  — HITL learning loop (suggest / record_feedback / retrain)

The Flask server lives at ../comply_verify_gui.py and imports from this
package as `from app import core, database, learning`.
"""

from . import core, database, learning  # noqa: F401

__all__ = ["core", "database", "learning"]
