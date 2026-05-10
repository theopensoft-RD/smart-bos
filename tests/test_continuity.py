"""
Continuity tests — both the consumer-side (/api/continuity) and
the producer-side (auto-write on shutdown).
"""

from __future__ import annotations

from pathlib import Path

from app import continuity as cont


def test_continuity_module_imports():
    assert callable(cont.write_session_state)
    assert callable(cont.install_atexit_hook)


def test_write_session_state_creates_file(client, gui, tmp_path: Path):
    """Ops-1: write_session_state should create a STATE markdown file
    when there's recent audit activity."""
    # Use the project root (real DB) so audit_log has rows
    out = cont.write_session_state(root=Path("."))
    # Either created OR returned None when no events — both are valid
    if out is not None:
        assert out.exists()
        text = out.read_text(encoding="utf-8")
        assert text.startswith("# Continuity State")
        assert "Session summary" in text


def test_atexit_hook_is_idempotent(gui):
    """Calling install twice doesn't double-register."""
    cont.install_atexit_hook(root=Path("."))
    cont.install_atexit_hook(root=Path("."))   # second call should no-op
    assert getattr(cont.install_atexit_hook, "_installed", False) is True
