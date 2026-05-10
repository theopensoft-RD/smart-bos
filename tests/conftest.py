"""
Shared fixtures for the Comply Verify smoke suite.

We share ONE booted Flask app across the whole session (booting the app
takes ~5–10 s because it indexes 101 PDFs + reads the 660-row xlsx).
That's safe for read-only smoke tests; mutating tests should be
isolated explicitly.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

import pytest

# Make sure the project root is importable
ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))


@pytest.fixture(scope="session")
def gui():
    """The booted comply_verify_gui module."""
    # Default to claude_code provider; the test client never actually
    # hits Claude unless we ask for /api/claude/stream
    os.environ.setdefault("COMPLY_LLM", "claude_code")
    import comply_verify_gui as _gui  # noqa: PLC0415 — boot is slow, must be lazy

    _gui.boot()
    return _gui


@pytest.fixture(scope="session")
def client(gui):
    """Flask test client (read-only smoke surface)."""
    return gui.app.test_client()
