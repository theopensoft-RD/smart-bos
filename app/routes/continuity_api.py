"""
continuity_api.py — Blueprint for /api/continuity.

Surfaces the latest ``_continuity/STATE_<ts>.md`` handoff document so
the UI can show a topbar banner and the Claude provider can pull it
into the system prompt.
"""

from __future__ import annotations

from datetime import datetime

from flask import Blueprint, current_app, jsonify

bp = Blueprint("continuity_api", __name__)


@bp.route("/api/continuity")
def get():
    cont_root = current_app.config["COMPLY_ROOT"] / "_continuity"
    if not cont_root.exists():
        return jsonify({"ok": True, "available": False})

    files = sorted(cont_root.glob("STATE_*.md"))
    if not files:
        return jsonify({"ok": True, "available": False})

    latest = files[-1]
    try:
        text = latest.read_text(encoding="utf-8")
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

    headline = ""
    for line in text.splitlines():
        s = line.strip()
        if s.startswith("#"):
            continue
        if s:
            headline = s
            break

    return jsonify({
        "ok": True,
        "available": True,
        "filename": latest.name,
        "mtime": datetime.fromtimestamp(latest.stat().st_mtime).isoformat(timespec="seconds"),
        "byte_size": latest.stat().st_size,
        "headline": headline[:200],
        "markdown": text,
        "history": [
            {"filename": f.name,
             "mtime": datetime.fromtimestamp(f.stat().st_mtime).isoformat(timespec="seconds"),
             "byte_size": f.stat().st_size}
            for f in reversed(files[:20])
        ],
    })
