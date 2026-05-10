"""
export_api.py — Blueprint for /api/export/* (print-ready PDF package).

Reads catalogs + project info from ``app.catalog``, builds the package
via ``app.export``, writes results to ``<root>/_db/exports/`` and
serves them via /api/export/download.
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

from flask import Blueprint, current_app, jsonify, request, send_file

from .. import catalog
from .. import database as db
from .. import export as exp_mod

bp = Blueprint("export_api", __name__)


def _root() -> Path:
    return current_app.config["COMPLY_ROOT"]


def _output() -> Path:
    return current_app.config["COMPLY_OUTPUT"]


@bp.route("/api/export/preview")
def preview():
    """What would be in the package — counts only, no PDF generation."""
    proj = catalog.get_active_project()
    if not proj:
        return jsonify({"ok": False, "error": "no active project"}), 503

    mode = request.args.get("mode", "full")
    section_filter = request.args.get("section") or None
    bound_only = request.args.get("bound_only") == "1"
    include_audit = request.args.get("include_audit") == "1"

    items = catalog.list_catalogs(section=section_filter, limit=1000)
    if bound_only:
        with db.conn() as c:
            bound = {r["catalog_id"] for r in c.execute(
                "SELECT DISTINCT catalog_id FROM row_catalog_links "
                "WHERE project_id = ?", (proj["project_id"],))}
        items = [c for c in items if c["catalog_id"] in bound]
    items.sort(key=lambda c: (c.get("section_hint") or "zzz",
                                c.get("brand") or "",
                                c.get("model") or ""))

    out = {
        "ok": True,
        "project": {
            "name": proj.get("name"),
            "code": proj.get("code"),
            "company_name": proj.get("company_name"),
        },
        "mode": mode,
        "comply_pdf_present": False,
        "catalog_count": 0,
        "catalogs_per_section": {},
        "audit_count": 0,
    }

    if mode in ("full", "comply_only"):
        candidates = sorted(_output().glob("Comply spec*.pdf"))
        candidates = [p for p in candidates
                      if ".bak" not in p.name and "~$" not in p.name]
        if candidates:
            out["comply_pdf_present"] = True
            out["comply_pdf_name"] = candidates[0].name

    if mode in ("full", "catalogs_only"):
        out["catalog_count"] = len(items)
        per_section: dict[str, int] = {}
        for c in items:
            key = c.get("section_hint") or "Other"
            per_section[key] = per_section.get(key, 0) + 1
        out["catalogs_per_section"] = per_section

    if include_audit:
        with db.conn() as c:
            n = c.execute("SELECT COUNT(*) AS n FROM audit_log").fetchone()["n"]
        out["audit_count"] = min(n, 200)

    return jsonify(out)


@bp.route("/api/export/package", methods=["POST"])
def build():
    """Build the PDF package and return its download URL."""
    args = request.get_json(silent=True) or {}
    args.update(request.args.to_dict())

    mode = args.get("mode", "full")
    section_filter = args.get("section") or None
    bound_only = str(args.get("bound_only", "")) == "1"
    include_audit = str(args.get("include_audit", "")) == "1"

    proj = catalog.get_active_project()
    if not proj:
        return jsonify({"ok": False, "error": "no active project"}), 503

    output = _output()

    # Comply spec PDF (optional)
    comply_pdf_path = None
    if mode in ("full", "comply_only"):
        candidates = [p for p in sorted(output.glob("Comply spec*.pdf"))
                      if ".bak" not in p.name and "~$" not in p.name]
        comply_pdf_path = candidates[0] if candidates else None

    # Catalogs
    items: list[dict] = []
    if mode in ("full", "catalogs_only"):
        items = catalog.list_catalogs(section=section_filter, limit=2000)
        if bound_only:
            with db.conn() as c:
                bound = {r["catalog_id"] for r in c.execute(
                    "SELECT DISTINCT catalog_id FROM row_catalog_links "
                    "WHERE project_id = ?", (proj["project_id"],))}
            items = [c for c in items if c["catalog_id"] in bound]
        items.sort(key=lambda c: (c.get("section_hint") or "zzz",
                                    c.get("brand") or "",
                                    c.get("model") or ""))

    # Audit
    audit_entries: list[dict] = []
    if include_audit:
        with db.conn() as c:
            audit_entries = [dict(r) for r in c.execute(
                "SELECT * FROM audit_log ORDER BY ts DESC LIMIT 200")]

    # Output file
    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    safe = (proj.get("name") or "project").replace("/", "-").replace(" ", "_")
    filename = args.get("filename") or f"compliance-package-{safe}-{ts}.pdf"
    out_dir = _root() / "_db" / "exports"
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / filename

    # Version label = latest snapshot id (best-effort import from gui)
    version_label = ""
    try:
        # Local import to avoid cycle at module load time
        import comply_verify_gui as gui
        latest = gui.get_version_sync_status().get("latest")
        if latest:
            version_label = latest.get("id", "")
    except Exception:
        pass

    try:
        result = exp_mod.build_package(
            out_path=out_path,
            company_name=proj.get("company_name") or "Smart Solution",
            project_name=proj.get("name") or "Project",
            project_code=proj.get("code"),
            version=version_label,
            comply_pdf_path=comply_pdf_path,
            catalogs=items,
            output_root=output,
            include_audit=include_audit,
            audit_entries=audit_entries,
        )
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

    db.log_audit(action="export_package", target_type="project",
                 target_id=str(proj.get("project_id")),
                 details={"mode": mode, "filename": filename,
                          "page_count": result["page_count"],
                          "byte_size": result["byte_size"]},
                 actor="user")

    return jsonify({
        **result,
        "filename": filename,
        "download_url": f"/api/export/download?file={filename}",
    })


@bp.route("/api/export/download")
def download():
    name = request.args.get("file") or ""
    if not name or "/" in name or "\\" in name or ".." in name:
        return jsonify({"ok": False, "error": "bad filename"}), 400
    p = _root() / "_db" / "exports" / name
    if not p.exists():
        return jsonify({"ok": False, "error": "not found"}), 404
    return send_file(str(p), mimetype="application/pdf",
                     as_attachment=True, download_name=name)


@bp.route("/api/export/list")
def list_exports():
    out_dir = _root() / "_db" / "exports"
    if not out_dir.exists():
        return jsonify({"ok": True, "items": []})
    items = []
    for p in sorted(out_dir.glob("*.pdf"),
                     key=lambda x: x.stat().st_mtime, reverse=True):
        st = p.stat()
        items.append({
            "filename": p.name,
            "byte_size": st.st_size,
            "mtime": datetime.fromtimestamp(st.st_mtime).isoformat(timespec="seconds"),
            "download_url": f"/api/export/download?file={p.name}",
        })
    return jsonify({"ok": True, "items": items[:50]})
