"""
catalog_api.py — Blueprint for /api/catalogs/* + /api/companies/*
                  + /api/projects/* endpoints.

Pure CRUD over ``app.catalog`` + ``app.database``. No xlsx mutation,
no in-memory ROWS dependency. Lives here so the comply_verify_gui
monolith stays focused on the orchestration glue.
"""

from __future__ import annotations

from flask import Blueprint, current_app, jsonify, request

from .. import catalog
from .. import database as db

bp = Blueprint("catalog_api", __name__)


@bp.route("/api/catalogs")
def list_():
    """List catalogs with optional filters: brand/category/section/q/archived."""
    archived = request.args.get("archived") == "1"
    items = catalog.list_catalogs(
        brand=request.args.get("brand") or None,
        category=request.args.get("category") or None,
        section=request.args.get("section") or None,
        q=request.args.get("q") or None,
        archived=archived,
        limit=int(request.args.get("limit", 200)),
    )
    return jsonify({"ok": True, "items": items, "count": len(items)})


@bp.route("/api/catalogs/stats")
def stats():
    return jsonify(catalog.stats())


@bp.route("/api/catalogs/<int:catalog_id>")
def get(catalog_id: int):
    item = catalog.get_catalog(catalog_id)
    if not item:
        return jsonify({"ok": False, "error": "not found"}), 404
    return jsonify({"ok": True, "catalog": item})


@bp.route("/api/catalogs/<int:catalog_id>", methods=["PATCH"])
def update(catalog_id: int):
    if not catalog.get_catalog(catalog_id):
        return jsonify({"ok": False, "error": "not found"}), 404
    data = request.get_json(silent=True) or {}
    try:
        catalog.update_catalog(catalog_id, **data)
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    db.log_audit(action="catalog_update", target_type="catalog",
                 target_id=str(catalog_id), details=data, actor="user")
    return jsonify({"ok": True, "catalog": catalog.get_catalog(catalog_id)})


@bp.route("/api/catalogs/<int:catalog_id>/links")
def links(catalog_id: int):
    return jsonify({"ok": True,
                    "links": catalog.list_links_for_catalog(catalog_id)})


@bp.route("/api/catalogs/reingest", methods=["POST"])
def reingest():
    force = request.args.get("force") == "1"
    output = current_app.config["COMPLY_OUTPUT"]
    return jsonify(catalog.ingest_output_dir(output, force=force))


# ── Catalog annotations (DB-stored) ─────────────────────────────────

@bp.route("/api/catalogs/<int:catalog_id>/annotations")
def annots_list(catalog_id: int):
    page = request.args.get("page")
    page_n = int(page) if page else None
    return jsonify({"ok": True,
                    "annotations": catalog.list_annotations(catalog_id, page=page_n)})


@bp.route("/api/catalogs/<int:catalog_id>/annotations", methods=["POST"])
def annots_add(catalog_id: int):
    data = request.get_json(silent=True) or {}
    try:
        annot_id = catalog.add_annotation(
            catalog_id=catalog_id,
            page=int(data["page"]),
            type=data["type"],
            rect=data["rect"],
            contents=data.get("contents", ""),
            color=data.get("color"),
            border_width=float(data.get("border_width", 1.0)),
            anchor_text=data.get("anchor_text"),
        )
    except (ValueError, KeyError) as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    db.log_audit(action="catalog_annot_add", target_type="catalog",
                 target_id=str(catalog_id),
                 details={"annot_id": annot_id, "page": data.get("page")},
                 actor="user")
    return jsonify({"ok": True, "annot_id": annot_id})


@bp.route("/api/catalogs/<int:catalog_id>/annotations/<int:annot_id>",
          methods=["PATCH"])
def annots_update(catalog_id: int, annot_id: int):
    data = request.get_json(silent=True) or {}
    try:
        catalog.update_annotation(annot_id, **data)
    except ValueError as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    return jsonify({"ok": True})


@bp.route("/api/catalogs/<int:catalog_id>/annotations/<int:annot_id>",
          methods=["DELETE"])
def annots_delete(catalog_id: int, annot_id: int):
    catalog.delete_annotation(annot_id)
    return jsonify({"ok": True})


# ── Companies / projects ────────────────────────────────────────────

@bp.route("/api/companies")
def companies():
    return jsonify({"ok": True, "items": catalog.list_companies()})


@bp.route("/api/companies", methods=["POST"])
def companies_add():
    data = request.get_json(silent=True) or {}
    name = (data.get("name") or "").strip()
    if not name:
        return jsonify({"ok": False, "error": "name required"}), 400
    cid = catalog.upsert_company(name=name, code=data.get("code"))
    return jsonify({"ok": True, "company_id": cid})


@bp.route("/api/projects")
def projects():
    company_id = request.args.get("company_id")
    cid = int(company_id) if company_id else None
    return jsonify({"ok": True,
                    "items": catalog.list_projects(company_id=cid),
                    "active": catalog.get_active_project()})


@bp.route("/api/projects", methods=["POST"])
def projects_add():
    data = request.get_json(silent=True) or {}
    try:
        pid = catalog.upsert_project(
            company_id=int(data["company_id"]),
            name=data["name"],
            code=data.get("code"),
            xlsx_rel=data.get("xlsx_rel"),
            output_rel=data.get("output_rel", "output"),
        )
    except (KeyError, ValueError) as e:
        return jsonify({"ok": False, "error": str(e)}), 400
    return jsonify({"ok": True, "project_id": pid})


@bp.route("/api/projects/<int:project_id>/activate", methods=["POST"])
def projects_activate(project_id: int):
    catalog.set_active_project(project_id)
    return jsonify({"ok": True, "active": catalog.get_active_project()})
