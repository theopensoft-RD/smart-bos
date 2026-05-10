"""
catalog.py — catalog library (multi-company / multi-project support).

This module is the *additive layer* over the existing Comply Verify
system. The xlsx file remains canonical (Contract A); new tables in
``_db/comply.db`` add a searchable, editable catalog of the PDFs in
``output/`` plus a binding from project rows → catalogs.

Public surface
--------------
ingest_output_dir(output_root)              one-time / idempotent migration
list_catalogs(brand, category, section, q)  filtered listing
get_catalog(catalog_id)                     full detail + annotations
update_catalog(catalog_id, **fields)        edit metadata
search_catalogs(query, limit=20)            FTS-style match on text + meta

list_annotations(catalog_id)
add_annotation(catalog_id, page, **fields)
update_annotation(annot_id, **fields)
delete_annotation(annot_id)

bind_row_to_catalog(project_id, row, catalog_id, page, col_d)
get_row_link(project_id, row)
list_links_for_catalog(catalog_id)

Companies / projects:
list_companies(), upsert_company(name, code)
list_projects(company_id), upsert_project(...), set_active_project(project_id)
get_active_project()
"""

from __future__ import annotations

import hashlib
import json
import re
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from . import database as db

# ---------------------------------------------------------------------------
# Companies / projects
# ---------------------------------------------------------------------------

def list_companies() -> list[dict]:
    with db.conn() as c:
        return [dict(r) for r in c.execute(
            "SELECT * FROM companies ORDER BY name"
        )]


def upsert_company(name: str, code: str | None = None) -> int:
    """Insert (or update by code) and return company_id."""
    with db.conn() as c:
        if code:
            row = c.execute(
                "SELECT company_id FROM companies WHERE code=?", (code,)
            ).fetchone()
            if row:
                c.execute("UPDATE companies SET name=? WHERE company_id=?",
                          (name, row["company_id"]))
                return int(row["company_id"])
        cur = c.execute(
            "INSERT INTO companies (name, code) VALUES (?, ?)", (name, code)
        )
        return int(cur.lastrowid)


def list_projects(company_id: int | None = None) -> list[dict]:
    sql = """SELECT p.*, c.name AS company_name, c.code AS company_code
             FROM projects p JOIN companies c ON c.company_id = p.company_id"""
    params: tuple = ()
    if company_id is not None:
        sql += " WHERE p.company_id = ?"
        params = (company_id,)
    sql += " ORDER BY p.name"
    with db.conn() as c:
        return [dict(r) for r in c.execute(sql, params)]


def upsert_project(*, company_id: int, name: str, code: str | None = None,
                    xlsx_rel: str | None = None,
                    output_rel: str = "output") -> int:
    """Insert (or update by company_id+code) and return project_id."""
    with db.conn() as c:
        if code:
            row = c.execute(
                "SELECT project_id FROM projects WHERE company_id=? AND code=?",
                (company_id, code),
            ).fetchone()
            if row:
                c.execute(
                    """UPDATE projects SET name=?, xlsx_rel=?, output_rel=?,
                                            updated_at=?
                       WHERE project_id=?""",
                    (name, xlsx_rel, output_rel, _now(), row["project_id"]),
                )
                return int(row["project_id"])
        cur = c.execute(
            """INSERT INTO projects (company_id, name, code, xlsx_rel, output_rel)
               VALUES (?, ?, ?, ?, ?)""",
            (company_id, name, code, xlsx_rel, output_rel),
        )
        return int(cur.lastrowid)


def set_active_project(project_id: int) -> None:
    with db.conn() as c:
        c.execute("UPDATE projects SET is_active=0")
        c.execute("UPDATE projects SET is_active=1 WHERE project_id=?",
                  (project_id,))


def get_active_project() -> dict | None:
    with db.conn() as c:
        r = c.execute(
            """SELECT p.*, c.name AS company_name, c.code AS company_code
               FROM projects p JOIN companies c ON c.company_id = p.company_id
               WHERE p.is_active = 1 LIMIT 1"""
        ).fetchone()
    return dict(r) if r else None


# ---------------------------------------------------------------------------
# Catalog ingest (one-time migration: scan output/ → catalogs table)
# ---------------------------------------------------------------------------

# Reuse the same brand/model parsing the rule-based generator already does.
# We don't import comply_verify_gui here to avoid a circular dep — these
# heuristics are duplicated, but lightweight.

_BRAND_TOKEN_RE = re.compile(r"[A-Za-z][A-Za-z0-9-]{2,}")
_SECTION_RE = re.compile(r"^(\d+(?:\.\d+){1,4})\.?")


def _parse_section_from_path(rel: str) -> str | None:
    """Extract a section like '5.1.1.2' from a folder/file name."""
    parts = Path(rel).parts
    for p in parts:
        m = _SECTION_RE.match(p)
        if m:
            return m.group(1)
    return None


def _guess_brand_model_from_filename(stem: str) -> tuple[str | None, str | None]:
    """Best-effort: pick longest Latin token as brand, the rest as model."""
    tokens = _BRAND_TOKEN_RE.findall(stem)
    if not tokens:
        return None, None
    # Heuristic: first capitalized token is often the brand; remainder = model
    brand = tokens[0]
    model = " ".join(tokens[1:]) if len(tokens) > 1 else None
    return brand, model


def _sha256_file(path: Path, max_bytes: int = 4 * 1024 * 1024) -> str | None:
    """Hash up to first ``max_bytes`` for fast identity. Full hash is overkill."""
    try:
        h = hashlib.sha256()
        with open(path, "rb") as f:
            h.update(f.read(max_bytes))
        return h.hexdigest()
    except Exception:
        return None


def ingest_output_dir(output_root: Path, *, force: bool = False) -> dict:
    """Scan ``output_root`` for PDFs and ensure each has a catalog row.

    Idempotent: re-running won't duplicate. ``force=True`` re-extracts
    metadata even when a row already exists.

    Ops-2 (2026-05-10): fast-skip path — if a row already exists AND its
    stored ``pdf_sha256`` matches the current file's first 4 MB hash,
    skip everything (no fitz open, no metadata reparse). On 309-PDF /
    cold-cache GDrive this drops boot from ~10 s to ~0.5 s.

    Returns counts: {scanned, inserted, updated, skipped, skipped_sha}.
    """
    output_root = Path(output_root).resolve()
    if not output_root.exists():
        return {"ok": False, "error": f"{output_root} not found"}

    inserted = updated = skipped = skipped_sha = scanned = 0
    pdfs = sorted(output_root.rglob("*.pdf"))

    # Lazy import fitz only if there's at least one PDF (small speedup
    # on empty folders during tests)
    try:
        import fitz  # type: ignore
    except ImportError:
        fitz = None

    with db.conn() as c:
        # Pre-fetch existing (rel → (catalog_id, pdf_sha256)) to avoid
        # one query per file
        existing_map: dict[str, tuple[int, str | None]] = {}
        for r in c.execute("SELECT catalog_id, pdf_rel, pdf_sha256 FROM catalogs"):
            existing_map[r["pdf_rel"]] = (int(r["catalog_id"]),
                                            r["pdf_sha256"])

        for pdf in pdfs:
            scanned += 1
            try:
                rel = str(pdf.relative_to(output_root))
            except ValueError:
                continue

            existing = existing_map.get(rel)

            # Ops-2 fast path: existing row + sha matches → skip everything
            if existing and not force and existing[1]:
                cur_sha = _sha256_file(pdf)
                if cur_sha == existing[1]:
                    skipped_sha += 1
                    continue

            sha = _sha256_file(pdf)
            section = _parse_section_from_path(rel)
            brand, model = _guess_brand_model_from_filename(pdf.stem)

            pages = None
            if fitz is not None:
                try:
                    with fitz.open(pdf) as doc:
                        pages = len(doc)
                except Exception:
                    pages = None

            if existing and not force:
                skipped += 1
                continue
            if existing:
                c.execute(
                    """UPDATE catalogs SET pdf_sha256=?, pages=?, brand=?,
                                            model=?, section_hint=?,
                                            updated_at=?
                       WHERE catalog_id=?""",
                    (sha, pages, brand, model, section, _now(),
                     existing[0]),
                )
                updated += 1
            else:
                c.execute(
                    """INSERT INTO catalogs (pdf_rel, pdf_sha256, pages,
                                              brand, model, section_hint)
                       VALUES (?, ?, ?, ?, ?, ?)""",
                    (rel, sha, pages, brand, model, section),
                )
                inserted += 1
    return {"ok": True, "scanned": scanned, "inserted": inserted,
            "updated": updated, "skipped": skipped,
            "skipped_sha": skipped_sha}


# ---------------------------------------------------------------------------
# Catalog list / get / update
# ---------------------------------------------------------------------------

def list_catalogs(*, brand: str | None = None, category: str | None = None,
                   section: str | None = None, q: str | None = None,
                   archived: bool = False, limit: int = 200) -> list[dict]:
    """Filtered list. ``q`` matches any of pdf_rel/brand/model/description."""
    sql = "SELECT * FROM catalogs WHERE 1=1"
    params: list[Any] = []
    if not archived:
        sql += " AND archived = 0"
    if brand:
        sql += " AND brand = ?"; params.append(brand)
    if category:
        sql += " AND category = ?"; params.append(category)
    if section:
        # Match section or any of its descendants (5.1 matches 5.1.1.2)
        sql += " AND (section_hint = ? OR section_hint LIKE ?)"
        params += [section, f"{section}.%"]
    if q:
        like = f"%{q}%"
        sql += (" AND (pdf_rel LIKE ? OR brand LIKE ? OR model LIKE ? "
                "OR description LIKE ? OR category LIKE ?)")
        params += [like] * 5
    sql += " ORDER BY section_hint, brand, model LIMIT ?"
    params.append(limit)
    with db.conn() as c:
        return [dict(r) for r in c.execute(sql, params)]


def get_catalog(catalog_id: int) -> dict | None:
    with db.conn() as c:
        r = c.execute(
            "SELECT * FROM catalogs WHERE catalog_id = ?", (catalog_id,)
        ).fetchone()
        if not r:
            return None
        out = dict(r)
        out["annotations"] = list_annotations(catalog_id)
        out["pages_text"] = [
            {"page": pr["page"], "excerpt": (pr["text_excerpt"] or "")[:300]}
            for pr in c.execute(
                "SELECT page, text_excerpt FROM catalog_pages "
                "WHERE catalog_id = ? ORDER BY page", (catalog_id,))
        ]
    return out


_EDITABLE_FIELDS = {"brand", "model", "category", "section_hint",
                    "description", "metadata_json", "archived"}


def update_catalog(catalog_id: int, **fields) -> bool:
    bad = set(fields) - _EDITABLE_FIELDS
    if bad:
        raise ValueError(f"non-editable fields: {bad}")
    if not fields:
        return False
    sets = ", ".join(f"{k}=?" for k in fields)
    params = list(fields.values()) + [_now(), catalog_id]
    sql = f"UPDATE catalogs SET {sets}, updated_at=? WHERE catalog_id=?"
    with db.conn() as c:
        c.execute(sql, params)
    return True


# UX-2: bulk catalog metadata cleanup ---------------------------------------

def bulk_update_brand(*, match: str, new_brand: str | None,
                       match_type: str = "exact",
                       only_field: str = "brand") -> dict:
    """Find catalogs whose ``brand`` equals/contains ``match`` and rewrite
    it to ``new_brand`` (or NULL if new_brand is empty).

    match_type: "exact" | "contains" | "prefix" | "regex"
    only_field: "brand" | "model" | "category" | "section_hint"

    Returns ``{matched: int, updated: int, ids: [int]}``.
    """
    if only_field not in ("brand", "model", "category", "section_hint"):
        raise ValueError("only_field must be brand/model/category/section_hint")

    # Build the WHERE clause based on match_type
    where_sql = ""
    where_params: list[Any] = []
    if match_type == "exact":
        if match == "":
            where_sql = f"({only_field} IS NULL OR {only_field} = '')"
        else:
            where_sql = f"{only_field} = ?"
            where_params = [match]
    elif match_type == "contains":
        where_sql = f"{only_field} LIKE ?"
        where_params = [f"%{match}%"]
    elif match_type == "prefix":
        where_sql = f"{only_field} LIKE ?"
        where_params = [f"{match}%"]
    elif match_type == "regex":
        # SQLite REGEXP needs an extension; fall back to LIKE-anchored
        # patterns the caller can express, or just iterate in Python.
        # For simplicity, list all candidates then filter in Python.
        import re as _re
        try:
            _ = _re.compile(match)
        except _re.error as e:
            raise ValueError(f"bad regex: {e}") from e
        return _bulk_update_regex(match, new_brand, only_field)
    else:
        raise ValueError(f"unknown match_type: {match_type}")

    new_val = new_brand if new_brand else None
    with db.conn() as c:
        rows = c.execute(
            f"SELECT catalog_id FROM catalogs WHERE archived = 0 AND {where_sql}",
            where_params,
        ).fetchall()
        ids = [int(r["catalog_id"]) for r in rows]
        if not ids:
            return {"matched": 0, "updated": 0, "ids": []}
        # Update via parametrized IN clause
        placeholders = ",".join("?" for _ in ids)
        params = [new_val, _now()] + ids
        c.execute(
            f"UPDATE catalogs SET {only_field} = ?, updated_at = ? "
            f"WHERE catalog_id IN ({placeholders})",
            params,
        )
    return {"matched": len(ids), "updated": len(ids), "ids": ids}


def _bulk_update_regex(pattern: str, new_value: str | None,
                       field: str) -> dict:
    """Regex-based variant — pull all rows then filter in Python."""
    import re as _re
    pat = _re.compile(pattern)
    new_val = new_value if new_value else None
    with db.conn() as c:
        rows = c.execute(
            f"SELECT catalog_id, {field} AS val FROM catalogs WHERE archived = 0"
        ).fetchall()
        ids = [int(r["catalog_id"]) for r in rows
               if r["val"] is not None and pat.search(r["val"])]
        if not ids:
            return {"matched": 0, "updated": 0, "ids": []}
        placeholders = ",".join("?" for _ in ids)
        params = [new_val, _now()] + ids
        c.execute(
            f"UPDATE catalogs SET {field} = ?, updated_at = ? "
            f"WHERE catalog_id IN ({placeholders})",
            params,
        )
    return {"matched": len(ids), "updated": len(ids), "ids": ids}


def bulk_match_preview(*, match: str, match_type: str = "exact",
                        field: str = "brand", limit: int = 50) -> list[dict]:
    """Preview rows that would be matched (for the UI dialog)."""
    if field not in ("brand", "model", "category", "section_hint"):
        raise ValueError("field must be brand/model/category/section_hint")

    if match_type == "regex":
        import re as _re
        try:
            pat = _re.compile(match)
        except _re.error as e:
            raise ValueError(f"bad regex: {e}") from e
        with db.conn() as c:
            all_rows = c.execute(
                f"SELECT catalog_id, brand, model, section_hint, pdf_rel, {field} "
                f"FROM catalogs WHERE archived = 0",
            ).fetchall()
            out = [dict(r) for r in all_rows
                   if r[field] is not None and pat.search(r[field])][:limit]
            return out

    if match_type == "exact":
        if match == "":
            where_sql = f"({field} IS NULL OR {field} = '')"
            params: list[Any] = []
        else:
            where_sql = f"{field} = ?"; params = [match]
    elif match_type == "contains":
        where_sql = f"{field} LIKE ?"; params = [f"%{match}%"]
    elif match_type == "prefix":
        where_sql = f"{field} LIKE ?"; params = [f"{match}%"]
    else:
        raise ValueError(f"unknown match_type: {match_type}")

    with db.conn() as c:
        rows = c.execute(
            f"SELECT catalog_id, brand, model, section_hint, pdf_rel "
            f"FROM catalogs WHERE archived = 0 AND {where_sql} "
            f"LIMIT ?", params + [limit],
        ).fetchall()
    return [dict(r) for r in rows]


# ---------------------------------------------------------------------------
# Catalog annotations (DB-stored)
# ---------------------------------------------------------------------------

def list_annotations(catalog_id: int, page: int | None = None) -> list[dict]:
    sql = ("SELECT * FROM catalog_annotations WHERE catalog_id = ? "
           "AND archived = 0")
    params: list[Any] = [catalog_id]
    if page is not None:
        sql += " AND page = ?"; params.append(page)
    sql += " ORDER BY page, annot_id"
    with db.conn() as c:
        out = []
        for r in c.execute(sql, params):
            d = dict(r)
            try:
                d["rect"] = json.loads(d.pop("rect_json")) if d.get("rect_json") else None
            except Exception:
                d["rect"] = None
            try:
                d["color"] = json.loads(d.pop("color_json")) if d.get("color_json") else None
            except Exception:
                d["color"] = None
            out.append(d)
        return out


def add_annotation(*, catalog_id: int, page: int, type: str,
                    rect: list[float], contents: str = "",
                    color: list[float] | None = None,
                    border_width: float = 1.0,
                    anchor_text: str | None = None) -> int:
    if type not in ("Square", "FreeText"):
        raise ValueError("type must be Square or FreeText")
    with db.conn() as c:
        cur = c.execute(
            """INSERT INTO catalog_annotations
               (catalog_id, page, type, rect_json, contents, color_json,
                border_width, anchor_text)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (catalog_id, page, type, json.dumps(rect), contents,
             json.dumps(color) if color else None,
             border_width, anchor_text),
        )
        return int(cur.lastrowid)


def update_annotation(annot_id: int, **fields) -> bool:
    allowed = {"page", "type", "rect", "contents", "color", "border_width",
               "anchor_text", "archived"}
    bad = set(fields) - allowed
    if bad:
        raise ValueError(f"unknown fields: {bad}")
    if not fields:
        return False
    # Convert structured fields to *_json
    if "rect" in fields:
        fields["rect_json"] = json.dumps(fields.pop("rect"))
    if "color" in fields:
        c = fields.pop("color")
        fields["color_json"] = json.dumps(c) if c else None
    sets = ", ".join(f"{k}=?" for k in fields)
    params = list(fields.values()) + [_now(), annot_id]
    sql = f"UPDATE catalog_annotations SET {sets}, updated_at=? WHERE annot_id=?"
    with db.conn() as c:
        c.execute(sql, params)
    return True


def delete_annotation(annot_id: int) -> bool:
    """Soft-delete (archived=1)."""
    with db.conn() as c:
        c.execute(
            "UPDATE catalog_annotations SET archived=1, updated_at=? "
            "WHERE annot_id=?", (_now(), annot_id))
    return True


# ---------------------------------------------------------------------------
# Row ↔ catalog binding
# ---------------------------------------------------------------------------

def bind_row_to_catalog(*, project_id: int, row_num: int, catalog_id: int,
                         page: int | None = None,
                         col_d_text: str | None = None) -> None:
    with db.conn() as c:
        c.execute(
            """INSERT INTO row_catalog_links
                  (project_id, row_num, catalog_id, page, col_d_text)
               VALUES (?, ?, ?, ?, ?)
               ON CONFLICT(project_id, row_num) DO UPDATE SET
                  catalog_id=excluded.catalog_id,
                  page=excluded.page,
                  col_d_text=excluded.col_d_text,
                  bound_at=CURRENT_TIMESTAMP""",
            (project_id, row_num, catalog_id, page, col_d_text),
        )


def get_row_link(project_id: int, row_num: int) -> dict | None:
    with db.conn() as c:
        r = c.execute(
            """SELECT l.*, c.brand, c.model, c.pdf_rel, c.section_hint
               FROM row_catalog_links l
               JOIN catalogs c USING(catalog_id)
               WHERE l.project_id=? AND l.row_num=?""",
            (project_id, row_num),
        ).fetchone()
    return dict(r) if r else None


def list_links_for_catalog(catalog_id: int) -> list[dict]:
    """Which rows in which projects currently use this catalog?"""
    with db.conn() as c:
        return [dict(r) for r in c.execute(
            """SELECT l.*, p.name AS project_name, p.code AS project_code,
                      co.name AS company_name
               FROM row_catalog_links l
               JOIN projects p USING(project_id)
               JOIN companies co USING(company_id)
               WHERE l.catalog_id = ?
               ORDER BY p.name, l.row_num""", (catalog_id,))]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def stats() -> dict:
    """Quick counts for the UI badge."""
    with db.conn() as c:
        d = {}
        d["companies"] = c.execute("SELECT COUNT(*) FROM companies").fetchone()[0]
        d["projects"] = c.execute("SELECT COUNT(*) FROM projects").fetchone()[0]
        d["catalogs"] = c.execute(
            "SELECT COUNT(*) FROM catalogs WHERE archived = 0").fetchone()[0]
        d["annotations"] = c.execute(
            "SELECT COUNT(*) FROM catalog_annotations WHERE archived = 0").fetchone()[0]
        d["row_links"] = c.execute(
            "SELECT COUNT(*) FROM row_catalog_links").fetchone()[0]
        # Catalog by brand (top 5)
        d["top_brands"] = [
            {"brand": r["brand"] or "(unknown)", "count": r["n"]}
            for r in c.execute(
                """SELECT brand, COUNT(*) AS n FROM catalogs
                   WHERE archived = 0 AND brand IS NOT NULL
                   GROUP BY brand ORDER BY n DESC LIMIT 5""")
        ]
    return d
