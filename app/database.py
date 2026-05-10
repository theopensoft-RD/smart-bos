"""
comply_db.py — SQLite-backed data layer for the Comply Verify tool.

The xlsx + filesystem (TOR pdf, catalog pdfs, snapshots) remain the canonical
source of truth.  This module mirrors them into a queryable database that
lives alongside the project (``_db/comply.db``) so we can:

  • answer "show me every row in section 5.1.2 with status≠pass" with a
    single SQL query instead of scanning ROWS in Python on every request
  • keep verification status, audit trail, and auto-annotate plan history
    in one place that survives restarts and works across machines via
    Google Drive
  • provide full-text search over Col B/C/D content via SQLite FTS5
  • record an immutable audit log of every change (status edit, auto-annotate
    apply, PDF edit, snapshot, restore) so users can see what happened when

Population is one-way (filesystem → DB) on boot and on `/api/refresh`.
Writes-back (status, audit, plans) flow through this module.
"""

from __future__ import annotations

import json
import os
import sqlite3
import sys
import time
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path

DB_VERSION = 3  # bumped 2026-05-10: submissions table (multi-bidder per project)

_DB_PATH: Path | None = None


# ---------------------------------------------------------------------------
# Schema
# ---------------------------------------------------------------------------

SCHEMA = """
CREATE TABLE IF NOT EXISTS schema_version (
    version    INTEGER PRIMARY KEY,
    applied_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- Rows from comply.xlsx (mirrored, refreshed on boot/refresh) ---------------
CREATE TABLE IF NOT EXISTS rows (
    row_num            INTEGER PRIMARY KEY,
    col_a              TEXT,
    col_b              TEXT,
    col_c              TEXT,
    col_d              TEXT,
    col_e              TEXT,
    col_f              TEXT,
    section            TEXT,
    section_inherited  INTEGER DEFAULT 0,
    pdf_rel            TEXT,
    pdf_inherited      INTEGER DEFAULT 0,
    needs_col_d        INTEGER DEFAULT 0,
    parsed_type        TEXT,
    parsed_brand       TEXT,
    parsed_model       TEXT,
    parsed_ref         TEXT,
    parsed_page        INTEGER,
    parsed_item        INTEGER,
    parsed_subitem     INTEGER,
    last_synced_at     TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
CREATE INDEX IF NOT EXISTS idx_rows_section ON rows(section);
CREATE INDEX IF NOT EXISTS idx_rows_pdf     ON rows(pdf_rel);
CREATE INDEX IF NOT EXISTS idx_rows_type    ON rows(parsed_type);
CREATE INDEX IF NOT EXISTS idx_rows_vendor  ON rows(col_e);

-- Catalog PDFs --------------------------------------------------------------
CREATE TABLE IF NOT EXISTS pdfs (
    pdf_id            INTEGER PRIMARY KEY AUTOINCREMENT,
    rel_path          TEXT UNIQUE NOT NULL,
    folder_key        TEXT,
    section_prefix    TEXT,
    size              INTEGER,
    mtime             REAL,
    num_pages         INTEGER,
    detected_brand    TEXT,
    detected_model    TEXT,
    last_indexed_at   TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
CREATE INDEX IF NOT EXISTS idx_pdfs_folder  ON pdfs(folder_key);
CREATE INDEX IF NOT EXISTS idx_pdfs_section ON pdfs(section_prefix);

-- Annotations parsed from each PDF ------------------------------------------
CREATE TABLE IF NOT EXISTS pdf_annotations (
    annot_id        INTEGER PRIMARY KEY AUTOINCREMENT,
    pdf_id          INTEGER NOT NULL REFERENCES pdfs(pdf_id) ON DELETE CASCADE,
    page_num        INTEGER NOT NULL,
    xref            INTEGER,
    annot_type      TEXT,
    rect_x0         REAL, rect_y0 REAL, rect_x1 REAL, rect_y1 REAL,
    contents        TEXT,
    is_inline       INTEGER DEFAULT 0,
    last_indexed_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
CREATE INDEX IF NOT EXISTS idx_annots_pdf_page ON pdf_annotations(pdf_id, page_num);

-- TOR section index ---------------------------------------------------------
CREATE TABLE IF NOT EXISTS tor_sections (
    section          TEXT PRIMARY KEY,
    page_num         INTEGER NOT NULL,
    last_indexed_at  TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- TOR per-page normalized text (for substring search) -----------------------
CREATE TABLE IF NOT EXISTS tor_pages (
    page_num         INTEGER PRIMARY KEY,
    normalized_text  TEXT NOT NULL,
    indexed_at       TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- Cached TOR row→page lookups ----------------------------------------------
CREATE TABLE IF NOT EXISTS tor_row_matches (
    row_num         INTEGER PRIMARY KEY REFERENCES rows(row_num) ON DELETE CASCADE,
    page_num        INTEGER,
    rects_json      TEXT,
    hits            INTEGER,
    needle          TEXT,
    cached_at       TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- Verification status (replaces verification_status.json) -------------------
CREATE TABLE IF NOT EXISTS verification_status (
    row_num     INTEGER PRIMARY KEY REFERENCES rows(row_num) ON DELETE CASCADE,
    status      TEXT NOT NULL,
    notes       TEXT,
    updated_at  TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_by  TEXT
);
CREATE INDEX IF NOT EXISTS idx_status_status ON verification_status(status);

-- Project-level snapshots (mirrors _versions/snapshots/) --------------------
CREATE TABLE IF NOT EXISTS snapshots (
    snapshot_id   TEXT PRIMARY KEY,
    tag           TEXT,
    kind          TEXT,
    timestamp     TEXT,
    size          INTEGER,
    n_files       INTEGER,
    n_output      INTEGER,
    has_tarball   INTEGER DEFAULT 0,
    last_seen_at  TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- Per-PDF edit history (mirrors output/_pdf_history/) -----------------------
CREATE TABLE IF NOT EXISTS pdf_history (
    history_id        INTEGER PRIMARY KEY AUTOINCREMENT,
    pdf_rel           TEXT NOT NULL,
    snapshot_filename TEXT NOT NULL,
    ts                TIMESTAMP,
    tag               TEXT,
    size              INTEGER,
    UNIQUE (pdf_rel, snapshot_filename)
);
CREATE INDEX IF NOT EXISTS idx_pdf_history_rel ON pdf_history(pdf_rel);

-- Auto-annotate plan history (preview + apply outcomes) ---------------------
CREATE TABLE IF NOT EXISTS auto_annotate_plans (
    plan_id          INTEGER PRIMARY KEY AUTOINCREMENT,
    row_num          INTEGER NOT NULL REFERENCES rows(row_num) ON DELETE CASCADE,
    proposed_c       TEXT,
    proposed_d       TEXT,
    annotations_json TEXT,
    role             TEXT,
    warnings_json    TEXT,
    generated_at     TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    applied_at       TIMESTAMP,
    apply_result     TEXT
);
CREATE INDEX IF NOT EXISTS idx_plans_row ON auto_annotate_plans(row_num);
CREATE INDEX IF NOT EXISTS idx_plans_applied ON auto_annotate_plans(applied_at);

-- Immutable audit log: every meaningful change ------------------------------
CREATE TABLE IF NOT EXISTS audit_log (
    log_id        INTEGER PRIMARY KEY AUTOINCREMENT,
    ts            TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    action        TEXT NOT NULL,
    target_type   TEXT,
    target_id     TEXT,
    details_json  TEXT,
    before_value  TEXT,
    after_value   TEXT,
    actor         TEXT
);
CREATE INDEX IF NOT EXISTS idx_audit_ts     ON audit_log(ts);
CREATE INDEX IF NOT EXISTS idx_audit_action ON audit_log(action);
CREATE INDEX IF NOT EXISTS idx_audit_target ON audit_log(target_type, target_id);

-- Full-text search over rows (Col B/C/D + section) -------------------------
CREATE VIRTUAL TABLE IF NOT EXISTS rows_fts USING fts5(
    row_num UNINDEXED, section, col_a, col_b, col_c, col_d, col_e,
    tokenize = 'unicode61'
);

-- HITL learning loop ========================================================
-- learning_feedback: append-only record of every (Core suggestion → user
-- action) pair. The retrain step distills these into learned_patterns that
-- get re-applied to future suggestions. Implements the "learn from
-- corrections" half of the HITL contract.
CREATE TABLE IF NOT EXISTS learning_feedback (
    fb_id              INTEGER PRIMARY KEY AUTOINCREMENT,
    ts                 TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    row_num            INTEGER,
    section            TEXT,

    -- Input context (what generator saw)
    input_b            TEXT,
    input_pdf_rel      TEXT,
    input_role         TEXT,
    input_filename     TEXT,

    -- Core's suggestion
    suggested_c        TEXT,
    suggested_d        TEXT,
    suggested_annots   TEXT,            -- JSON
    confidence         REAL,
    generator          TEXT,            -- "rules-v1" / "rules+pattern-v1" / "llm-…"
    provenance         TEXT,            -- JSON: which patterns/rules contributed

    -- What the user did
    user_action        TEXT,            -- accepted / edited / rejected
    final_c            TEXT,
    final_d            TEXT,
    final_annots       TEXT,            -- JSON

    -- For analysis / retrain
    edit_distance_d    INTEGER,
    correction_kind    TEXT             -- format / page / brand / model / vendor / multi
);
CREATE INDEX IF NOT EXISTS idx_fb_row     ON learning_feedback(row_num);
CREATE INDEX IF NOT EXISTS idx_fb_section ON learning_feedback(section);
CREATE INDEX IF NOT EXISTS idx_fb_action  ON learning_feedback(user_action);
CREATE INDEX IF NOT EXISTS idx_fb_kind    ON learning_feedback(correction_kind);

-- learned_patterns: distilled rules extracted from feedback. retrain()
-- updates this table by mining learning_feedback for repeating corrections.
-- Each pattern carries a confidence (samples_correct / samples_total) and
-- can be enabled/disabled by the user.
CREATE TABLE IF NOT EXISTS learned_patterns (
    pattern_id        INTEGER PRIMARY KEY AUTOINCREMENT,
    pattern_type      TEXT NOT NULL,    -- e.g. "filename_brand", "section_vendor",
                                        --      "row_format_d", "annot_position"
    trigger_key       TEXT NOT NULL,    -- the matching key (token / section / etc.)
    trigger_extra     TEXT,             -- optional secondary key (JSON)
    output_value      TEXT,             -- the corrected value to apply
    samples_total     INTEGER DEFAULT 1,
    samples_correct   INTEGER DEFAULT 1,
    confidence        REAL DEFAULT 1.0, -- samples_correct / samples_total
    enabled           INTEGER DEFAULT 1,
    created_at        TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    last_used_at      TIMESTAMP,
    note              TEXT,
    UNIQUE (pattern_type, trigger_key, trigger_extra)
);
CREATE INDEX IF NOT EXISTS idx_lp_type    ON learned_patterns(pattern_type);
CREATE INDEX IF NOT EXISTS idx_lp_enabled ON learned_patterns(enabled);

-- llm_calls: every Claude (or other LLM) API call's metrics, used for
-- budget enforcement, cost analytics, and quality A/B over time.
CREATE TABLE IF NOT EXISTS llm_calls (
    call_id          INTEGER PRIMARY KEY AUTOINCREMENT,
    ts               TIMESTAMP NOT NULL,
    row_num          INTEGER,
    model            TEXT,                  -- e.g. "claude-sonnet-4-5-20250929"
    stop_reason      TEXT,                  -- end_turn / tool_use / max_tokens / ...
    input_tokens     INTEGER DEFAULT 0,
    output_tokens    INTEGER DEFAULT 0,
    cache_write_tokens INTEGER DEFAULT 0,
    cache_read_tokens  INTEGER DEFAULT 0,
    cost_usd         REAL DEFAULT 0,
    elapsed_ms       INTEGER DEFAULT 0,
    tool_calls_json  TEXT,                  -- JSON array of {name, input}
    response_text    TEXT,                  -- truncated prose, if any
    prompt_size_chars INTEGER DEFAULT 0
);
CREATE INDEX IF NOT EXISTS idx_llm_ts    ON llm_calls(ts);
CREATE INDEX IF NOT EXISTS idx_llm_row   ON llm_calls(row_num);
CREATE INDEX IF NOT EXISTS idx_llm_model ON llm_calls(model);

-- ─────────────────────────────────────────────────────────────────────────
-- Multi-company / catalog library (Phase 2 — added 2026-05-10)
-- ─────────────────────────────────────────────────────────────────────────
-- Companies own projects; projects use catalogs from the shared library.
-- Catalogs are reusable across projects (same product datasheet for many
-- compliance assignments). Annotations live in the DB so the user can
-- edit them per-catalog without re-baking the PDF every time.

CREATE TABLE IF NOT EXISTS companies (
    company_id  INTEGER PRIMARY KEY AUTOINCREMENT,
    name        TEXT NOT NULL,
    code        TEXT UNIQUE,                  -- short slug, e.g. "PATTAYA"
    created_at  TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS projects (
    project_id  INTEGER PRIMARY KEY AUTOINCREMENT,
    company_id  INTEGER NOT NULL,
    name        TEXT NOT NULL,                -- "Smart Plant 1"
    code        TEXT,                          -- "SP1"
    xlsx_rel    TEXT,                          -- relative to project root
    output_rel  TEXT,                          -- "output" by default
    is_active   INTEGER NOT NULL DEFAULT 0,    -- one project at a time gets =1
    created_at  TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at  TIMESTAMP,
    FOREIGN KEY (company_id) REFERENCES companies(company_id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_projects_company ON projects(company_id);
CREATE INDEX IF NOT EXISTS idx_projects_active  ON projects(is_active);

CREATE TABLE IF NOT EXISTS catalogs (
    catalog_id    INTEGER PRIMARY KEY AUTOINCREMENT,
    pdf_rel       TEXT NOT NULL UNIQUE,        -- relative to OUTPUT root
    pdf_sha256    TEXT,                         -- de-dup key (NULL during ingest)
    pages         INTEGER,
    -- Curated metadata (editable in UI; populated by ingest with best-effort)
    brand         TEXT,
    model         TEXT,
    category      TEXT,                         -- "Server", "Switch", "Rack", ...
    section_hint  TEXT,                         -- inferred section "5.1.1.2"
    description   TEXT,
    -- Free-form vendor specs / notes
    metadata_json TEXT,                         -- JSON dict
    -- Soft-delete + audit
    archived      INTEGER NOT NULL DEFAULT 0,
    created_at    TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at    TIMESTAMP
);
CREATE INDEX IF NOT EXISTS idx_catalogs_brand    ON catalogs(brand);
CREATE INDEX IF NOT EXISTS idx_catalogs_category ON catalogs(category);
CREATE INDEX IF NOT EXISTS idx_catalogs_section  ON catalogs(section_hint);
CREATE INDEX IF NOT EXISTS idx_catalogs_archived ON catalogs(archived);

-- Per-page text excerpt for FTS-style search
CREATE TABLE IF NOT EXISTS catalog_pages (
    catalog_id  INTEGER NOT NULL,
    page        INTEGER NOT NULL,
    text_excerpt TEXT,
    PRIMARY KEY (catalog_id, page),
    FOREIGN KEY (catalog_id) REFERENCES catalogs(catalog_id) ON DELETE CASCADE
);

-- Catalog-level annotations stored in DB (editable independent of PDF baking).
-- These are the "templates" — when applied to a project row, they're written
-- through to the actual PDF via the existing apply_pdf_edits flow.
CREATE TABLE IF NOT EXISTS catalog_annotations (
    annot_id      INTEGER PRIMARY KEY AUTOINCREMENT,
    catalog_id    INTEGER NOT NULL,
    page          INTEGER NOT NULL,
    type          TEXT NOT NULL,                -- "Square" | "FreeText"
    rect_json     TEXT NOT NULL,                -- "[x0,y0,x1,y1]"
    contents      TEXT,
    color_json    TEXT,                          -- "[r,g,b]" 0..1; default red
    border_width  REAL DEFAULT 1.0,
    -- Anchor metadata (Phase C: lets annot survive PDF page reflow)
    anchor_text   TEXT,
    -- Soft-delete
    archived      INTEGER NOT NULL DEFAULT 0,
    created_at    TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at    TIMESTAMP,
    FOREIGN KEY (catalog_id) REFERENCES catalogs(catalog_id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_catalog_annots ON catalog_annotations(catalog_id, page);

-- Linkage: which catalog (and page) is currently bound to each project row.
-- Composite PK so each (project, row) has exactly one active link.
--
-- 2026-05-10 (v3): the practical scope is per-SUBMISSION, not per-project.
-- A single project (Smart Plant 1) can have multiple submissions
-- (TRIO_SR_Solution, Take_IT) where the same row gets DIFFERENT
-- catalog bindings. The PK is now (submission_id, row_num); the
-- legacy project_id column stays for backward compat. New writes
-- always set submission_id; reads should prefer submission_id when
-- the active submission is set.
CREATE TABLE IF NOT EXISTS row_catalog_links (
    project_id     INTEGER NOT NULL,
    row_num        INTEGER NOT NULL,
    catalog_id     INTEGER NOT NULL,
    page           INTEGER,                       -- which page anchors this row
    col_d_text     TEXT,                           -- generated Col D string
    bound_at       TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    submission_id  INTEGER,                        -- NEW v3: per-bidder scope
    PRIMARY KEY (project_id, row_num),
    FOREIGN KEY (project_id) REFERENCES projects(project_id) ON DELETE CASCADE,
    FOREIGN KEY (catalog_id) REFERENCES catalogs(catalog_id)
);
CREATE INDEX IF NOT EXISTS idx_row_links_catalog ON row_catalog_links(catalog_id);
CREATE INDEX IF NOT EXISTS idx_row_links_submission ON row_catalog_links(submission_id);

-- Submissions (NEW v3, 2026-05-10) ──────────────────────────────────
-- A "submission" is a specific bidder/consortium's response to a
-- project's TOR. Same project, multiple submissions = multiple
-- vendor proposals — each with its own xlsx + catalog bindings.
-- E.g.: Smart Plant 1 has TRIO_SR_Solution and Take_IT submissions.
--
-- Each submission lives in a subdirectory of output/ with its own
-- "Comply spec*.xlsx" file. Auto-discovered at boot.

CREATE TABLE IF NOT EXISTS submissions (
    submission_id  INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id     INTEGER NOT NULL,
    name           TEXT NOT NULL,                 -- "TRIO_SR_Solution"
    code           TEXT,                           -- "TRIO" (slug for display)
    output_subdir  TEXT NOT NULL,                  -- "TRIO_SR_Solution" relative to output/
    xlsx_rel       TEXT NOT NULL,                  -- "TRIO_SR_Solution/Comply spec ... TRIO.xlsx"
    is_active      INTEGER NOT NULL DEFAULT 0,     -- exactly 1 active per project
    notes          TEXT,
    created_at     TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at     TIMESTAMP,
    FOREIGN KEY (project_id) REFERENCES projects(project_id) ON DELETE CASCADE,
    UNIQUE (project_id, name)
);
CREATE INDEX IF NOT EXISTS idx_submissions_project ON submissions(project_id);
CREATE INDEX IF NOT EXISTS idx_submissions_active ON submissions(is_active);
"""


# ---------------------------------------------------------------------------
# Connection management
# ---------------------------------------------------------------------------

def init_db(db_path: str | Path) -> None:
    """Create the database file (if missing) and apply the schema."""
    global _DB_PATH
    _DB_PATH = Path(db_path)
    _DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    with _connect() as c:
        # ── Pre-migrate v2 → v3 BEFORE executescript ─────────────────
        # CREATE INDEX in SCHEMA references row_catalog_links.submission_id
        # which doesn't exist on a v2-era DB. Add the column first so the
        # CREATE INDEX inside executescript() succeeds.
        try:
            tbl = c.execute(
                "SELECT name FROM sqlite_master WHERE type='table' "
                "AND name='row_catalog_links'").fetchone()
            if tbl:
                cols = {r["name"] for r in c.execute(
                    "PRAGMA table_info(row_catalog_links)")}
                if "submission_id" not in cols:
                    c.execute(
                        "ALTER TABLE row_catalog_links "
                        "ADD COLUMN submission_id INTEGER")
        except Exception:
            pass

        c.executescript(SCHEMA)
        c.execute("INSERT OR IGNORE INTO schema_version (version) VALUES (?)",
                  (DB_VERSION,))


def db_path() -> Path | None:
    return _DB_PATH


def _connect() -> sqlite3.Connection:
    """Plain connection (caller manages commit/close).

    Timestamps are stored as ISO-8601 strings (not Python datetime), so we
    deliberately *don't* enable PARSE_DECLTYPES — otherwise sqlite3's bundled
    timestamp converter chokes on whatever ISO format we used.
    """
    if _DB_PATH is None:
        raise RuntimeError("DB not initialised — call init_db(path) first")
    c = sqlite3.connect(str(_DB_PATH), timeout=10.0)
    c.row_factory = sqlite3.Row
    c.execute("PRAGMA foreign_keys = ON")
    c.execute("PRAGMA journal_mode = WAL")
    return c


@contextmanager
def conn():
    """Auto-commit + close context manager."""
    c = _connect()
    try:
        yield c
        c.commit()
    except Exception:
        c.rollback()
        raise
    finally:
        c.close()


# ---------------------------------------------------------------------------
# Sync from external sources
# ---------------------------------------------------------------------------

def sync_rows(rows: list[dict]) -> None:
    """Mirror the in-memory rows list into the DB (full replace)."""
    with conn() as c:
        c.execute("DELETE FROM rows")
        c.execute("DELETE FROM rows_fts")
        for r in rows:
            p = r.get("parsed") or {}
            c.execute(
                """INSERT INTO rows
                   (row_num, col_a, col_b, col_c, col_d, col_e, col_f,
                    section, section_inherited, pdf_rel, pdf_inherited,
                    needs_col_d, parsed_type, parsed_brand, parsed_model,
                    parsed_ref, parsed_page, parsed_item, parsed_subitem)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (
                    r.get("row"),
                    _s(r.get("A")), _s(r.get("B")), _s(r.get("C")),
                    _s(r.get("D")), _s(r.get("E")), _s(r.get("F")),
                    _s(r.get("section_inferred")),
                    1 if r.get("section_inherited") else 0,
                    _s(r.get("pdf_rel")),
                    1 if r.get("pdf_inherited") else 0,
                    1 if r.get("needs_col_d") else 0,
                    _s(p.get("type")), _s(p.get("brand")), _s(p.get("model")),
                    _s(p.get("ref")), p.get("page"),
                    p.get("item"), p.get("subitem"),
                ),
            )
            c.execute(
                """INSERT INTO rows_fts (row_num, section, col_a, col_b, col_c, col_d, col_e)
                   VALUES (?, ?, ?, ?, ?, ?, ?)""",
                (
                    r.get("row"),
                    _s(r.get("section_inferred")),
                    _s(r.get("A")), _s(r.get("B")), _s(r.get("C")),
                    _s(r.get("D")), _s(r.get("E")),
                ),
            )


def sync_pdfs(pdf_records: list[dict]) -> None:
    """Mirror catalog PDFs (full replace).

    Each record is ``{rel_path, folder_key, section_prefix, size, mtime,
    num_pages, brand, model, annotations}``.  Annotations are written into
    pdf_annotations.
    """
    with conn() as c:
        c.execute("DELETE FROM pdf_annotations")
        c.execute("DELETE FROM pdfs")
        for rec in pdf_records:
            cur = c.execute(
                """INSERT INTO pdfs
                   (rel_path, folder_key, section_prefix, size, mtime,
                    num_pages, detected_brand, detected_model)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                (
                    rec["rel_path"], _s(rec.get("folder_key")),
                    _s(rec.get("section_prefix")), rec.get("size"),
                    rec.get("mtime"), rec.get("num_pages"),
                    _s(rec.get("brand")), _s(rec.get("model")),
                ),
            )
            pdf_id = cur.lastrowid
            for ann in rec.get("annotations") or []:
                rect = ann.get("rect") or [0, 0, 0, 0]
                c.execute(
                    """INSERT INTO pdf_annotations
                       (pdf_id, page_num, xref, annot_type, rect_x0, rect_y0,
                        rect_x1, rect_y1, contents, is_inline)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    (
                        pdf_id, ann.get("page", 0), ann.get("xref", 0),
                        _s(ann.get("type")),
                        rect[0], rect[1], rect[2], rect[3],
                        _s(ann.get("contents")),
                        1 if ann.get("_inline") else 0,
                    ),
                )


def sync_tor(section_index: dict, page_texts: dict) -> None:
    with conn() as c:
        c.execute("DELETE FROM tor_sections")
        for sec, page in section_index.items():
            c.execute("INSERT INTO tor_sections (section, page_num) VALUES (?, ?)",
                      (sec, page))
        c.execute("DELETE FROM tor_pages")
        for page, text in page_texts.items():
            c.execute("INSERT INTO tor_pages (page_num, normalized_text) VALUES (?, ?)",
                      (page, text))


def sync_snapshots(snaps: list[dict]) -> None:
    with conn() as c:
        c.execute("DELETE FROM snapshots")
        for s in snaps:
            c.execute(
                """INSERT INTO snapshots
                   (snapshot_id, tag, kind, timestamp, size, n_files, n_output, has_tarball)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                (
                    s.get("id"), _s(s.get("tag")), _s(s.get("kind")),
                    _s(s.get("timestamp")), s.get("size", 0),
                    s.get("n_tracked", 0), s.get("n_output", 0),
                    1 if s.get("has_tarball") else 0,
                ),
            )


def sync_pdf_history(records: list[dict]) -> None:
    with conn() as c:
        c.execute("DELETE FROM pdf_history")
        for r in records:
            c.execute(
                """INSERT OR IGNORE INTO pdf_history
                   (pdf_rel, snapshot_filename, ts, tag, size)
                   VALUES (?, ?, ?, ?, ?)""",
                (r["pdf_rel"], r["snapshot_filename"],
                 r.get("ts"), _s(r.get("tag")), r.get("size")),
            )


# ---------------------------------------------------------------------------
# Verification status (write-through API)
# ---------------------------------------------------------------------------

def import_status_from_json(path: Path) -> int:
    """One-time migration from verification_status.json into the DB."""
    if not path.exists():
        return 0
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return 0
    n = 0
    with conn() as c:
        for row, entry in (data or {}).items():
            try:
                rn = int(row)
            except (TypeError, ValueError):
                continue
            c.execute(
                """INSERT OR REPLACE INTO verification_status
                   (row_num, status, notes, updated_at, updated_by)
                   VALUES (?, ?, ?, ?, ?)""",
                (rn, _s(entry.get("status")), _s(entry.get("notes")),
                 _s(entry.get("updated_at")), _s(entry.get("updated_by"))),
            )
            n += 1
    return n


def get_all_status() -> dict:
    """Return ``{row_num_str: {status, notes, updated_at}}``."""
    out = {}
    with conn() as c:
        for r in c.execute("SELECT * FROM verification_status"):
            out[str(r["row_num"])] = {
                "status": r["status"],
                "notes": r["notes"] or "",
                "updated_at": r["updated_at"] or "",
                "updated_by": r["updated_by"] or "",
            }
    return out


def set_status(row_num: int, status: str | None = None,
               notes: str | None = None, actor: str = "user") -> dict:
    """Update verification status; logs an entry in audit_log."""
    with conn() as c:
        prev = c.execute(
            "SELECT status, notes FROM verification_status WHERE row_num = ?",
            (row_num,),
        ).fetchone()
        prev_status = prev["status"] if prev else None
        prev_notes = prev["notes"] if prev else None

        new_status = status if status is not None else prev_status
        new_notes = notes if notes is not None else prev_notes
        ts = datetime.now().isoformat(timespec="seconds")

        c.execute(
            """INSERT INTO verification_status (row_num, status, notes, updated_at, updated_by)
               VALUES (?, ?, ?, ?, ?)
               ON CONFLICT(row_num) DO UPDATE SET
                 status = excluded.status,
                 notes  = excluded.notes,
                 updated_at = excluded.updated_at,
                 updated_by = excluded.updated_by""",
            (row_num, new_status, new_notes, ts, actor),
        )

        if status is not None and status != prev_status:
            log_audit(c, action="status_change", target_type="row",
                      target_id=str(row_num),
                      before=prev_status or "unverified",
                      after=status or "unverified",
                      details={"notes": new_notes}, actor=actor)
        elif notes is not None and notes != prev_notes:
            log_audit(c, action="notes_update", target_type="row",
                      target_id=str(row_num), actor=actor,
                      details={"length": len(notes or "")})

    return {"row": row_num, "status": new_status, "notes": new_notes,
            "updated_at": ts}


# ---------------------------------------------------------------------------
# Auto-annotate plan tracking
# ---------------------------------------------------------------------------

def record_plan(row_num: int, plan: dict) -> int:
    with conn() as c:
        cur = c.execute(
            """INSERT INTO auto_annotate_plans
               (row_num, proposed_c, proposed_d, annotations_json, role, warnings_json)
               VALUES (?, ?, ?, ?, ?, ?)""",
            (
                row_num, _s(plan.get("proposed_c")), _s(plan.get("proposed_d")),
                json.dumps(plan.get("annotations") or [], ensure_ascii=False),
                _s((plan.get("role") or {}).get("role")),
                json.dumps(plan.get("warnings") or [], ensure_ascii=False),
            ),
        )
        return cur.lastrowid


def mark_plan_applied(plan_id: int, result: dict) -> None:
    with conn() as c:
        c.execute(
            """UPDATE auto_annotate_plans
               SET applied_at = ?, apply_result = ?
               WHERE plan_id = ?""",
            (datetime.now().isoformat(timespec="seconds"),
             json.dumps(result, ensure_ascii=False), plan_id),
        )


# ---------------------------------------------------------------------------
# Audit log
# ---------------------------------------------------------------------------

def log_audit(c: sqlite3.Connection | None = None, *,
              action: str,
              target_type: str | None = None,
              target_id: str | None = None,
              before: str | None = None,
              after: str | None = None,
              details: dict | None = None,
              actor: str = "system") -> None:
    """Insert an audit row.  Pass an existing connection ``c`` to share a
    transaction; otherwise a new connection is opened."""
    payload = (
        action, target_type, target_id,
        json.dumps(details, ensure_ascii=False) if details else None,
        before, after, actor,
    )
    sql = """INSERT INTO audit_log
             (action, target_type, target_id, details_json,
              before_value, after_value, actor)
             VALUES (?, ?, ?, ?, ?, ?, ?)"""
    if c is None:
        with conn() as cc:
            cc.execute(sql, payload)
    else:
        c.execute(sql, payload)


def recent_audit(limit: int = 100, since_ts: str | None = None,
                 action_filter: str | None = None) -> list[dict]:
    sql = "SELECT * FROM audit_log"
    params: list = []
    where = []
    if since_ts:
        where.append("ts >= ?"); params.append(since_ts)
    if action_filter:
        where.append("action = ?"); params.append(action_filter)
    if where:
        sql += " WHERE " + " AND ".join(where)
    sql += " ORDER BY ts DESC LIMIT ?"
    params.append(limit)
    out = []
    with conn() as c:
        for r in c.execute(sql, params):
            out.append({
                "log_id": r["log_id"],
                "ts": r["ts"],
                "action": r["action"],
                "target_type": r["target_type"],
                "target_id": r["target_id"],
                "details": json.loads(r["details_json"]) if r["details_json"] else None,
                "before": r["before_value"],
                "after": r["after_value"],
                "actor": r["actor"],
            })
    return out


# ---------------------------------------------------------------------------
# Search / stats
# ---------------------------------------------------------------------------

def fts_search(query: str, limit: int = 50) -> list[dict]:
    """Full-text search over Col A/B/C/D/E + section.  Returns matched
    row_num list with snippets."""
    if not query or not query.strip():
        return []
    safe_q = query.replace('"', '""')
    sql = """SELECT row_num, section,
                    snippet(rows_fts, 3, '<mark>', '</mark>', ' … ', 16) AS snippet_b,
                    snippet(rows_fts, 5, '<mark>', '</mark>', ' … ', 16) AS snippet_d,
                    bm25(rows_fts) AS rank
             FROM rows_fts
             WHERE rows_fts MATCH ?
             ORDER BY rank
             LIMIT ?"""
    out = []
    with conn() as c:
        try:
            for r in c.execute(sql, (safe_q, limit)):
                out.append({
                    "row": r["row_num"],
                    "section": r["section"] or "",
                    "snippet_b": r["snippet_b"] or "",
                    "snippet_d": r["snippet_d"] or "",
                    "rank": r["rank"],
                })
        except sqlite3.OperationalError as e:
            sys.stderr.write(f"[fts_search] query error: {e}\n")
    return out


def stats_summary() -> dict:
    """One-shot summary used by the dashboard."""
    with conn() as c:
        total_rows = c.execute("SELECT COUNT(*) c FROM rows").fetchone()["c"]
        with_pdf = c.execute(
            "SELECT COUNT(*) c FROM rows WHERE pdf_rel IS NOT NULL"
        ).fetchone()["c"]
        needs_d = c.execute(
            "SELECT COUNT(*) c FROM rows WHERE needs_col_d = 1"
        ).fetchone()["c"]
        status_counts = {}
        for r in c.execute(
            "SELECT status, COUNT(*) c FROM verification_status GROUP BY status"
        ):
            status_counts[r["status"]] = r["c"]
        type_counts = {}
        for r in c.execute(
            "SELECT parsed_type, COUNT(*) c FROM rows GROUP BY parsed_type"
        ):
            type_counts[r["parsed_type"] or "(empty)"] = r["c"]
        n_pdfs = c.execute("SELECT COUNT(*) c FROM pdfs").fetchone()["c"]
        n_annots = c.execute(
            "SELECT COUNT(*) c FROM pdf_annotations"
        ).fetchone()["c"]
        n_snap = c.execute("SELECT COUNT(*) c FROM snapshots").fetchone()["c"]
        n_audit = c.execute("SELECT COUNT(*) c FROM audit_log").fetchone()["c"]
        n_plans = c.execute(
            "SELECT COUNT(*) c FROM auto_annotate_plans WHERE applied_at IS NOT NULL"
        ).fetchone()["c"]
    return {
        "rows_total": total_rows,
        "rows_with_pdf": with_pdf,
        "rows_needs_col_d": needs_d,
        "status_counts": status_counts,
        "type_counts": type_counts,
        "pdfs": n_pdfs,
        "annotations": n_annots,
        "snapshots": n_snap,
        "audit_entries": n_audit,
        "auto_annotates_applied": n_plans,
    }


def section_progress() -> list[dict]:
    """Per-section verification progress."""
    sql = """
        SELECT r.section, COUNT(*) AS total,
               SUM(CASE WHEN s.status = 'pass'      THEN 1 ELSE 0 END) AS pass,
               SUM(CASE WHEN s.status = 'fail'      THEN 1 ELSE 0 END) AS fail,
               SUM(CASE WHEN s.status = 'need_fix'  THEN 1 ELSE 0 END) AS need_fix,
               SUM(CASE WHEN s.status = 'skip'      THEN 1 ELSE 0 END) AS skip
        FROM rows r
        LEFT JOIN verification_status s ON r.row_num = s.row_num
        WHERE r.section IS NOT NULL
        GROUP BY r.section
        ORDER BY r.section
    """
    out = []
    with conn() as c:
        for r in c.execute(sql):
            out.append({k: r[k] for k in r.keys()})
    return out


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _s(v) -> str | None:
    """Coerce to str (or None) for SQLite TEXT columns."""
    if v is None:
        return None
    if isinstance(v, str):
        return v
    return str(v)
