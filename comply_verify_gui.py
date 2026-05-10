#!/usr/bin/env python3
"""
Comply Verify GUI — visual verification tool for Smart Plant 1 Comply spec.

Reads:
  - output/Comply spec Smart Plant 1.xlsx
  - output/**/*.pdf (catalog PDFs with annotations)
  - TOR/*.pdf (optional)
  - BOQ/*.xlsx (optional)

Usage:
  python3 comply_verify_gui.py
  → open http://127.0.0.1:5173 in browser

Persistence:
  output/verification_status.json
"""

from __future__ import annotations

import hashlib
import io
import json
import os
import re
import shutil
import subprocess
import sys
import threading
import time
import urllib.parse
import webbrowser
from datetime import datetime
from pathlib import Path

import openpyxl
from flask import Flask, Response, abort, jsonify, render_template, request, send_file

try:
    import fitz  # PyMuPDF
except ImportError:
    sys.stderr.write("PyMuPDF (fitz) is required. Install: pip3 install --user pymupdf\n")
    sys.exit(1)

try:
    from PIL import Image, ImageDraw
except ImportError:
    sys.stderr.write("Pillow is required. Install: pip3 install --user pillow\n")
    sys.exit(1)

from app import catalog, core, learning
from app import database as db

ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "output"


def _load_dotenv(path: Path) -> None:
    """Minimal .env loader so we don't pull in python-dotenv as a dep.
    Lines like KEY=value (no quotes needed). Skips comments + blanks.
    Existing env vars take precedence (so shell exports win)."""
    if not path.exists():
        return
    try:
        for line in path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            k = k.strip()
            v = v.strip().strip('"').strip("'")
            os.environ.setdefault(k, v)
    except Exception as e:
        sys.stderr.write(f"[_load_dotenv] {e}\n")

# Layout detection — supports both:
#   1. comply-module/ layout (current): TOR, BOQ, scripts, _versions sit
#      alongside comply_verify_gui.py (PROJECT == ROOT)
#   2. Legacy Tools/ layout: TOR, BOQ live in the parent (PROJECT ==
#      ROOT.parent)
PROJECT = ROOT if (ROOT / "TOR").exists() else ROOT.parent
TOR_DIR = PROJECT / "TOR"
BOQ_DIR = PROJECT / "BOQ"

# Project-level snapshot system (scripts/version.py + _versions/snapshots/)
SCRIPTS_DIR    = ROOT / "scripts"
VERSIONS_DIR   = ROOT / "_versions"
SNAPS_DIR      = VERSIONS_DIR / "snapshots"
VERSION_SCRIPT = SCRIPTS_DIR / "version.py"

XLSX_PATH = OUTPUT / "Comply spec Smart Plant 1.xlsx"
STATUS_PATH = OUTPUT / "verification_status.json"

# SQLite database — local cache + audit log + FTS search index. Lives in
# the project; survives across boots; safe to delete (rebuilt from xlsx +
# filesystem on next boot).
DB_DIR  = ROOT / "_db"
DB_PATH = DB_DIR / "comply.db"

# Project singleton — OOP entry point that owns everything else. The legacy
# globals (ROWS, PDF_INDEX, etc.) coexist for now; new code should prefer
# this object. ``project.rows`` mirrors ROWS automatically via load_rows().
PROJECT_OBJ = core.Project(ROOT)

_TEMPLATES_DIR = ROOT / "app" / "server" / "templates"
_STATIC_DIR = ROOT / "app" / "server" / "static"
app = Flask(
    __name__,
    template_folder=str(_TEMPLATES_DIR),
    static_folder=str(_STATIC_DIR) if _STATIC_DIR.exists() else None,
    static_url_path="/static" if _STATIC_DIR.exists() else None,
)
app.config["JSON_AS_ASCII"] = False

# T2.2 (2026-05-10): split-blueprint registration. Catalog / export /
# continuity routes live in app/routes/*; the gui module keeps only
# the routes that touch xlsx mutation state (apply_catalog, status,
# auto_annotate, manual_annotate, etc.).
from app.routes import register_all as _register_route_blueprints  # noqa: E402

_register_route_blueprints(app, root=ROOT, output_root=OUTPUT)


# ---------------------------------------------------------------------------
# Indexing
# ---------------------------------------------------------------------------

# ref-key (e.g. "5.1.2-2") → list[Path]
PDF_INDEX: dict[str, list[Path]] = {}
# section-key (e.g. "5.1.2") → list[Path]   (section-level matches, no -N)
SECTION_INDEX: dict[str, list[Path]] = {}
# model-substring → list[Path]   (for Col D that is just a model name)
MODEL_INDEX: dict[str, list[Path]] = {}
# rows from xlsx
ROWS: list[dict] = []
# header info (TOR PDF, BOQ xlsx)
EXTRA_REFS: dict = {}


def build_pdf_index() -> None:
    """Index PDFs by ref-key (5.1.2-1), section (5.1.6.2), and model substrings.

    Folder convention:
      - Parent rack folder: '{section}.-{N}'   (e.g. '5.1.2.-1', '5.2.1.-9')
      - Sub-item PDFs sit beside parent in same folder: '{section}.M-{n} ...pdf'
      - Sensor/LED/เสา folder: '{section}. {thai-name}' (no -N suffix)
    """
    PDF_INDEX.clear()
    SECTION_INDEX.clear()
    MODEL_INDEX.clear()
    if not OUTPUT.exists():
        return
    for pdf in OUTPUT.rglob("*.pdf"):
        if pdf.name.startswith("~"):
            continue
        if pdf.name == "Comply spec Smart Plant 1.pdf":
            continue
        stem = pdf.stem
        parent = pdf.parent.name

        # Folder pattern A: "{section}.-{N}" — parent rack folder
        # (allow optional whitespace before the dash, e.g. "5.1.3.1. -1")
        fm = re.match(r"^(\d+(?:\.\d+){1,3})\.?\s*-(\d+)\b", parent)
        if fm:
            sec, n = fm.group(1), fm.group(2)
            key_parent = f"{sec}-{n}"
            # Is this the parent file? filename ends with "-N"
            tail = re.search(r"-(\d+)$", stem)
            if tail and tail.group(1) == n:
                PDF_INDEX.setdefault(key_parent, []).insert(0, pdf)
            else:
                # Sub-item PDF — filename starts with sub-section then -n
                im = re.match(r"^(\d+(?:\.\d+){2,3})-(\d+)\b", stem)
                if im:
                    PDF_INDEX.setdefault(f"{im.group(1)}-{im.group(2)}", []).append(pdf)
                else:
                    PDF_INDEX.setdefault(key_parent, []).append(pdf)
        else:
            # Folder pattern B: "{section}. {thai-name}" — section-only ref (sensor/LED/เสา)
            sm = re.match(r"^(\d+(?:\.\d+){1,3})", parent)
            if sm:
                SECTION_INDEX.setdefault(sm.group(1), []).append(pdf)
            else:
                # Fallback: index by stem section
                sm2 = re.match(r"^(\d+(?:\.\d+){1,3})", stem)
                if sm2:
                    SECTION_INDEX.setdefault(sm2.group(1), []).append(pdf)

        # Model index — for Col D that names just a model (e.g. "UFC9312A")
        text = f"{parent} {stem}"
        for tok in re.findall(r"[A-Za-z][\w\-]{2,}", text):
            if len(tok) >= 3 and not tok.isdigit():
                MODEL_INDEX.setdefault(tok.lower(), []).append(pdf)


def parse_col_d(d: str | None) -> dict:
    """Parse Col D — return dict with type + extracted fields."""
    out: dict = {"raw": d or "", "type": "empty"}
    if not d:
        return out
    s = d.strip()
    out["raw"] = s
    # Type classification
    if s.startswith("ยี่ห้อ"):
        out["type"] = "brand_model"
        m = re.match(r"ยี่ห้อ\s+(.+?)\s+รุ่น\s+(.+)$", s)
        if m:
            out["brand"] = m.group(1).strip()
            out["model"] = m.group(2).strip()
    elif s.startswith("เทียบเท่าข้อกำหนด"):
        out["type"] = "equivalent"
    elif s.startswith("สูงกว่าข้อกำหนด"):
        out["type"] = "higher"
    elif s.startswith("ยินดีปฏิบัติ"):
        out["type"] = "commitment"
    elif re.match(r"^\d+(?:\.\d+){2,3}", s):
        out["type"] = "filename_format"
    else:
        # fall back — could be model-only string (e.g., "UFC9312A")
        out["type"] = "model_only"

    # Extract reference fields if present
    m = re.search(r"เอกสาร\s+(\d+(?:\.\d+){1,3}(?:-\d+)?)", s)
    if m:
        out["ref"] = m.group(1)
    m = re.search(r"หน้า\s+(\d+)", s)
    if m:
        out["page"] = int(m.group(1))
    m = re.search(r"ข้อ\s+(\d+)\)\s*ข้อย่อย\s+(\d+)", s)
    if m:
        out["item"] = int(m.group(1))
        out["subitem"] = int(m.group(2))
    else:
        m = re.search(r"ข้อ\s+(\d+)\)", s)
        if m:
            out["item"] = int(m.group(1))
        m = re.search(r"ข้อย่อย\s+(\d+)", s)
        if m:
            out["subitem"] = int(m.group(1))

    return out


def _section_to_folder_key(section: str) -> str | None:
    """Map a section number to the PDF_INDEX key for its folder.

    "5.2.1.9" → "5.2.1-9"  (single-row item under 5.2.1.-9/ folder)
    "5.1.1.2" → "5.1.1-2"  (parent rack folder 5.1.1.-2/)
    """
    parts = section.split(".")
    if len(parts) < 3:
        return None
    return f"{'.'.join(parts[:-1])}-{parts[-1]}"


def _pick_best_in_folder(paths: list[Path], section: str, raw: str) -> Path | None:
    """When PDF_INDEX[key] has multiple PDFs, prefer the one whose filename
    matches Col D's tokens most distinctively (avoids picking sub-item PDFs
    when the row points to the parent, or vice versa)."""
    if not paths:
        return None
    if len(paths) == 1:
        return paths[0]
    # Distinctive tokens from Col D (English/numeric, ≥4 chars or has digit)
    raw_tokens = []
    for tok in re.findall(r"[A-Za-z][\w\-]{2,}", raw or ""):
        if len(tok) >= 4 or any(c.isdigit() for c in tok):
            raw_tokens.append(tok.lower())
    # Score each PDF by token overlap with its filename
    scored = []
    for p in paths:
        stem_l = p.stem.lower()
        score = sum(1 for t in raw_tokens if t in stem_l)
        # Bonus: filename starts with the exact section prefix
        if p.stem.startswith(section + " ") or p.stem.startswith(section + "."):
            score += 5
        # Bonus: filename ends with `-N` matching last section digit (parent file)
        if section.split(".")[-1] in p.stem.rsplit("-", 1)[-1]:
            score += 1
        scored.append((score, p))
    scored.sort(key=lambda x: (-x[0], str(x[1])))
    return scored[0][1]


def _ref_to_folder_keys(ref: str) -> list[str]:
    """Yield candidate PDF_INDEX/SECTION_INDEX keys for a Col D ref.

    Users mix two conventions:
      • dash form  "5.1.1-2"   → PDF_INDEX["5.1.1-2"]   (parent rack folder)
      • dot form   "5.1.1.2"   → folder "5.1.1.-2/"     → "5.1.1-2"
      • dot form   "5.1.6.2"   → SECTION_INDEX["5.1.6.2"]  (sensor/LED)
    The clone/auto-annotate scripts emit dot form; the legacy convention is
    dash form. Try both.
    """
    if not ref:
        return []
    out = [ref]
    # If ref is dot form depth ≥3, also try last-segment-as-N
    if "-" not in ref:
        parts = ref.split(".")
        if len(parts) >= 3:
            try:
                int(parts[-1])
                folder = ".".join(parts[:-1]) + "-" + parts[-1]
                out.append(folder)
            except ValueError:
                pass
    # If ref is dash form, also try expanded dot form
    if "-" in ref:
        m = re.match(r"^(\d+(?:\.\d+)*)-(\d+)$", ref)
        if m:
            out.append(m.group(1) + "." + m.group(2))
    return out


def resolve_ref_to_pdf(parsed: dict, raw: str) -> Path | None:
    """Find PDF path from parsed Col D."""
    ref = parsed.get("ref")
    if ref:
        for k in _ref_to_folder_keys(ref):
            if k in PDF_INDEX:
                return _pick_best_in_folder(PDF_INDEX[k], k.replace("-", "."), raw)
            if k in SECTION_INDEX:
                return SECTION_INDEX[k][0]

    if parsed.get("type") == "filename_format":
        # (a) Direct {section}-{N} sub-item pattern at start of D
        m = re.match(r"^(\d+(?:\.\d+){2,3}-\d+)", raw)
        if m and m.group(1) in PDF_INDEX:
            return _pick_best_in_folder(PDF_INDEX[m.group(1)],
                                        m.group(1).replace("-", "."), raw)
        # (b) Single-row format "5.2.1.9 ..." → check folder 5.2.1.-9
        msec = re.match(r"^(\d+(?:\.\d+){2,3})\b", raw)
        if msec:
            sec = msec.group(1)
            # Section-only folder (sensor, LED, เสา) takes priority
            if sec in SECTION_INDEX:
                return SECTION_INDEX[sec][0]
            # Otherwise look in `parent.-N/` folder
            key = _section_to_folder_key(sec)
            if key and key in PDF_INDEX:
                return _pick_best_in_folder(PDF_INDEX[key], sec, raw)
        # (c) Older shorter section "5.1.3.2." → SECTION_INDEX
        m2 = re.match(r"^(\d+(?:\.\d+){1,3})", raw)
        if m2 and m2.group(1) in SECTION_INDEX:
            return SECTION_INDEX[m2.group(1)][0]

    # model-only: scan tokens, but ONLY accept a hit that's also unique
    # (multiple PDFs share many tokens — picking the first match silently
    # routes rows to the wrong catalog).
    if parsed.get("type") in ("model_only", "filename_format"):
        candidate_pdfs: list[Path] = []
        for tok in re.findall(r"[A-Za-z][\w\-]{2,}", raw):
            paths = MODEL_INDEX.get(tok.lower())
            if paths:
                # Only trust unique matches OR matches that contain the
                # exact token as a "model" suffix (e.g. UF-2010A)
                if len(set(paths)) == 1:
                    return paths[0]
                candidate_pdfs.extend(paths)
        # Score remaining candidates (same heuristic as folder match)
        if candidate_pdfs:
            return _pick_best_in_folder(list(set(candidate_pdfs)), "", raw)
    return None


def load_rows() -> None:
    ROWS.clear()
    if not XLSX_PATH.exists():
        sys.stderr.write(f"WARN: {XLSX_PATH} not found\n")
        return
    wb = openpyxl.load_workbook(XLSX_PATH, data_only=True)
    ws = wb.active
    # rows 1-4 are header — start at row 5
    for r in range(5, ws.max_row + 1):
        a = ws.cell(r, 1).value
        b = ws.cell(r, 2).value
        c = ws.cell(r, 3).value
        d = ws.cell(r, 4).value
        e = ws.cell(r, 5).value
        f = ws.cell(r, 6).value
        if not any([a, b, c, d, e, f]):
            continue
        parsed = parse_col_d(d)
        pdf = resolve_ref_to_pdf(parsed, str(d) if d else "")
        # Detect section from B (e.g., "5.1.2." or "          1) ..." etc.)
        section = None
        if a:
            section = str(a).strip()
        elif b:
            mb = re.match(r"\s*(\d+(?:\.\d+)+)\.?", str(b))
            if mb:
                section = mb.group(1)
        ROWS.append({
            "row": r,
            "A": a,
            "B": b,
            "C": c,
            "D": d,
            "E": e,
            "F": f,                                       # existing status
            "parsed": parsed,
            "pdf_rel": str(pdf.relative_to(OUTPUT)) if pdf else None,
            "pdf_inherited": False,
            "section_inferred": section,
        })

    # Second pass: brand_model parent rows inherit PDF from next sub-item row
    # (so clicking "ยี่ห้อ Schneider รุ่น QO..." shows the Schneider/QO catalog page)
    for i, r in enumerate(ROWS):
        if r["parsed"].get("type") == "brand_model" and not r["pdf_rel"]:
            for j in range(i + 1, min(i + 6, len(ROWS))):
                nxt = ROWS[j]
                if nxt["pdf_rel"]:
                    r["pdf_rel"] = nxt["pdf_rel"]
                    r["pdf_inherited"] = True
                    r["parsed"].setdefault("page", 1)
                    break
            # Fallback: use the row's own section to find the parent rack
            # folder (handles brand_model headers when sub-items haven't been
            # auto-annotated yet)
            if not r["pdf_rel"]:
                sec = r.get("section_inferred")
                if sec:
                    for k in _ref_to_folder_keys(sec):
                        if k in PDF_INDEX:
                            r["pdf_rel"] = str(PDF_INDEX[k][0].relative_to(OUTPUT))
                            r["pdf_inherited"] = True
                            r["parsed"].setdefault("page", 1)
                            break
                        if k in SECTION_INDEX:
                            r["pdf_rel"] = str(SECTION_INDEX[k][0].relative_to(OUTPUT))
                            r["pdf_inherited"] = True
                            r["parsed"].setdefault("page", 1)
                            break

    # Third pass: propagate inferred section to sub-item rows without one.
    # When inheriting, prefer the MORE SPECIFIC of:
    #   • last_section (from preceding section-header row)
    #   • D-ref-derived section (e.g. ref "5.1.2-2" → "5.1.2")
    # Otherwise a row in section 5.1.1.2 with D ref "5.1.1-2" would be
    # mis-inferred as "5.1.1" (less specific) instead of "5.1.1.2".
    def _section_depth(s: str) -> int:
        return len(s.split("."))

    last_section: str | None = None
    for r in ROWS:
        if r["section_inferred"]:
            last_section = r["section_inferred"]
            continue

        candidates: list[str] = []
        if last_section:
            candidates.append(last_section)
        d_ref = r["parsed"].get("ref")
        if d_ref:
            m = re.match(r"^(\d+(?:\.\d+){1,3})", d_ref)
            if m:
                candidates.append(m.group(1))

        if candidates:
            # Pick deepest section number; ties broken by D-ref preference
            chosen = max(candidates, key=lambda s: (_section_depth(s), s != last_section))
            r["section_inferred"] = chosen
            r["section_inherited"] = True

    # Fourth pass: rows that don't yet have a pdf_rel but DO have a section
    # → try to find a candidate catalog by folder convention.
    #
    # Also runs for "ยินดีปฏิบัติตามข้อกำหนด" (commitment) rows so the user
    # can SEE the catalog when invoking 📍 Mark in the manual-annotate
    # workflow. Col D itself isn't changed — only the surface pdf_rel mapping.
    #
    # Uses _candidate_pdfs_for_section which tries multiple lookup strategies:
    #   • exact section ("5.1.6.2" → SECTION_INDEX)
    #   • folder convention ("5.1.1.2" → PDF_INDEX["5.1.1-2"])
    #   • wildcard for sections with multiple sub-catalogs
    #     ("5.1.2"  → first of 5.1.2-1, 5.1.2-2, …)
    #     ("5.1.6.1" → first of 5.1.6.1-1, …)
    #   • parent fallback ("5.1.6.1" → 5.1.6 sensors/เสา)
    for r in ROWS:
        if r["pdf_rel"]:
            continue
        d_value = (r.get("D") or "").strip()
        is_commit = d_value.startswith("ยินดีปฏิบัติ")
        if d_value and not is_commit:
            continue
        sec = r.get("section_inferred")
        if not sec:
            continue
        candidates = _candidate_pdfs_for_section(sec)
        if not candidates:
            continue
        r["pdf_rel"] = str(candidates[0].relative_to(OUTPUT))
        r["pdf_inherited"] = True
        if not d_value:
            r["needs_col_d"] = True

    # Mirror into the OOP Project singleton so new code paths can use it.
    PROJECT_OBJ.set_rows([core.from_legacy_dict(r) for r in ROWS])


def collect_extra_refs() -> None:
    EXTRA_REFS.clear()
    EXTRA_REFS["tor_pdfs"] = sorted(str(p.relative_to(PROJECT)) for p in TOR_DIR.glob("*.pdf")) if TOR_DIR.exists() else []
    EXTRA_REFS["boq_xlsx"] = sorted(str(p.relative_to(PROJECT)) for p in BOQ_DIR.glob("*.xlsx")) if BOQ_DIR.exists() else []
    EXTRA_REFS["comply_pdf"] = "Comply spec Smart Plant 1.pdf" if (OUTPUT / "Comply spec Smart Plant 1.pdf").exists() else None


# ---------------------------------------------------------------------------
# Tree builder — section hierarchy + sub-items grouped under parent
# ---------------------------------------------------------------------------

def _section_parts(sec: str) -> list[str]:
    return [p for p in sec.split(".") if p]


def build_tree() -> dict:
    """Return tree of {key, label, children, rows[], type}.

    Hierarchy:
      Smart Plant 1
      └─ 5
         └─ 5.1
            └─ 5.1.2
               ├─ R62 (parent header)
               ├─ R63 (item 1))
               ├─ R64 (item 2))
               │  ├─ R65 (sub-item 1.)
               │  ├─ R66 (sub-item 2.)
               │  ...
    """
    root = {"key": "ROOT", "label": "Smart Plant 1", "children": {}, "rows": [], "type": "root"}

    def get_or_create(node, key, label, ntype="section"):
        if key not in node["children"]:
            node["children"][key] = {"key": key, "label": label,
                                     "children": {}, "rows": [], "type": ntype}
        return node["children"][key]

    # First pass: section nodes
    last_section_node = root
    last_item_node = None  # within section, current ข้อ N) parent
    for r in ROWS:
        sec = r.get("section_inferred")
        b = (r.get("B") or "").strip()
        d = r.get("D") or ""
        if sec:
            parts = _section_parts(sec)
            node = root
            for i, _ in enumerate(parts):
                key = ".".join(parts[: i + 1])
                node = get_or_create(node, key, key)
            # this row IS the section header (or first row of section)
            node["rows"].append(r["row"])
            last_section_node = node
            last_item_node = None
        else:
            # sub-row — could be "1) ..." (parent item) or "1. ..." (sub-item)
            m_parent = re.match(r"^\s*(\d+)\s*\)", b)
            m_sub = re.match(r"^\s*(\d+)\s*\.", b)
            if m_parent:
                # ข้อ N) — child of last_section_node
                key = f"{last_section_node['key']}#{m_parent.group(1)})"
                item = get_or_create(last_section_node, key,
                                     f"ข้อ {m_parent.group(1)})", ntype="item")
                item["rows"].append(r["row"])
                last_item_node = item
            elif m_sub and last_item_node is not None:
                # ข้อย่อย N. — child of last_item_node
                key = f"{last_item_node['key']}#{m_sub.group(1)}."
                sub = get_or_create(last_item_node, key,
                                    f"ข้อย่อย {m_sub.group(1)}.", ntype="sub")
                sub["rows"].append(r["row"])
            else:
                # fall-through — attach to nearest known parent
                target = last_item_node or last_section_node
                target["rows"].append(r["row"])

    return root


# ---------------------------------------------------------------------------
# TOR text search (find page by Col B content)
# ---------------------------------------------------------------------------

TOR_PATH: Path | None = None
TOR_CACHE: dict[int, dict] = {}              # row_num → {page, rects}
TOR_SECTION_INDEX: dict[str, int] = {}       # section → first page (1-indexed)
TOR_PAGE_TEXTS: dict[int, str] = {}          # page → normalized full text (for substring search)


def normalize_thai_text(s: str) -> str:
    """Aggressive normalization for cross-document Thai matching.

    Handles:
      - SARA AM (precomposed ำ U+0E33) ↔ NIKHAHIT + SARA AA (ํา) — TOR
        often uses the decomposed form while xlsx uses precomposed.
      - Combining-mark order: TOR sometimes writes NIKHAHIT before tone
        marks (`นํ้า`), xlsx writes tone before SARA AM (`น้ำ`). Unicode
        canonical ordering doesn't fix this because NIKHAHIT has
        combining class 0. We swap explicitly.
      - Whitespace collapse + case fold.
    """
    if not s:
        return ""
    # SARA AM decomposition equivalence
    s = s.replace("ำ", "ํา")  # always force decomposed form
    # NIKHAHIT (U+0E4D) before tone marks (0E48..0E4B) → tone before NIKHAHIT
    # (this normalizes to the order typically produced by xlsx)
    s = re.sub(r"ํ([่-๋])", r"\1ํ", s)
    # collapse all whitespace (incl. zero-width and NBSP) to single space
    s = re.sub(r"[\s ​]+", " ", s)
    return s.strip().lower()


def index_tor_text() -> None:
    """Pre-extract + normalize every TOR page so we can do substring search
    independent of Unicode form."""
    global TOR_PATH, TOR_PAGE_TEXTS
    TOR_PAGE_TEXTS = {}
    if TOR_PATH is None:
        TOR_PATH = find_tor_pdf()
    if TOR_PATH is None:
        return
    try:
        doc = fitz.open(TOR_PATH)
    except Exception:
        return
    for pno in range(len(doc)):
        try:
            txt = doc[pno].get_text("text")
        except Exception:
            txt = ""
        TOR_PAGE_TEXTS[pno + 1] = normalize_thai_text(txt)


def find_tor_pdf() -> Path | None:
    if TOR_DIR.exists():
        pdfs = sorted(TOR_DIR.glob("*.pdf"))
        if pdfs:
            return pdfs[0]
    return None


# Match a section heading at line start. Anchor with a dot or whitespace
# AFTER the section to avoid catching arbitrary numbers like "85" or "5mm".
# Acceptable: "5.", "5.1", "5.1.2.", "5.1.2 ตู้..." (digit + dot/space + non-digit).
_SECTION_HEAD = re.compile(
    r"(?m)^\s*(\d+(?:\.\d+){0,4})\s*[\.\)]?\s+(?=[^\d])"
)


def _is_valid_section_for_index(sec: str) -> bool:
    """Reject "85" "100.00" etc. — keep only realistic section paths."""
    parts = sec.split(".")
    if len(parts) > 5:
        return False
    for p in parts:
        if not p.isdigit():
            return False
        n = int(p)
        if n < 1 or n > 50:
            return False
    return True


def build_tor_section_index() -> None:
    """Scan TOR once, map each section heading to its first occurrence page."""
    global TOR_PATH, TOR_SECTION_INDEX
    TOR_SECTION_INDEX = {}
    if TOR_PATH is None:
        TOR_PATH = find_tor_pdf()
    if TOR_PATH is None:
        return
    try:
        doc = fitz.open(TOR_PATH)
    except Exception:
        return
    for pno in range(len(doc)):
        try:
            text = doc[pno].get_text("text")
        except Exception:
            continue
        for m in _SECTION_HEAD.finditer(text):
            sec = m.group(1)
            if not _is_valid_section_for_index(sec):
                continue
            # Only record FIRST page of each section
            if sec not in TOR_SECTION_INDEX:
                TOR_SECTION_INDEX[sec] = pno + 1


def get_tor_section_range(section: str | None) -> tuple[int, int] | None:
    """Return (start, end) inclusive page range of TOR for the given section.

    Falls back to parent section if the exact section isn't indexed.
    The end page is determined by the next section at the same depth or
    shallower; if nothing follows, default to start + 6.
    """
    if not section or not TOR_SECTION_INDEX or TOR_PATH is None:
        return None
    parts = section.split(".")
    start = TOR_SECTION_INDEX.get(section)
    used_section = section
    if start is None:
        for i in range(len(parts) - 1, 0, -1):
            parent = ".".join(parts[:i])
            if parent in TOR_SECTION_INDEX:
                start = TOR_SECTION_INDEX[parent]
                used_section = parent
                break
    if start is None:
        return None

    # End: smallest page > start belonging to a section at depth <= used_section
    target_depth = len(used_section.split("."))
    end = None
    for sec, pg in TOR_SECTION_INDEX.items():
        if sec == used_section:
            continue
        if pg <= start:
            continue
        sec_depth = len(sec.split("."))
        # Sibling or shallower → defines the boundary
        if sec_depth <= target_depth:
            if end is None or pg < end:
                end = pg - 1
    if end is None:
        try:
            doc = fitz.open(TOR_PATH)
            end = min(start + 6, len(doc))
        except Exception:
            end = start + 6
    return (start, end)


def _normalize_for_search(s: str) -> str:
    """Strip leading whitespace + numbering prefix; pick a stable substring."""
    s = (s or "").strip()
    # remove leading "N)" or "N." or "5.1.2."
    s = re.sub(r"^\s*\d+(?:\.\d+)*\.?\s*", "", s)
    s = re.sub(r"^\s*\d+\s*[\)\.]\s*", "", s)
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def _thai_variants(s: str) -> list[str]:
    """TOR PDFs sometimes use decomposed Thai (NIKHAHIT + SARA AA), while xlsx
    uses precomposed SARA AM (U+0E33). NFD doesn't break U+0E33 (it's only a
    compatibility decomposition) — must use NFKD."""
    import unicodedata
    out = [s]
    for form in ("NFC", "NFD", "NFKC", "NFKD"):
        v = unicodedata.normalize(form, s)
        if v not in out:
            out.append(v)
    # also: explicit U+0E33 → U+0E4D + U+0E32 (TOR's preferred form)
    explicit = s.replace("ำ", "ํา")
    if explicit not in out:
        out.append(explicit)
    return out


def _try_get_rects(page, needle: str) -> list:
    """Best-effort rect lookup for `needle` on a page. Try the original needle
    plus several Thai variants. Return the first non-empty result."""
    for variant in _thai_variants(needle):
        try:
            rects = page.search_for(variant, quads=False)
        except Exception:
            rects = []
        if rects:
            return rects
    return []


def _extract_search_tokens(text: str) -> list[str]:
    """Distinctive tokens used as a fallback when the full phrase doesn't
    match. Prefer English words, model codes, numeric values with units."""
    s = re.sub(r"^\s*\d+\s*[\)\.]\s*", "", text or "")
    tokens: list[str] = []
    seen = set()

    def add(t):
        t = t.strip()
        if len(t) < 3:
            return
        k = t.lower()
        if k in seen:
            return
        seen.add(k)
        tokens.append(t)

    # English words / model codes (≥3 chars, must contain a letter)
    for m in re.finditer(r"[A-Za-z][A-Za-z0-9]{2,}(?:[-+/.][A-Za-z0-9]+)*", s):
        tok = m.group(0)
        if len(tok) >= 4 or any(c.isdigit() for c in tok):
            add(tok)
    # Standards
    for m in re.finditer(r"(?:IEC|ISO|ASTM|ANSI|UL|TIS|มอก\.?)\s*\d+(?:[-/.\s]\d+)*", s):
        add(m.group(0))
    # IP rating / DC voltage / typical units
    for m in re.finditer(r"\b(?:IP\d+|DC\s*\d+(?:-\d+)?\s*V|AC\s*\d+\s*V)\b", s):
        add(m.group(0))
    # Numeric + unit
    for m in re.finditer(r"\d+(?:[\.,]\d+)?\s*(?:mA|MHz|GHz|kHz|Hz|°C|°F|mm|cm|m|kg|hr|min|kV|V|A|W|kW)", s, re.IGNORECASE):
        add(m.group(0))
    return tokens


def _next_sibling_section(section: str) -> str | None:
    """Compute the next sibling section number at the same depth.
    "5.1.1.2" → "5.1.1.3". Used as a y-bound when filtering matches."""
    if not section:
        return None
    parts = section.split(".")
    try:
        last = int(parts[-1])
    except ValueError:
        return None
    return ".".join(parts[:-1] + [str(last + 1)])


def _section_y_bounds_on_page(page, section: str | None,
                              next_section: str | None) -> tuple[float, float] | None:
    """Find the y-range of `section`'s content on this page.

    Returns (y_top, y_bottom). y_top = y0 of section header marker (if found
    on this page; else 0). y_bottom = y0 of next-section marker (if any on
    this page; else page height).

    A rect entirely below y_bottom (or above y_top) belongs to a different
    section's content — the matcher should reject it.
    """
    if not section:
        return None
    page_h = page.rect.height
    y_top = 0.0
    y_bottom = page_h

    # Section header on this page? (e.g. "5.1.1.2." or "5.1.1.2 ")
    sec_hits: list = []
    for needle in (f"{section}.", f"{section} "):
        try:
            sec_hits.extend(page.search_for(needle, quads=False))
        except Exception:
            pass
    if sec_hits:
        # Section may have multiple hits (header + citations) — take topmost
        y_top = min(h.y0 for h in sec_hits)

    # Next sibling header on this page = upper bound for section's content
    if next_section:
        nx_hits: list = []
        for needle in (f"{next_section}.", f"{next_section} "):
            try:
                nx_hits.extend(page.search_for(needle, quads=False))
            except Exception:
                pass
        # Only consider next-section hits that come AFTER section header
        nx_hits = [h for h in nx_hits if h.y0 > y_top + 1]
        if nx_hits:
            y_bottom = min(h.y0 for h in nx_hits)

    return (y_top, y_bottom)


def _filter_rects_by_bounds(rects: list, bounds: tuple[float, float] | None,
                            margin: float = 4.0) -> list:
    if not bounds or not rects:
        return rects
    y_top, y_bottom = bounds
    return [r for r in rects
            if (r.y0 >= y_top - margin and r.y1 <= y_bottom + margin)]


def _densest_y_cluster(rects: list, max_span: float = 35.0) -> list:
    """Among scattered rects, return the tightest y-range cluster.

    When the full phrase can't be located by ``search_for`` (e.g. Thai
    combining-mark mismatches), the token fallback collects rects from many
    individual matches across the page. Most of those tokens repeat in
    headers/footers/citations; only the cluster on the actual spec line is
    relevant. This function picks the tightest band containing the most
    rects so we highlight a single line, not every occurrence on the page.
    """
    if len(rects) <= 1:
        return rects
    def y0(r):
        return r.y0 if hasattr(r, "y0") else r[1]
    sorted_rects = sorted(rects, key=y0)
    best_start, best_count = 0, 0
    for i, r in enumerate(sorted_rects):
        top = y0(r)
        count = sum(1 for r2 in sorted_rects if top <= y0(r2) <= top + max_span)
        if count > best_count:
            best_count = count
            best_start = i
    top = y0(sorted_rects[best_start])
    return [r for r in sorted_rects if top <= y0(r) <= top + max_span]


def _anchor_cluster_by_rarest(token_hits: dict, line_span: float = 14.0) -> list:
    """Pick the line containing the rarest token, then keep only token rects
    on that same line (within ``line_span`` of the anchor's y).

    ``token_hits`` is ``{token: [rect, ...]}``. Generic words like "Firewall"
    repeat across the section header + multiple sub-items, so densest-cluster
    can still span 3 lines. Anchoring on the RAREST token (the one that
    appears once on the whole page) pinpoints the actual content line.
    """
    if not token_hits:
        return []
    # Sort tokens by hit count ascending (rarest first); break ties by token
    # length (longer = more distinctive).
    ranked = sorted(token_hits.items(),
                    key=lambda kv: (len(kv[1]), -len(kv[0])))
    # If even the rarest appears many times, fall back to densest cluster
    rarest_tok, rarest_rects = ranked[0]
    if len(rarest_rects) > 4:
        flat = [r for rs in token_hits.values() for r in rs]
        return _densest_y_cluster(flat, max_span=line_span * 2)

    def y0(r):
        return r.y0 if hasattr(r, "y0") else r[1]

    # Try each occurrence of the rarest token as a potential anchor and
    # pick the anchor that gathers the most other-token rects on its line.
    best: tuple[list, int] = ([], -1)
    for anchor in rarest_rects:
        anchor_y = y0(anchor)
        line_rects = []
        for tok, rs in token_hits.items():
            for r in rs:
                if abs(y0(r) - anchor_y) <= line_span:
                    line_rects.append(r)
        if len(line_rects) > best[1]:
            best = (line_rects, len(line_rects))
    return best[0]


def _search_tor_in_pages(text: str, page_nums: list[int],
                          section: str | None = None) -> dict | None:
    """Two-stage search across the given page list:

    1. **Normalized substring** match against the pre-indexed TOR text.
       Robust to Thai Unicode form differences (handles SARA AM ↔ NIKHAHIT
       and combining-mark order).
    2. Once we know which page contains the match, use ``page.search_for``
       with multiple variants to recover rect coordinates for highlight.
    3. **Filter rects to within the section's y-bounds** on that page so
       phrases that repeat across sections don't bleed into the highlight.
    """
    if TOR_PATH is None or not text or not page_nums or not TOR_PAGE_TEXTS:
        return None
    needle_raw = _normalize_for_search(text)
    if len(needle_raw) < 6:
        return None
    try:
        doc = fitz.open(TOR_PATH)
    except Exception:
        return None

    next_section = _next_sibling_section(section) if section else None

    # Try full needle, then progressively shorter prefixes
    for length in (90, 60, 40, 25, 15):
        head = needle_raw[:length].strip()
        if not head:
            continue
        head_n = normalize_thai_text(head)
        if not head_n:
            continue
        for pno in page_nums:
            if pno < 1 or pno > len(doc):
                continue
            page_text = TOR_PAGE_TEXTS.get(pno, "")
            if not page_text:
                continue
            if head_n in page_text:
                page = doc[pno - 1]
                bounds = _section_y_bounds_on_page(page, section, next_section)
                rects = _try_get_rects(page, head)
                rects = _filter_rects_by_bounds(rects, bounds)
                if not rects:
                    for sub_len in (30, 20, 12):
                        if sub_len >= len(head):
                            continue
                        rects = _try_get_rects(page, head[:sub_len])
                        rects = _filter_rects_by_bounds(rects, bounds)
                        if rects:
                            break
                if not rects:
                    # Token-rect fallback — anchor on the rarest token to
                    # pin a single line. Generic words ("Firewall", "Power")
                    # repeat across the section header + every sub-item; an
                    # uncommon token like "Throughput" or a numeric+unit
                    # appears only on the actual spec line.
                    token_hits: dict = {}
                    for tok in _extract_search_tokens(needle_raw)[:8]:
                        rs = _filter_rects_by_bounds(
                            _try_get_rects(page, tok), bounds)
                        if rs:
                            token_hits[tok] = rs
                    rects = _anchor_cluster_by_rarest(token_hits, line_span=14)
                if not rects and bounds:
                    # If page contains the phrase but every match is outside
                    # the section's y-range, skip this page (it's a different
                    # section's repetition of the phrase) — keep looking on
                    # other pages.
                    continue
                return {"page": pno,
                        "rects": [[r.x0, r.y0, r.x1, r.y1] for r in rects],
                        "needle": head}

    # Phrase fallback failed — try distinctive-token scoring
    tokens = _extract_search_tokens(needle_raw)
    if tokens:
        best = None  # (page, score, rects)
        for pno in page_nums:
            if pno < 1 or pno > len(doc):
                continue
            page_text = TOR_PAGE_TEXTS.get(pno, "")
            if not page_text:
                continue
            score = 0
            matched_tokens = []
            for tok in tokens:
                if normalize_thai_text(tok) in page_text:
                    score += 1
                    matched_tokens.append(tok)
            if score and (best is None or score > best[1]):
                best = (pno, score, matched_tokens)
        if best and best[1] >= max(2, min(3, len(tokens))):
            page = doc[best[0] - 1]
            bounds = _section_y_bounds_on_page(page, section, next_section)
            token_hits: dict = {}
            for tok in best[2][:6]:
                rs = _filter_rects_by_bounds(_try_get_rects(page, tok), bounds)
                if rs:
                    token_hits[tok] = rs
            rects = _anchor_cluster_by_rarest(token_hits, line_span=14)
            if rects:
                return {"page": best[0],
                        "rects": [[r.x0, r.y0, r.x1, r.y1] for r in rects],
                        "needle": " + ".join(best[2])}
    return None


def _chapter_root(section: str | None) -> str | None:
    """Top-level chapter, e.g. "5.1.2" → "5"."""
    if not section:
        return None
    return section.split(".")[0]


def find_in_tor(text: str, section: str | None = None) -> dict | None:
    """Search TOR PDF for `text`, biased to the row's section.

    Strategy (in priority order):
      1. Within the row's exact section range (e.g. 5.1.2 → pages 12–13).
      2. Section range widened by ±3 pages.
      3. From section start forward to end of the chapter root (e.g. all of 5).
      4. **No** full-document fallback — wrong-chapter matches are worse than
         no match. If nothing is found, default to the section's first page.
    """
    global TOR_PATH
    if TOR_PATH is None:
        TOR_PATH = find_tor_pdf()
    if TOR_PATH is None or not text:
        return None
    try:
        doc = fitz.open(TOR_PATH)
    except Exception:
        return None
    n_pages = len(doc)

    rng = get_tor_section_range(section)
    chapter_rng = get_tor_section_range(_chapter_root(section))

    page_groups: list[list[int]] = []
    if rng:
        s, e = rng
        # 1) exact section pages, in order
        page_groups.append(list(range(s, e + 1)))
        # 2) widen FORWARD only (content often spills past the section heading
        #    of the next sub-section, but rarely appears before the heading)
        page_groups.append(list(range(s, min(n_pages, e + 5) + 1)))
    if chapter_rng:
        cs, ce = chapter_rng
        # 3) forward from section start (or chapter start) to chapter end
        forward_start = rng[0] if rng else cs
        page_groups.append(list(range(forward_start, min(n_pages, ce) + 1)))
        # 4) backward within chapter, only if forward yielded nothing
        if rng:
            page_groups.append(list(range(cs, rng[0])))
        else:
            page_groups.append(list(range(cs, min(n_pages, ce) + 1)))
    if not page_groups:
        # Section unknown — search full doc as low-confidence fallback
        page_groups.append(list(range(1, n_pages + 1)))

    seen: set[int] = set()
    for group in page_groups:
        fresh = [p for p in group if p not in seen]
        seen.update(group)
        if not fresh:
            continue
        result = _search_tor_in_pages(text, fresh, section=section)
        if result:
            result["section_range"] = rng
            return result

    # Nothing matched within the chapter — return the section's first page
    # so the user lands in the right neighbourhood (no highlight).
    if rng:
        return {"page": rng[0], "rects": [], "section_range": rng,
                "fallback": "section_start"}
    return None


def render_tor_page(page_num: int, highlight_rects: list | None = None,
                    dpi: int = 110) -> bytes:
    if TOR_PATH is None:
        return b""
    doc = fitz.open(TOR_PATH)
    if page_num < 1 or page_num > len(doc):
        page_num = max(1, min(len(doc), page_num))
    page = doc[page_num - 1]
    pix = page.get_pixmap(dpi=dpi, annots=False)
    img_bytes = pix.tobytes("png")
    if not highlight_rects:
        return img_bytes
    im = Image.open(io.BytesIO(img_bytes)).convert("RGBA")
    overlay = Image.new("RGBA", im.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)
    scale = dpi / 72.0
    for r in highlight_rects:
        x0, y0, x1, y1 = r[0] * scale, r[1] * scale, r[2] * scale, r[3] * scale
        # yellow translucent fill + outline
        draw.rectangle([x0, y0, x1, y1], fill=(255, 230, 0, 90),
                       outline=(255, 180, 0, 220), width=3)
    im2 = Image.alpha_composite(im, overlay)
    buf = io.BytesIO()
    im2.convert("RGB").save(buf, "PNG")
    buf.seek(0)
    return buf.read()


# ---------------------------------------------------------------------------
# Auto checks (light — flag obvious problems without judging content)
# ---------------------------------------------------------------------------

SOFTWARE_KEYWORDS = ["ซอฟต์แวร์", "Software", "Package Application"]
INSTALL_KEYWORDS = ["งานติดตั้ง", "ติดตั้งและทดสอบ", "งานเดินสาย"]
COMPARE_PHRASES = ["หรือดีกว่า", "ไม่น้อยกว่า", "ไม่น้อยไปกว่า", "ไม่มากกว่า", "ต้องสามารถ", "จะต้อง"]


def auto_check_row(row: dict) -> list[dict]:
    """Return list of {level, message} flags."""
    flags = []
    b, c, d = (row.get("B") or ""), (row.get("C") or ""), (row.get("D") or "")
    parsed = row.get("parsed", {})
    # Col D format
    if d and parsed.get("type") == "empty":
        flags.append({"level": "warn", "msg": "Col D ไม่อยู่ใน 6 รูปแบบที่กำหนด"})
    # Software/install rows should be commitment + no PDF rect
    is_software = any(k in b for k in SOFTWARE_KEYWORDS) or any(k in c for k in SOFTWARE_KEYWORDS) if (b or c) else False
    is_install = any(k in b for k in INSTALL_KEYWORDS) or any(k in c for k in INSTALL_KEYWORDS) if (b or c) else False
    if (is_software or is_install) and parsed.get("type") not in ("commitment", "empty"):
        flags.append({"level": "info", "msg": f"Software/install row — Col D ควรเป็น 'ยินดีปฏิบัติตามข้อกำหนด' (พบ: {parsed.get('type')})"})
    # Col C still has compare phrases?
    if c:
        for ph in COMPARE_PHRASES:
            if ph in c:
                flags.append({"level": "warn", "msg": f"Col C ยังมีคำเปรียบเทียบ: '{ph}'"})
    # Col B copy from TOR — Col B should have leading whitespace (TOR indentation), Col C should not
    # gentle hint only
    if b and c and isinstance(b, str) and isinstance(c, str):
        if b.lstrip() == c.lstrip() and parsed.get("type") in ("equivalent", "higher", "brand_model"):
            # Col C identical to Col B (after trim) — not necessarily wrong but worth noting
            pass
    # PDF resolution
    if parsed.get("type") in ("equivalent", "higher", "filename_format", "model_only") and not row.get("pdf_rel"):
        flags.append({"level": "warn", "msg": "ไม่พบ PDF ที่ Col D อ้างอิง"})
    # Catalog exists but Col D is still empty (user added catalogs but hasn't
    # updated the spreadsheet yet — we resolve the PDF via folder convention)
    if row.get("needs_col_d"):
        flags.append({"level": "warn",
                      "msg": "พบ catalog ใน folder แต่ Col D ว่าง — ต้องเติม Col D ให้ตรงตาม convention"})
    return flags


# ---------------------------------------------------------------------------
# Status persistence
# ---------------------------------------------------------------------------

def load_status() -> dict:
    """Verification status now lives in SQLite (canonical source). The legacy
    verification_status.json is migrated on boot and kept as backup."""
    if db.db_path() is not None:
        try:
            return db.get_all_status()
        except Exception as e:
            sys.stderr.write(f"WARN: db get_all_status: {e}\n")
    # Fallback: read legacy JSON if DB isn't ready yet
    if STATUS_PATH.exists():
        try:
            return json.loads(STATUS_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def save_status(data: dict) -> None:
    """Backup write to JSON for compat with manual diffs. Not the source of
    truth — DB is."""
    try:
        STATUS_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2),
                               encoding="utf-8")
    except Exception as e:
        sys.stderr.write(f"WARN: cannot save status JSON: {e}\n")


# ---------------------------------------------------------------------------
# PDF rendering
# ---------------------------------------------------------------------------

def safe_iter_annots(page) -> list:
    """Iterate annotations skipping ones with broken appearance streams.

    PyMuPDF's ``page.annots()`` generator can fail entirely (returning zero
    annots) when the page mixes drawable annots with broken Image/Form
    annotations — common in our catalog PDFs. To stay robust we bypass the
    generator and call ``load_annot(xref)`` per entry from ``annot_xrefs()``
    with individual try/except. A single bad annot only loses itself, never
    the rest of the page.
    """
    out = []
    try:
        xrefs = page.annot_xrefs()
    except Exception:
        # fall back to the old generator path if annot_xrefs is unavailable
        try:
            gen = page.annots()
            if gen is None:
                return out
            while True:
                try:
                    ann = next(gen)
                except StopIteration:
                    break
                except Exception:
                    continue
                if ann is None:
                    continue
                out.append(ann)
        except Exception:
            pass
        return out

    for entry in xrefs:
        # entry is typically (xref, anntype, name); fall back gracefully
        if isinstance(entry, tuple):
            xref = entry[0]
        else:
            try:
                xref = int(entry)
            except Exception:
                continue
        if not xref:
            # xref=0 means an inline annot — handled by parse_inline_annots
            continue
        try:
            ann = page.load_annot(int(xref))
        except Exception:
            continue
        if ann is None:
            continue
        out.append(ann)
    return out


# ---------------------------------------------------------------------------
# Inline annotation parser
#
# Some PDFs (notably the เสา catalogs in 5.1.3.2 / 5.1.4.3 / 5.1.6.3) embed
# their annotations DIRECTLY in the page's /Annots array as literal
# dictionaries instead of indirect references. PyMuPDF's standard
# Page.annots() iterator skips these (xref reported as 0), so they were
# invisible to the highlight matcher and to the front-end annotation list.
# This parser extracts them by reading the raw /Annots value and decoding
# Subtype, Rect, and Contents fields out of each <<...>> block.
# ---------------------------------------------------------------------------

def _split_annots_array_items(arr_text: str) -> list[tuple[str, str]]:
    """Split a page /Annots array into ordered items, preserving the
    distinction between inline blocks and indirect references.

    Returns list of (kind, text) where kind ∈ {'ref', 'inline'} so callers
    can splice and rebuild the array correctly when removing an inline
    block (the indirect refs must keep their positions).
    """
    s = arr_text.strip()
    if s.startswith("["):
        s = s[1:]
    if s.endswith("]"):
        s = s[:-1]
    items: list[tuple[str, str]] = []
    i = 0
    L = len(s)
    while i < L:
        c = s[i]
        if c.isspace():
            i += 1
            continue
        if c == "<" and i + 1 < L and s[i + 1] == "<":
            depth = 1
            j = i + 2
            while j + 1 < L and depth > 0:
                if s[j] == "<" and s[j + 1] == "<":
                    depth += 1; j += 2
                elif s[j] == ">" and s[j + 1] == ">":
                    depth -= 1; j += 2
                else:
                    j += 1
            items.append(("inline", s[i:j]))
            i = j
        else:
            m = re.match(r"\d+\s+\d+\s+R", s[i:])
            if m:
                items.append(("ref", m.group(0)))
                i += m.end()
            else:
                i += 1
    return items


def _delete_inline_annot(doc, page, inline_index: int) -> bool:
    """Remove the N-th inline annotation from this page's /Annots array.

    PyMuPDF can't address inline annots via load_annot (they have xref=0),
    so we rewrite the array string in place with xref_set_key.
    """
    try:
        page_xref = page.xref
        kind, value = doc.xref_get_key(page_xref, "Annots")
    except Exception:
        return False
    if kind != "array" or not value:
        return False
    items = _split_annots_array_items(value)
    seen_inline = -1
    target_pos = -1
    for k, (knd, _val) in enumerate(items):
        if knd == "inline":
            seen_inline += 1
            if seen_inline == inline_index:
                target_pos = k
                break
    if target_pos < 0:
        return False
    items.pop(target_pos)
    new_value = "[" + " ".join(v for _, v in items) + "]"
    try:
        doc.xref_set_key(page_xref, "Annots", new_value)
    except Exception:
        return False
    return True


def _split_dict_blocks(arr_text: str) -> list[str]:
    """Split '[<<...>><<...>>]' into ['<<...>>', '<<...>>'] at depth 1."""
    s = arr_text.strip()
    if s.startswith("["):
        s = s[1:]
    if s.endswith("]"):
        s = s[:-1]
    blocks: list[str] = []
    depth = 0
    start = -1
    i = 0
    L = len(s)
    while i < L:
        if i + 1 < L and s[i] == "<" and s[i + 1] == "<":
            if depth == 0:
                start = i
            depth += 1
            i += 2
        elif i + 1 < L and s[i] == ">" and s[i + 1] == ">":
            depth -= 1
            i += 2
            if depth == 0 and start >= 0:
                blocks.append(s[start:i])
                start = -1
        else:
            i += 1
    return blocks


def _decode_pdf_string_literal(s: str) -> str:
    """Decode a (literal-string) PDF value with backslash escapes."""
    out = []
    i = 0
    L = len(s)
    while i < L:
        c = s[i]
        if c == "\\" and i + 1 < L:
            nxt = s[i + 1]
            if nxt in "()\\":
                out.append(nxt); i += 2; continue
            if nxt == "n": out.append("\n"); i += 2; continue
            if nxt == "r": out.append("\r"); i += 2; continue
            if nxt == "t": out.append("\t"); i += 2; continue
            if nxt.isdigit():
                # octal escape, up to 3 digits
                j = i + 1
                while j < L and j - i <= 3 and s[j].isdigit():
                    j += 1
                try:
                    out.append(chr(int(s[i + 1:j], 8)))
                except Exception:
                    pass
                i = j
                continue
            out.append(nxt); i += 2
        else:
            out.append(c); i += 1
    return "".join(out)


def _decode_pdf_hex_string(h: str) -> str:
    h = re.sub(r"\s+", "", h)
    if not h:
        return ""
    if len(h) % 2:
        h += "0"
    try:
        b = bytes.fromhex(h)
    except ValueError:
        return ""
    if b[:2] == b"\xfe\xff":
        try:
            return b[2:].decode("utf-16-be", errors="replace")
        except Exception:
            return ""
    if b[:2] == b"\xff\xfe":
        try:
            return b[2:].decode("utf-16-le", errors="replace")
        except Exception:
            return ""
    try:
        return b.decode("latin-1", errors="replace")
    except Exception:
        return ""


_RECT_RE = re.compile(
    r"/Rect\s*\[\s*([\-\d.]+)\s+([\-\d.]+)\s+([\-\d.]+)\s+([\-\d.]+)\s*\]"
)
_SUBTYPE_RE = re.compile(r"/Subtype\s*/(\w+)")
_CONTENTS_LITERAL_RE = re.compile(r"/Contents\s*\((.*?)(?<!\\)\)", re.DOTALL)
_CONTENTS_HEX_RE = re.compile(r"/Contents\s*<([0-9A-Fa-f\s]+)>")


def _parse_one_inline_annot(block: str) -> dict | None:
    """Parse a single <<...>> annotation dict block into our metadata shape."""
    if "/Type/Annot" not in block and "/Type /Annot" not in block:
        return None
    m_subtype = _SUBTYPE_RE.search(block)
    m_rect = _RECT_RE.search(block)
    if not m_subtype or not m_rect:
        return None
    out = {
        "type": m_subtype.group(1),
        "rect": [float(m_rect.group(i)) for i in (1, 2, 3, 4)],
        "contents": "",
    }
    # Hex string takes precedence (Thai content is usually hex-encoded)
    m_hex = _CONTENTS_HEX_RE.search(block)
    if m_hex:
        out["contents"] = _decode_pdf_hex_string(m_hex.group(1))
    else:
        m_lit = _CONTENTS_LITERAL_RE.search(block)
        if m_lit:
            out["contents"] = _decode_pdf_string_literal(m_lit.group(1))
    return out


def parse_inline_annots(doc, page) -> list[dict]:
    """Return inline-annotation dicts for a page (those without xrefs).

    PDF annotation /Rect values are in user-space (Y axis goes BOTTOM-UP in
    the standard PDF coordinate system). PyMuPDF's regular ``Annot.rect``
    auto-transforms to page-space (top-down) using
    ``page.transformation_matrix``; we replicate that transformation here so
    the rects we return are directly comparable to image-pixel coords.
    """
    try:
        page_xref = page.xref
        kind, value = doc.xref_get_key(page_xref, "Annots")
    except Exception:
        return []
    if kind != "array" or not value:
        return []
    try:
        mat = page.transformation_matrix
    except Exception:
        mat = None
    out = []
    for block in _split_dict_blocks(value):
        ann = _parse_one_inline_annot(block)
        if ann is None:
            continue
        if mat is not None:
            x0, y0, x1, y1 = ann["rect"]
            try:
                p0 = fitz.Point(x0, y0) * mat
                p1 = fitz.Point(x1, y1) * mat
                ann["rect"] = [
                    min(p0.x, p1.x), min(p0.y, p1.y),
                    max(p0.x, p1.x), max(p0.y, p1.y),
                ]
            except Exception:
                pass
        out.append(ann)
    return out


def all_annots_on_page(doc, page, page_num: int) -> list[dict]:
    """Combine real (xref'd) annotations and inline ones into a single list of
    dicts shaped like the rest of the codebase expects. Real annots come first
    so callers that key by xref don't lose them."""
    out: list[dict] = []
    seen_signatures: set[tuple] = set()  # to deduplicate
    for ann in safe_iter_annots(page):
        try:
            xref = ann.xref
            t = ann.type[1]
            r = [round(c, 2) for c in ann.rect]
            c = ann.info.get("content", "") or ""
        except Exception:
            continue
        sig = (t, tuple(r), c)
        seen_signatures.add(sig)
        out.append({
            "xref": int(xref),
            "page": page_num,
            "type": t,
            "rect": r,
            "contents": c,
        })
    inline_idx = 0
    for ann in parse_inline_annots(doc, page):
        r = [round(c, 2) for c in ann["rect"]]
        sig = (ann["type"], tuple(r), ann["contents"])
        # Even if the signature matches a real annot we keep the inline_index
        # counter advancing — its value must mirror the position of the inline
        # block in the page's /Annots array for delete-path lookups.
        if sig in seen_signatures:
            inline_idx += 1
            continue
        out.append({
            "xref": 0,            # 0 = inline (not editable via load_annot)
            "page": page_num,
            "type": ann["type"],
            "rect": r,
            "contents": ann["contents"],
            "_inline": True,
            "inline_index": inline_idx,
        })
        inline_idx += 1
    return out


# Patterns used to parse annotation labels into structured fields:
#   "5.1.2 ข้อ 3) ข้อย่อย 1."   → {item: 3, subitem: 1}
#   "5.1.2 ข้อ 3)"              → {item: 3, subitem: None}
#   "5.1.2 ข้อย่อย 4."          → {item: None, subitem: 4}
#   "ยี่ห้อ" / "รุ่น"            → {kind: 'brand_or_model', literal: ...}
_LABEL_ITEM_RE    = re.compile(r"ข้อ\s*(\d+)\s*\)")
_LABEL_SUBITEM_RE = re.compile(r"ข้อย่อย\s*(\d+)")
_LABEL_LITERALS   = ("ยี่ห้อ", "รุ่น")


def _parse_label(s: str) -> dict:
    """Parse query OR annotation content into structured fields.

    Returns ``{}`` for unrecognized strings (caller can fall back to substring).
    """
    out: dict = {}
    if not s:
        return out
    s = s.strip()
    # Brand/model literal (label content is typically just "ยี่ห้อ" or "รุ่น")
    if s in _LABEL_LITERALS:
        return {"kind": "literal", "literal": s}
    item = _LABEL_ITEM_RE.search(s)
    sub  = _LABEL_SUBITEM_RE.search(s)
    if item:
        out["item"] = int(item.group(1))
    if sub:
        out["subitem"] = int(sub.group(1))
    if out:
        out["kind"] = "section_label"
    return out


def _match_annot_label(query: str, content: str) -> bool:
    """Decide whether an annotation's content matches a highlight query.

    Goals:
      • exact digit match (so "ข้อย่อย 1" doesn't accidentally hit "ข้อย่อย 10")
      • when the query is a sub-item, also light up the parent rect
        ("ข้อ X)") so the user gets context — but NOT siblings
      • when the query is brand/model, only literal labels match (the paired
        Square gets added separately by spatial pairing in the caller)
    """
    qf = _parse_label(query)
    cf = _parse_label(content)

    # Brand/model literal — exact only
    if qf.get("kind") == "literal":
        return cf.get("kind") == "literal" and cf["literal"] == qf["literal"]

    # If we couldn't parse the query as a section label, fall back to
    # boundary-aware substring matching (e.g. tag content "5.1.2(2-model)").
    if qf.get("kind") != "section_label":
        if query == content:
            return True
        # Use word-boundary check so "ข้อย่อย 1" doesn't match "ข้อย่อย 10"
        # when callers pass raw substrings.
        try:
            return re.search(re.escape(query) + r"(?!\d)", content) is not None
        except re.error:
            return query in content

    if cf.get("kind") != "section_label":
        return False

    # Now both sides parsed.
    q_item, q_sub = qf.get("item"), qf.get("subitem")
    c_item, c_sub = cf.get("item"), cf.get("subitem")

    # SUB-ITEM QUERY (e.g., "ข้อ 3) ข้อย่อย 1.")
    if q_sub is not None:
        if c_sub is None:
            # Light up the parent rect of the SAME item (gives visual context)
            if q_item is not None and c_item == q_item:
                return True
            return False
        # Both have subitem — must match exactly
        if c_sub != q_sub:
            return False
        # If query specifies item, content's item must match too
        if q_item is not None and c_item is not None and c_item != q_item:
            return False
        return True

    # PARENT-ONLY QUERY ("ข้อ X)")
    if q_item is not None:
        if c_item != q_item:
            return False
        # Match the parent rect itself (no subitem in content). Skipping
        # subitem rects keeps the highlight focused.
        return c_sub is None

    return False


def render_pdf_page_png(pdf_path: Path, page_num: int, dpi: int = 130,
                        highlight: str | None = None,
                        no_annots: bool = False) -> tuple[bytes, dict]:
    """Render PDF page → (png_bytes, info).

    info = {"y0": float|None, "y1": float|None}  — image-pixel Y range of
    matched annotations (None if no highlight match).

    When ``no_annots`` is True the page is rendered WITHOUT baked annotations
    so the frontend can draw an interactive SVG overlay on top.
    """
    info: dict = {"y0": None, "y1": None}
    doc = fitz.open(pdf_path)
    if page_num < 1 or page_num > len(doc):
        page_num = max(1, min(len(doc), page_num))
    page = doc[page_num - 1]

    # Render page (annotations included unless explicitly suppressed)
    pix = page.get_pixmap(dpi=dpi, annots=not no_annots)
    img_bytes = pix.tobytes("png")

    if no_annots or not highlight:
        return img_bytes, info

    # Highlight may be a single label or "|"-separated alternatives
    queries = [q.strip() for q in highlight.split("|") if q.strip()]

    # Collect ALL annotations on this page — both real (xref'd) and inline
    # (literal-dict) so the matcher can find e.g. เสา catalog rects.
    doc_for_annots = page.parent
    annots = all_annots_on_page(doc_for_annots, page, page_num)

    # Find matching annotations using a structured, boundary-aware matcher.
    matched: list[dict] = []
    seen_ids: set[int] = set()
    for ann in annots:
        c = (ann.get("contents") or "").strip()
        if not c:
            continue
        for q in queries:
            if _match_annot_label(q, c):
                aid = id(ann)
                if aid not in seen_ids:
                    matched.append(ann)
                    seen_ids.add(aid)
                break

    # For each matched FreeText, pair with the nearest Square (by minimum
    # rect-edge distance). Handles all label layouts: right of square, below
    # square (เสา P2), above, or overlapping.
    def _rect_edge_distance(r1: list, r2: list) -> float:
        x_gap = max(0.0, r1[0] - r2[2], r2[0] - r1[2])
        y_gap = max(0.0, r1[1] - r2[3], r2[1] - r1[3])
        return (x_gap * x_gap + y_gap * y_gap) ** 0.5

    paired: list[dict] = []
    for m in matched:
        if m.get("type") != "FreeText":
            continue
        best = None
        best_d = 1e9
        for o in annots:
            if o.get("type") != "Square":
                continue
            d_ = _rect_edge_distance(m["rect"], o["rect"])
            if d_ < best_d:
                best_d, best = d_, o
        # Threshold: 60pt — close enough to be the labelled square but far
        # enough to not pair with random squares elsewhere on the page.
        if best is not None and best_d <= 60.0 and best not in matched and best not in paired:
            paired.append(best)

    if not matched and not paired:
        return img_bytes, info

    # Overlay highlights + collect Y range
    im = Image.open(io.BytesIO(img_bytes)).convert("RGBA")
    overlay = Image.new("RGBA", im.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)
    scale = dpi / 72.0
    ys: list[float] = []
    for ann in matched + paired:
        r = ann.get("rect")
        if not r or len(r) < 4:
            continue
        x0, y0, x1, y1 = r[0] * scale, r[1] * scale, r[2] * scale, r[3] * scale
        ys.append(y0); ys.append(y1)
        for w, alpha in [(8, 60), (5, 100), (2, 220)]:
            draw.rectangle([x0 - w, y0 - w, x1 + w, y1 + w],
                           outline=(255, 220, 0, alpha), width=w)
    if ys:
        info["y0"] = min(ys)
        info["y1"] = max(ys)
    im2 = Image.alpha_composite(im, overlay)
    buf = io.BytesIO()
    im2.convert("RGB").save(buf, "PNG")
    buf.seek(0)
    return buf.read(), info


# ---------------------------------------------------------------------------
# PDF editing — apply edits + per-PDF versioning
# ---------------------------------------------------------------------------

PDF_HISTORY = OUTPUT / "_pdf_history"


def _flatten_rel(rel: str) -> str:
    """Turn a relative PDF path into a single safe directory name."""
    return re.sub(r"[^\w\-]", "_", rel)[:160]


def snapshot_pdf(pdf_path: Path, tag: str = "") -> Path | None:
    """Copy the PDF into _pdf_history/<flat>/<timestamp>.pdf."""
    try:
        rel = str(pdf_path.relative_to(OUTPUT))
    except ValueError:
        return None
    flat = _flatten_rel(rel)
    hist_dir = PDF_HISTORY / flat
    hist_dir.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    name = f"{ts}{('_' + tag) if tag else ''}.pdf"
    target = hist_dir / name
    try:
        shutil.copy2(pdf_path, target)
    except Exception as e:
        sys.stderr.write(f"[snapshot] {e}\n")
        return None
    return target


def list_pdf_snapshots(pdf_path: Path) -> list[dict]:
    try:
        rel = str(pdf_path.relative_to(OUTPUT))
    except ValueError:
        return []
    flat = _flatten_rel(rel)
    hist_dir = PDF_HISTORY / flat
    if not hist_dir.exists():
        return []
    out = []
    for p in sorted(hist_dir.glob("*.pdf"), reverse=True):
        st = p.stat()
        out.append({
            "name": p.name,
            "id": p.stem,
            "size": st.st_size,
            "mtime": st.st_mtime,
        })
    return out


def restore_pdf_snapshot(pdf_path: Path, snapshot_id: str) -> bool:
    rel = str(pdf_path.relative_to(OUTPUT))
    flat = _flatten_rel(rel)
    src = PDF_HISTORY / flat / f"{snapshot_id}.pdf"
    if not src.exists():
        return False
    # Snapshot current state before overwriting
    snapshot_pdf(pdf_path, tag="pre_restore")
    shutil.copy2(src, pdf_path)
    return True


# Standard appearance string for newly-created Square / FreeText annots,
# matching the convention defined in SKILL.md (red, Helvetica-Bold 9pt).
_DA_BRAND   = "1 0 0 rg /HeBo 9 Tf"
_DS_BRAND   = "font: bold Helvetica,sans-serif 9.0pt; text-align:left; color:#FF0000"


def apply_pdf_edits(pdf_path: Path, edits: list[dict]) -> dict:
    """Apply a list of edit ops to a PDF. Snapshots before write.

    Each edit is one of:
      {action: "update",  xref: int, rect?: [x0,y0,x1,y1], contents?: str}
      {action: "delete",  xref: int}
      {action: "create",  page: int, type: "Square"|"FreeText",
                          rect: [...], contents?: str, fontsize?: int}
    """
    if not edits:
        return {"applied": 0, "errors": 0, "snapshot": None}

    snap = snapshot_pdf(pdf_path, tag="pre_edit")

    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        return {"applied": 0, "errors": len(edits), "error": str(e)}

    # Map xref → page_index without keeping annot references alive (those go
    # stale after intervening operations). When we need an annot we re-load
    # it fresh via page.load_annot(int_xref).
    xref_to_page: dict[int, int] = {}
    for pno in range(len(doc)):
        try:
            for entry in doc[pno].annot_xrefs():
                xref = entry[0] if isinstance(entry, tuple) else int(entry)
                xref_to_page[int(xref)] = pno
        except Exception:
            continue

    # Cache page objects so the annots returned by load_annot stay bound
    # (PyMuPDF detaches Annot objects when their page goes out of scope).
    _pages: dict[int, fitz.Page] = {}

    def _get_page(pno: int):
        if pno not in _pages:
            _pages[pno] = doc[pno]
        return _pages[pno]

    def _find_annot(xref: int):
        pno = xref_to_page.get(xref)
        if pno is None:
            return None, None, None
        page = _get_page(pno)
        try:
            ann = page.load_annot(int(xref))
        except Exception:
            return None, None, None
        return pno, page, ann

    applied = 0
    errors = 0
    error_msgs: list[str] = []
    new_xrefs: dict[str, int] = {}  # client_id → real xref (for "create" results)

    for edit in edits:
        try:
            action = edit.get("action")
            if action == "update":
                xref = int(edit["xref"])
                _pno, _page, ann = _find_annot(xref)
                if ann is None:
                    errors += 1
                    error_msgs.append(f"update: xref {xref} not found")
                    continue
                if "rect" in edit:
                    ann.set_rect(fitz.Rect(*edit["rect"]))
                if "contents" in edit:
                    info = ann.info
                    info["content"] = str(edit["contents"])
                    ann.set_info(info)
                # Re-assert standard appearance so legacy annotations
                # (saved with default colors / no explicit stroke) pick up
                # the red theme next time the user edits them.
                try:
                    kind = ann.type[1] if isinstance(ann.type, tuple) else str(ann.type)
                except Exception:
                    kind = ""
                if kind == "Square":
                    try:
                        ann.set_colors(stroke=(1, 0, 0))
                        ann.set_border(width=1)
                    except Exception: pass
                elif kind == "FreeText":
                    try:
                        # In PyMuPDF FreeText, "stroke" is the TEXT color
                        ann.set_colors(stroke=(1, 0, 0))
                    except Exception: pass
                ann.update()
                applied += 1

            elif action == "delete":
                xref = int(edit["xref"])
                pno, page, ann = _find_annot(xref)
                if ann is None:
                    errors += 1
                    error_msgs.append(f"delete: xref {xref} not found")
                    continue
                page.delete_annot(ann)
                # remove from map so future ops don't try to use it
                xref_to_page.pop(xref, None)
                applied += 1

            elif action == "delete_inline":
                # Remove an inline annotation (xref=0) by rewriting the page's
                # /Annots array string. inline_index identifies which inline
                # block to drop, preserving order of indirect refs.
                pno = int(edit["page"]) - 1
                if pno < 0 or pno >= len(doc):
                    errors += 1
                    error_msgs.append("delete_inline: bad page")
                    continue
                page = _get_page(pno)
                idx = int(edit.get("inline_index", -1))
                if idx < 0:
                    errors += 1
                    error_msgs.append("delete_inline: missing inline_index")
                    continue
                if _delete_inline_annot(doc, page, idx):
                    applied += 1
                else:
                    errors += 1
                    error_msgs.append(
                        f"delete_inline: page {pno + 1} idx {idx} not found")

            elif action == "create":
                pno = int(edit["page"]) - 1
                if pno < 0 or pno >= len(doc):
                    errors += 1
                    continue
                page = _get_page(pno)
                rect = fitz.Rect(*edit["rect"])
                ann_type = edit.get("type", "Square")
                if ann_type == "Square":
                    a = page.add_rect_annot(rect)
                    a.set_colors(stroke=(1, 0, 0))
                    a.set_border(width=1)
                    info = a.info
                    info["content"] = str(edit.get("contents", ""))
                    a.set_info(info)
                    a.update()
                elif ann_type == "FreeText":
                    # Compute fontsize from rect height if not provided —
                    # matches the SVG overlay's clamp formula so what the
                    # user sees in edit mode is what they get on disk.
                    rect_h = max(0.0, rect.y1 - rect.y0)
                    default_fs = max(6.0, min(14.0, rect_h * 0.65))
                    fs_raw = edit.get("fontsize")
                    fs = float(fs_raw) if fs_raw not in (None, "") else default_fs
                    fs = max(6.0, min(14.0, fs))
                    # NOTE: PyMuPDF rejects border_color/fill_color unless
                    # rich_text=True. We deliberately render FreeText labels
                    # WITHOUT a border — the SVG overlay matches by drawing
                    # a transparent label rect (just the red text shows).
                    # The paired Square already provides the visible frame.
                    a = page.add_freetext_annot(
                        rect,
                        str(edit.get("contents", "")),
                        fontsize=fs,
                        fontname="helv",
                        text_color=(1, 0, 0),
                        align=0,
                    )
                    a.set_info({"content": str(edit.get("contents", ""))})
                    a.update()
                else:
                    errors += 1
                    continue
                client_id = edit.get("client_id")
                if client_id is not None:
                    new_xrefs[str(client_id)] = a.xref
                applied += 1

            else:
                errors += 1
                error_msgs.append(f"unknown action: {action}")
        except Exception as e:
            errors += 1
            error_msgs.append(f"{edit.get('action')}: {e}")
            sys.stderr.write(f"[apply_pdf_edits] {e}\n")

    # Save with garbage cleanup so removed annots actually leave the file
    tmp = pdf_path.with_suffix(".pdf.tmp")
    try:
        doc.save(str(tmp), garbage=4, clean=True, deflate=True)
    finally:
        doc.close()

    # Atomic-replace via os.open + write (avoids permission quirks on cloud
    # mounts like Google Drive)
    try:
        with open(tmp, "rb") as src:
            data = src.read()
        fd = os.open(str(pdf_path), os.O_WRONLY | os.O_TRUNC)
        try:
            os.write(fd, data)
        finally:
            os.close(fd)
    except Exception:
        # fallback: replace
        shutil.move(str(tmp), str(pdf_path))
    else:
        try:
            tmp.unlink()
        except Exception:
            pass

    return {
        "applied": applied,
        "errors": errors,
        "error_msgs": error_msgs,
        "snapshot": snap.name if snap else None,
        "new_xrefs": new_xrefs,
    }


# ---------------------------------------------------------------------------

def list_pdf_annots(pdf_path: Path) -> list[dict]:
    """Return annotations metadata for all pages.

    Includes both real (xref'd) and inline (literal-dict) annotations so the
    highlight matcher and frontend can see annotations from PDFs that embed
    their /Annots inline (e.g. the เสา catalogs).
    """
    doc = fitz.open(pdf_path)
    out: list[dict] = []
    for pno in range(len(doc)):
        page = doc[pno]
        out.extend(all_annots_on_page(doc, page, pno + 1))
    return out


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    # Template lives at app/server/templates/index.html (extracted from
    # the legacy INDEX_HTML embedded string in T2.1, 2026-05-10).
    return render_template("index.html")


@app.route("/manual")
def manual_page():
    """Render the user manual as a styled standalone page (TOC + search)."""
    return render_template("manual.html")


@app.route("/api/manual/raw")
def api_manual_raw():
    """Return the raw markdown of docs/MANUAL.md so the manual page can
    render it client-side via marked.js. Falls back to a stub message if
    the file is missing."""
    p = ROOT / "docs" / "MANUAL.md"
    if not p.exists():
        return Response(
            "# Manual not found\n\nExpected at `docs/MANUAL.md`.",
            mimetype="text/markdown; charset=utf-8",
        )
    return Response(p.read_text(encoding="utf-8"),
                    mimetype="text/markdown; charset=utf-8")


@app.route("/api/index")
def api_index():
    rows_payload = []
    for r in ROWS:
        # Derive structural role from Col B (mirrors detect_row_role) so the
        # frontend can branch on role without needing a separate API call.
        b_full = r.get("B") or ""
        b = b_full.strip()
        role = "unknown"
        if re.match(r"^\d+(?:\.\d+){2,3}\.\s+", b):
            role = "section_header"
        else:
            m_item = re.match(r"^(\d+)\s*\)\s*", b)
            m_sub = re.match(r"^(\d+)\s*\.\s*", b)
            if m_item and (not m_sub or m_item.end() <= m_sub.end()):
                role = "item"
            elif m_sub:
                role = "sub_item"
        rows_payload.append({
            "row": r["row"],
            "A": r["A"],
            "B": r["B"],
            "C": r["C"],
            "D": r["D"],
            "E": r["E"],
            "F": r["F"],
            "parsed": r["parsed"],
            "pdf_rel": r["pdf_rel"],
            "pdf_inherited": r.get("pdf_inherited", False),
            "section": r["section_inferred"],
            "role": role,
            "auto_flags": auto_check_row(r),
        })
    sections = sorted({r["section_inferred"] for r in ROWS if r.get("section_inferred")},
                      key=lambda s: [int(x) for x in re.findall(r"\d+", s)])
    # serialize tree (children dict → list, preserve insertion order)
    def _serialize(node):
        return {
            "key": node["key"],
            "label": node["label"],
            "type": node["type"],
            "rows": node["rows"],
            "children": [_serialize(c) for c in node["children"].values()],
        }
    return jsonify({
        "rows": rows_payload,
        "sections": sections,
        "tree": _serialize(build_tree()),
        "extra": EXTRA_REFS,
        "tor_available": TOR_PATH is not None or find_tor_pdf() is not None,
        "status": load_status(),
        "version_sync": get_version_sync_status(),
        "stats": {
            "total": len(ROWS),
            "with_pdf_ref": sum(1 for r in ROWS if r["pdf_rel"]),
            "by_type": _counter([r["parsed"].get("type") for r in ROWS]),
        },
    })


def _counter(seq):
    out = {}
    for s in seq:
        out[s] = out.get(s, 0) + 1
    return out


@app.route("/api/status", methods=["GET", "POST"])
def api_status():
    if request.method == "POST":
        data = request.get_json(silent=True) or {}
        try:
            row_num = int(data.get("row"))
        except (TypeError, ValueError):
            return jsonify({"error": "row required"}), 400
        # Write through DB (canonical) — this also writes an audit entry
        # whenever the status actually changes.
        entry = db.set_status(
            row_num,
            status=data.get("status") if "status" in data else None,
            notes=data.get("notes") if "notes" in data else None,
            actor="user",
        )
        # Mirror into legacy JSON for compatibility
        full = db.get_all_status()
        save_status(full)
        return jsonify({"ok": True, "entry": entry})
    return jsonify(load_status())


@app.route("/api/pdf_page")
def api_pdf_page():
    rel = request.args.get("rel")
    page = int(request.args.get("page", 1))
    dpi = int(request.args.get("dpi", 130))
    highlight = request.args.get("highlight")
    edit_mode = request.args.get("edit") == "1"
    # WYSIWYG: even in edit mode we bake annots into the rendered image so
    # the user sees the SAME thing they'd see in preview (yellow header
    # banners, custom-appearance labels, etc.). The SVG overlay only adds
    # invisible hit areas + handles for selected annots, plus visible
    # rect/text for newly-drawn (unsaved) ones. Suppressing annots
    # altogether (the old behaviour) made edit-mode look totally different
    # from preview because PyMuPDF's appearance streams (/AP) — including
    # backgrounds and custom fonts — were stripped.
    # An explicit ?bake=0 still forces stripping (used by the wizard when
    # actively editing existing annots so the user doesn't see ghost
    # copies during drag).
    bake_param = request.args.get("bake")
    if bake_param == "0":
        no_annots = True
    else:
        no_annots = False
    if not rel:
        abort(400, "rel required")
    p = (OUTPUT / rel).resolve()
    if not str(p).startswith(str(OUTPUT.resolve())):
        abort(403)
    if not p.exists():
        abort(404)
    png, info = render_pdf_page_png(p, page, dpi=dpi,
                                    highlight=highlight,
                                    no_annots=no_annots)
    resp = send_file(io.BytesIO(png), mimetype="image/png")
    if info.get("y0") is not None:
        resp.headers["X-Highlight-Y0"] = f"{info['y0']:.0f}"
        resp.headers["X-Highlight-Y1"] = f"{info['y1']:.0f}"
    resp.headers["Access-Control-Expose-Headers"] = "X-Highlight-Y0, X-Highlight-Y1"
    # Cache-bust on edit-mode changes by setting no-store
    if edit_mode:
        resp.headers["Cache-Control"] = "no-store"
    return resp


def _resolve_pdf_rel(rel: str) -> Path:
    if not rel:
        abort(400, "rel required")
    p = (OUTPUT / rel).resolve()
    if not str(p).startswith(str(OUTPUT.resolve())):
        abort(403)
    if not p.exists():
        abort(404)
    return p


@app.route("/api/pdf_save", methods=["POST"])
def api_pdf_save():
    """Apply a list of annotation edits to a PDF and snapshot the prior state.

    Body: {"rel": "...", "edits": [...]}
    """
    data = request.get_json(silent=True) or {}
    rel = data.get("rel")
    edits = data.get("edits") or []
    if not rel or not isinstance(edits, list):
        return jsonify({"error": "rel and edits[] required"}), 400
    p = _resolve_pdf_rel(rel)
    result = apply_pdf_edits(p, edits)
    try:
        db.log_audit(action="pdf_edit", target_type="pdf", target_id=rel,
                     details={"edits": len(edits),
                              "applied": result.get("applied", 0),
                              "errors": result.get("errors", 0),
                              "snapshot": result.get("snapshot")},
                     actor="user")
    except Exception: pass
    return jsonify({"ok": True, **result})


@app.route("/api/pdf_history")
def api_pdf_history():
    rel = request.args.get("rel")
    if not rel:
        abort(400)
    p = _resolve_pdf_rel(rel)
    return jsonify({
        "rel": rel,
        "snapshots": list_pdf_snapshots(p),
    })


@app.route("/api/pdf_restore", methods=["POST"])
def api_pdf_restore():
    data = request.get_json(silent=True) or {}
    rel = data.get("rel")
    snap_id = data.get("snapshot")
    if not rel or not snap_id:
        return jsonify({"error": "rel + snapshot required"}), 400
    p = _resolve_pdf_rel(rel)
    ok = restore_pdf_snapshot(p, snap_id)
    if not ok:
        return jsonify({"error": "snapshot not found"}), 404
    return jsonify({"ok": True})


@app.route("/api/pdf_snapshot", methods=["POST"])
def api_pdf_snapshot():
    """Manual snapshot trigger ('save version' button)."""
    data = request.get_json(silent=True) or {}
    rel = data.get("rel")
    tag = (data.get("tag") or "manual").strip()
    tag = re.sub(r"[^\w\-]", "_", tag)[:40]
    if not rel:
        return jsonify({"error": "rel required"}), 400
    p = _resolve_pdf_rel(rel)
    snap = snapshot_pdf(p, tag=tag)
    return jsonify({"ok": True, "name": snap.name if snap else None})


@app.route("/api/pdf_meta")
def api_pdf_meta():
    rel = request.args.get("rel")
    if not rel:
        abort(400)
    p = (OUTPUT / rel).resolve()
    if not str(p).startswith(str(OUTPUT.resolve())):
        abort(403)
    if not p.exists():
        abort(404)
    doc = fitz.open(p)
    meta = {
        "rel": rel,
        "pages": len(doc),
        "annots": list_pdf_annots(p),
        "page_sizes": [(round(doc[i].rect.width, 1), round(doc[i].rect.height, 1))
                       for i in range(len(doc))],
    }
    return jsonify(meta)


@app.route("/api/raw_pdf")
def api_raw_pdf():
    rel = request.args.get("rel")
    if not rel:
        abort(400)
    # Allow PDFs from output, TOR, BOQ
    candidates = [OUTPUT / rel, PROJECT / rel]
    for p in candidates:
        if p.exists() and p.is_file() and p.suffix.lower() == ".pdf":
            return send_file(p, mimetype="application/pdf")
    abort(404)


@app.route("/api/tor_page")
def api_tor_page():
    """Render TOR PDF page with highlight on Col B text from a given row."""
    row = int(request.args.get("row", 0))
    page = request.args.get("page")
    dpi = int(request.args.get("dpi", 110))

    target = next((r for r in ROWS if r["row"] == row), None)
    if not target:
        abort(404, "row not found")
    text = target.get("B") or target.get("C") or ""
    section = target.get("section_inferred")

    # Cache lookup
    if row in TOR_CACHE:
        info = TOR_CACHE[row]
    else:
        info = find_in_tor(text, section=section) or {"page": 1, "rects": []}
        TOR_CACHE[row] = info

    use_page = int(page) if page else info["page"]
    rects = info["rects"] if (use_page == info["page"]) else []
    png = render_tor_page(use_page, rects, dpi=dpi)
    if not png:
        abort(404, "TOR not found")

    resp = send_file(io.BytesIO(png), mimetype="image/png")
    resp.headers["X-TOR-Page"] = str(info["page"])
    resp.headers["X-TOR-Hits"] = str(len(info["rects"]))
    rng = info.get("section_range") or get_tor_section_range(section)
    if rng:
        resp.headers["X-Section-Range"] = f"{rng[0]}-{rng[1]}"
    if section:
        resp.headers["X-Section"] = section
    # image-pixel Y of first matched rect → for auto-scroll
    if rects:
        scale = dpi / 72.0
        y0 = min(r[1] for r in rects) * scale
        y1 = max(r[3] for r in rects) * scale
        resp.headers["X-Highlight-Y0"] = f"{y0:.0f}"
        resp.headers["X-Highlight-Y1"] = f"{y1:.0f}"
    # Allow client JS to read these custom headers (CORS isn't an issue
    # since same-origin, but some browsers still gate non-standard headers)
    resp.headers["Access-Control-Expose-Headers"] = (
        "X-TOR-Page, X-TOR-Hits, X-Highlight-Y0, X-Highlight-Y1, "
        "X-Section, X-Section-Range"
    )
    return resp


@app.route("/api/tor_meta")
def api_tor_meta():
    global TOR_PATH
    if TOR_PATH is None:
        TOR_PATH = find_tor_pdf()
    if TOR_PATH is None:
        return jsonify({"available": False})
    doc = fitz.open(TOR_PATH)
    return jsonify({
        "available": True,
        "filename": TOR_PATH.name,
        "pages": len(doc),
    })


@app.route("/api/row/col_d", methods=["POST"])
def api_row_col_d():
    """Inline Col D edit (from xlsx preview double-click). Records the
    change as HITL feedback so the learner can mine repeated patterns."""
    data = request.get_json(silent=True) or {}
    try:
        row_num = int(data["row"])
    except (TypeError, ValueError, KeyError):
        return jsonify({"ok": False, "error": "row required"}), 400
    new_d = (data.get("col_d") or "").strip()
    original = (data.get("original") or "").strip()

    row = next((r for r in ROWS if r["row"] == row_num), None)
    if not row:
        return jsonify({"ok": False, "error": "row not found"}), 404

    # Pre-snapshot for safety
    _run_version_cmd(["snap", f"pre-inline-row-{row_num}"], timeout=60)

    try:
        wb = openpyxl.load_workbook(XLSX_PATH)
        ws = wb.active
        ws.cell(row_num, 4).value = new_d
        tmp = XLSX_PATH.with_suffix(".xlsx.tmp")
        wb.save(str(tmp))
        with open(tmp, "rb") as f:
            buf = f.read()
        fd = os.open(str(XLSX_PATH), os.O_WRONLY | os.O_TRUNC)
        try:
            os.write(fd, buf)
        finally:
            os.close(fd)
        tmp.unlink(missing_ok=True)
    except Exception as e:
        return jsonify({"ok": False, "error": f"xlsx write failed: {e}"}), 500

    # Refresh + log
    load_rows()
    sync_db_from_memory()
    try:
        db.log_audit(action="col_d_inline_edit",
                     target_type="row", target_id=str(row_num),
                     before=original, after=new_d, actor="user")
    except Exception: pass
    try:
        learning.record_feedback(
            row_num=row_num,
            section=row.get("section_inferred"),
            input_b=row.get("B", "") or "",
            input_pdf_rel=row.get("pdf_rel"),
            input_role=detect_row_role(row_num).get("role", ""),
            input_filename=(Path(row["pdf_rel"]).name if row.get("pdf_rel") else None),
            suggested_c=row.get("C", "") or "",
            suggested_d=original,
            suggested_annots=[],
            confidence=0.0,
            generator="prior",
            provenance={},
            user_action="edited",
            final_c=row.get("C", "") or "",
            final_d=new_d,
            final_annots=[],
        )
    except Exception as e:
        sys.stderr.write(f"[col_d_inline] feedback: {e}\n")

    return jsonify({"ok": True, "row": row_num, "new_d": new_d})


@app.route("/api/row/col_d/suggest")
def api_row_col_d_suggest():
    """Phase B3: live autocomplete for the Col D inline editor.

    Returns up to 6 ranked candidates: the AI's auto-annotate proposal
    (top-priority), Col D shapes from same-section verified rows, and
    a couple of generic shape templates so users can tab-complete.

    Cheap path — no LLM call here. The expensive proposal comes from
    auto_annotate_plan() which is already cache-friendly via section
    + filename derivations.
    """
    try:
        row_num = int(request.args.get("row", "0"))
    except ValueError:
        return jsonify({"ok": False, "error": "row required"}), 400
    q = (request.args.get("q") or "").strip()
    row = next((r for r in ROWS if r["row"] == row_num), None)
    if not row:
        return jsonify({"ok": False, "error": "row not found"}), 404

    suggestions: list[dict] = []
    section = row.get("section_inferred") or ""

    # 1. AI proposal from auto_annotate_plan (rule+pattern, no LLM)
    try:
        plan = auto_annotate_plan(row_num)
        if plan.get("ok") and plan.get("proposed_d"):
            suggestions.append({
                "text": plan["proposed_d"],
                "kind": "ai",
                "label": "AI proposal",
                "confidence": round(float(plan.get("confidence") or 0.0), 2),
                "generator": plan.get("generator", "rules"),
            })
    except Exception:
        pass

    # 2. Col D values from verified rows in the same section root
    #    (e.g. for section 5.1.2, suggest Col D values from any verified
    #    5.1.* row). This gives users a "match the neighbor's style"
    #    completion. De-dupe on text.
    seen = {s["text"].strip() for s in suggestions}
    if section:
        section_root = ".".join(section.split(".")[:2])  # "5.1.2" → "5.1"
        try:
            statuses = db.get_all_status()
        except Exception:
            statuses = {}
        for r in ROWS:
            if r["row"] == row_num: continue
            sec = r.get("section_inferred") or ""
            if not sec.startswith(section_root): continue
            d = (r.get("D") or "").strip()
            if not d or d in seen: continue
            # Prefer rows with a non-unverified verdict
            st = (statuses.get(str(r["row"]), {}) or {}).get("status", "unverified")
            if st in ("pass", "need_fix"):  # known good/edited shapes
                suggestions.append({
                    "text": d, "kind": "neighbor",
                    "label": f"R{r['row']} · {sec}",
                    "confidence": 0.7 if st == "pass" else 0.5,
                    "generator": f"section:{sec}",
                })
                seen.add(d)
            if len(suggestions) >= 6: break

    # 3. Shape templates if we still have room — handy when a user
    #    types a section ref and wants to expand to canonical form.
    if len(suggestions) < 6:
        templates = [
            {"text": f"เอกสาร {section} ... หน้า ?", "kind": "shape",
             "label": "doc-ref (dot form)", "confidence": 0.3},
            {"text": "ยินดีปฏิบัติตามข้อกำหนด", "kind": "shape",
             "label": "commitment", "confidence": 0.3},
        ]
        for t in templates:
            if t["text"] in seen: continue
            suggestions.append(t)
            seen.add(t["text"])
            if len(suggestions) >= 6: break

    # If user has typed a query, prefer suggestions whose text contains q
    # (case-insensitive), but always keep AI proposal at top.
    if q:
        ql = q.casefold()
        def _score(s):
            t = s["text"].casefold()
            kind_w = {"ai": 100, "neighbor": 50, "shape": 10}.get(s["kind"], 0)
            match_w = 30 if ql in t else 0
            start_w = 20 if t.startswith(ql) else 0
            return -(kind_w + match_w + start_w + s.get("confidence", 0) * 10)
        suggestions.sort(key=_score)

    return jsonify({"ok": True, "row": row_num, "section": section,
                    "suggestions": suggestions[:6]})


@app.route("/api/row_context")
def api_row_context():
    """Return rows surrounding the target row for the xlsx preview pane."""
    row = int(request.args.get("row", 0))
    radius = int(request.args.get("radius", 6))
    idx = next((i for i, r in enumerate(ROWS) if r["row"] == row), -1)
    if idx < 0:
        abort(404)
    lo = max(0, idx - radius)
    hi = min(len(ROWS), idx + radius + 1)
    out = []
    for i in range(lo, hi):
        r = ROWS[i]
        out.append({
            "row": r["row"],
            "A": r["A"], "B": r["B"], "C": r["C"],
            "D": r["D"], "E": r["E"], "F": r["F"],
            "type": r["parsed"].get("type"),
        })
    return jsonify({"target": row, "rows": out})


@app.route("/api/refresh")
def api_refresh():
    """Re-scan filesystem + xlsx (in case user edited files)."""
    build_pdf_index()
    load_rows()
    collect_extra_refs()
    build_tor_section_index()
    index_tor_text()
    TOR_CACHE.clear()
    sync_db_from_memory()
    try:
        db.log_audit(action="refresh", target_type="project",
                     details={"rows": len(ROWS), "pdfs": len(PDF_INDEX)},
                     actor="user")
    except Exception: pass
    return jsonify({"ok": True, "rows": len(ROWS), "pdfs": len(PDF_INDEX),
                    "tor_sections": len(TOR_SECTION_INDEX),
                    "tor_pages_indexed": len(TOR_PAGE_TEXTS)})


# ---------------------------------------------------------------------------
# Auto-annotate pipeline
#
# For rows that have a resolved catalog PDF but empty Col D, we can:
#   1. Extract brand/model from the catalog filename
#   2. Detect the row's role (section header / item / sub-item)
#   3. For sub-items: text-search the PDF for distinctive Col B tokens to
#      pinpoint which page the spec lives on; compute a content rect.
#   4. Generate proposed Col D + (optional) PDF Square + FreeText annotations
#   5. Snapshot before writing, then commit changes to xlsx + PDF
#
# Operates in dry-run by default (returns a "plan") so the user can review
# each row before applying. Batch mode processes all rows with the
# `needs_col_d` flag.
# ---------------------------------------------------------------------------

# Tokens that look like brand names in filenames but actually describe the
# product category. Skipped when picking the brand.
_FILENAME_UNITS = {
    "kVA", "KVA", "kW", "kWh", "kHz", "MHz", "GHz", "Hz", "kV", "mV", "mA",
    "mm", "cm", "kg", "ms", "us", "ns", "Mbps", "Gbps", "DC", "AC", "VAC",
    "VDC", "ms.", "rpm", "RPM", "lm", "lx", "ID", "OD",
    # Generic product-category acronyms (NAS / PoE / AP / etc.)
    "NAS", "PoE", "POE", "AP", "IP", "USB", "HDMI", "RJ45", "SFP", "SFP+",
    "L2", "L3", "PC", "TB", "GB", "MB", "RAM", "ROM", "CPU", "GPU", "UPS",
    "NVR", "DVR", "CCTV", "PDU", "FO", "VCT", "EMT", "HDPE",
}

# Compare phrases per SKILL.md (used when generating Col C from Col B)
_COMPARE_PHRASES_STRIP = [
    ("หรือดีกว่า", ""), ("ไม่น้อยไปกว่า", ""), ("ไม่น้อยกว่า", ""),
    ("ไม่มากกว่า", ""), ("ต้องสามารถ", "สามารถ"), ("จะต้อง", ""),
]


def parse_brand_model_from_filename(stem: str) -> tuple[str, str]:
    """Heuristic: strip section + Thai description, take first Latin token as
    brand and the rest as model.

    Examples:
      '5.1.1.2. เครื่องคอมพิวเตอร์แม่ข่าย แบบที่ 2 Lenovo ThinkSystem SR630 V4'
        → ('Lenovo', 'ThinkSystem SR630 V4')
      '5.1.1.7. เครื่องสำรองไฟฟ้า ขนาด 10 kVA (...) cleanline t10k33lv2'
        → ('Cleanline', 't10k33lv2')
    """
    s = stem.strip()
    # Strip leading section number
    s = re.sub(r"^\d+(?:\.\d+){1,3}\.?\s+", "", s)
    # Strip "แบบที่ N" markers (variant labels)
    s = re.sub(r"แบบที่\s*\d+\s*", "", s)
    # Strip parenthesised phrases (often Thai notes)
    s = re.sub(r"\([^)]*\)", "", s)
    s = re.sub(r"\s+", " ", s).strip()

    tokens = s.split()
    brand_idx = -1
    for i, tok in enumerate(tokens):
        if tok in _FILENAME_UNITS:
            continue
        # Token must start with ≥2 ASCII letters (skips raw digits + units)
        m = re.match(r"^[A-Za-z]{2,}", tok)
        if m:
            brand_idx = i
            break
    if brand_idx < 0:
        return ("", "")
    brand = tokens[brand_idx]
    # Title-case the brand if it's all-lowercase (e.g. "cleanline" → "Cleanline")
    if brand and brand.islower():
        brand = brand[0].upper() + brand[1:]
    model = " ".join(tokens[brand_idx + 1:]).strip()
    return (brand, model)


def extract_section_name(stem: str, section: str) -> str:
    """Strip section prefix + brand/model tail from filename, leaving the
    Thai description (used in Col D between the catalog ref and the page).

    '5.1.1.2. เครื่องคอมพิวเตอร์แม่ข่าย แบบที่ 2 Lenovo ThinkSystem SR630 V4'
      → 'เครื่องคอมพิวเตอร์แม่ข่าย แบบที่ 2'
    """
    brand, model = parse_brand_model_from_filename(stem)
    s = stem.strip()
    s = re.sub(r"^\d+(?:\.\d+){1,3}\.?\s+", "", s)
    if brand:
        # remove brand and everything after
        idx = s.find(brand)
        if idx > 0:
            s = s[:idx].strip()
    # Strip trailing "-N" if filename was a parent rack
    s = re.sub(r"\s*-\s*\d+\s*$", "", s).strip()
    return s


def detect_row_role(row_num: int) -> dict:
    """Determine whether this row is a section header, an item (N)), or a
    sub-item (N.). Returns the structural location used by Col D format
    generation."""
    row = next((r for r in ROWS if r["row"] == row_num), None)
    if not row:
        return {"role": "unknown"}
    b_full = row.get("B") or ""
    b = b_full.strip()
    section = row.get("section_inferred", "")

    # Section header: B starts with full section number + period (3+ levels)
    if re.match(r"^\d+(?:\.\d+){2,3}\.\s+", b):
        return {"role": "section_header", "section": section}

    # Item: B starts with "N)" (after some indent)
    m_item = re.match(r"^(\d+)\s*\)\s*", b)
    # Sub-item: B starts with "N." (after deeper indent)
    m_sub = re.match(r"^(\d+)\s*\.\s*", b)

    # Distinguish item vs sub-item by indent depth in original Col B
    indent_a = len(b_full) - len(b_full.lstrip())
    if m_item and (not m_sub or m_item.end() <= m_sub.end()):
        return {"role": "item", "item_num": int(m_item.group(1)),
                "section": section, "indent": indent_a}
    if m_sub:
        # Find the most recent preceding item-row in the same section
        parent_item = None
        for prev in ROWS:
            if prev["row"] >= row_num:
                continue
            if prev.get("section_inferred") != section:
                continue
            pb = (prev.get("B") or "").strip()
            pm = re.match(r"^(\d+)\s*\)\s*", pb)
            if pm:
                parent_item = int(pm.group(1))
        return {"role": "sub_item", "sub_num": int(m_sub.group(1)),
                "parent_item": parent_item, "section": section, "indent": indent_a}

    return {"role": "unknown", "section": section}


def col_b_search_tokens(b: str) -> list[str]:
    """Distinctive tokens from Col B that should also appear in catalog text.
    Prefers English/code/numeric+unit because Thai compound matching is
    fragile across catalog encodings."""
    s = re.sub(r"^\s*\d+\s*[\)\.]\s*", "", b or "")
    s = re.sub(r"^\s*\d+(?:\.\d+)*\.\s+", "", s)
    seen: set[str] = set()
    out: list[str] = []

    def add(t: str):
        t = t.strip()
        if len(t) < 3 or t.lower() in seen:
            return
        seen.add(t.lower())
        out.append(t)

    for m in re.finditer(r"[A-Za-z][A-Za-z0-9]{2,}(?:[-+/.][A-Za-z0-9]+)*", s):
        add(m.group(0))
    for m in re.finditer(r"(?:IEC|ISO|ASTM|ANSI|UL|TIS|มอก\.?)\s*\d+(?:[-/.\s]\d+)*", s):
        add(m.group(0))
    for m in re.finditer(r"\b(?:IP\d+|DC\s*\d+(?:-\d+)?\s*V|AC\s*\d+\s*V)\b", s):
        add(m.group(0))
    for m in re.finditer(
        r"\d+(?:[\.,]\d+)?\s*(?:mA|MHz|GHz|kHz|Hz|°C|°F|mm|cm|m|kg|hr|min|kV|V|A|W|kW|Mbps|Gbps|MB|GB|TB|byte)",
        s, re.IGNORECASE,
    ):
        add(m.group(0))
    return out


def find_text_match_in_pdf(pdf_path: Path, b_text: str,
                            min_score: int = 2) -> dict | None:
    """Find best-matching page for `b_text` in `pdf_path`. Returns
    {page, rects, score, matched_tokens} or None."""
    tokens = col_b_search_tokens(b_text)
    if not tokens or not pdf_path.exists():
        return None
    try:
        doc = fitz.open(pdf_path)
    except Exception:
        return None
    best = None
    for pno in range(len(doc)):
        page = doc[pno]
        score = 0
        rects: list = []
        matched: list[str] = []
        for tok in tokens:
            for variant in _thai_variants(tok):
                try:
                    rs = page.search_for(variant, quads=False)
                except Exception:
                    rs = []
                if rs:
                    score += 1
                    matched.append(tok)
                    rects.extend(rs[:2])
                    break
        if score and (best is None or score > best["score"]):
            best = {"page": pno + 1, "score": score,
                    "matched_tokens": matched,
                    "rects": [[r.x0, r.y0, r.x1, r.y1] for r in rects]}
    if best and best["score"] >= max(min_score, len(tokens) // 4):
        return best
    return None


def merge_rects_to_content(rects: list, padding: float = 4.0,
                           max_height: float = 80.0) -> list:
    """Combine matched-token rects into a single content rect. If the rects
    span too tall a region (>max_height), use only the topmost cluster to
    avoid over-broad highlights."""
    if not rects:
        return [0, 0, 0, 0]
    sorted_by_y = sorted(rects, key=lambda r: r[1])
    cluster = [sorted_by_y[0]]
    for r in sorted_by_y[1:]:
        if r[1] - cluster[-1][3] < 25:  # within 25pt of previous → same cluster
            cluster.append(r)
        else:
            break
    x0 = min(r[0] for r in cluster) - padding
    y0 = min(r[1] for r in cluster) - padding
    x1 = max(r[2] for r in cluster) + padding
    y1 = max(r[3] for r in cluster) + padding
    return [max(0, x0), max(0, y0), x1, min(y1, y0 + max_height)]


def compute_label_rect(content_rect: list, page_w: float, page_h: float,
                       label_w: float = 130, label_h: float = 14) -> list:
    """Place a FreeText label in whitespace next to the content rect.
    Prefers right side; falls back to below if right doesn't fit."""
    x0, y0, x1, y1 = content_rect
    # Right of rect
    rx0 = x1 + 5
    if rx0 + label_w <= page_w - 5:
        ry0 = max(0, (y0 + y1) / 2 - label_h / 2)
        return [rx0, ry0, rx0 + label_w, ry0 + label_h]
    # Below rect
    by0 = y1 + 4
    if by0 + label_h <= page_h - 5:
        bx0 = x0
        return [bx0, by0, bx0 + label_w, by0 + label_h]
    # Above rect
    ay1 = y0 - 4
    return [x0, ay1 - label_h, x0 + label_w, ay1]


def make_col_d_for_row(row: dict, role_info: dict, pdf_path: Path,
                       page: int | None) -> str:
    """Compose Col D string per SKILL.md convention."""
    role = role_info.get("role")
    section = role_info.get("section") or ""
    stem = pdf_path.stem
    cat_full = stem  # full filename (no .pdf)

    # Determine the "ref" portion: prefer "{parent}-{N}" form when filename
    # ends in -N, else use the full filename (newer convention from
    # clone_*.py scripts).
    ref = cat_full
    sec_name = ""
    parent_folder = pdf_path.parent.name
    fm = re.match(r"^(\d+(?:\.\d+){1,3})\.?\s*-(\d+)\b", parent_folder)
    if fm:
        ref = f"{fm.group(1)}-{fm.group(2)}"
        sec_name = extract_section_name(stem, section)

    if role == "section_header":
        brand, model = parse_brand_model_from_filename(stem)
        if brand and model:
            return f"ยี่ห้อ {brand} รุ่น {model}"
        if brand:
            return f"ยี่ห้อ {brand}"
        return ""

    if role == "item":
        n = role_info.get("item_num", 1)
        if page is None:
            return ""
        if sec_name:
            return f"เทียบเท่าข้อกำหนด เอกสาร {ref} {sec_name} หน้า {page} ข้อ {section} ข้อ {n})"
        return f"เทียบเท่าข้อกำหนด เอกสาร {ref} หน้า {page} ข้อ {section} ข้อ {n})"

    if role == "sub_item":
        sub = role_info.get("sub_num", 1)
        parent = role_info.get("parent_item")
        if page is None:
            return ""
        if parent is not None:
            tail = f"ข้อ {section} ข้อ {parent}) ข้อย่อย {sub}."
        else:
            tail = f"ข้อ {section} ข้อย่อย {sub}."
        if sec_name:
            return f"เทียบเท่าข้อกำหนด เอกสาร {ref} {sec_name} หน้า {page} {tail}"
        return f"เทียบเท่าข้อกำหนด เอกสาร {ref} หน้า {page} {tail}"

    return ""


def make_col_c_from_b(b: str) -> str:
    """Strip indentation, leading numbering, and compare phrases."""
    if not b:
        return ""
    s = b.strip()
    for needle, repl in _COMPARE_PHRASES_STRIP:
        s = s.replace(needle, repl)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def auto_annotate_plan(row_num: int) -> dict:
    """Generate a dry-run plan for one row. Caller can inspect + apply.

    HITL pipeline:
      1. Apply learned patterns first (highest priority — user-validated).
      2. Fall through to rule-based generation (filename parsing, text
         search, etc.) for the long tail.
      3. (Optional) ask the LLM provider when both above produce low
         confidence.
      4. Compute a confidence score + provenance trail for the UI.
    """
    row = next((r for r in ROWS if r["row"] == row_num), None)
    if row is None:
        return {"ok": False, "error": f"row {row_num} not found"}

    plan: dict = {
        "ok": True,
        "row": row_num,
        "section": row.get("section_inferred"),
        "current_d": row.get("D") or "",
        "proposed_d": "",
        "current_c": row.get("C") or "",
        "proposed_c": "",
        "annotations": [],
        "warnings": [],
        "generator": "rules",
        "provenance": {},
        "confidence": 0.0,
    }

    pdf_rel = row.get("pdf_rel")
    if not pdf_rel:
        plan["ok"] = False
        plan["warnings"].append("ไม่มี PDF ที่อ้างอิง — ไม่สามารถ auto-annotate ได้")
        return plan
    pdf_path = OUTPUT / pdf_rel
    plan["pdf_rel"] = pdf_rel
    plan["pdf_filename"] = pdf_path.name

    role_info = detect_row_role(row_num)
    plan["role"] = role_info

    b = row.get("B") or ""
    proposed_c = make_col_c_from_b(b)
    plan["proposed_c"] = proposed_c

    role = role_info.get("role")

    if role == "section_header":
        # 1. learned filename_brand pattern wins if it fires
        learned_brand, prov = learning.apply_learned_brand(pdf_path.stem)
        if learned_brand:
            # use the rule-based model parser to get the model part, but
            # override the brand with what the user historically picked
            _, model = parse_brand_model_from_filename(pdf_path.stem)
            proposed_d = (f"ยี่ห้อ {learned_brand} รุ่น {model}".strip()
                          if model else f"ยี่ห้อ {learned_brand}")
            plan["proposed_d"] = proposed_d
            plan["generator"] = "rules+pattern"
            plan["provenance"] = {"brand": prov}
        else:
            proposed_d = make_col_d_for_row(row, role_info, pdf_path, None)
            plan["proposed_d"] = proposed_d
        plan["meta"] = {"brand_model_only": True}
        if not plan["proposed_d"]:
            plan["warnings"].append("ไม่สามารถ parse brand/model จาก filename")
        plan["confidence"] = learning.confidence_score(
            generator=plan["generator"], provenance=plan["provenance"],
            role="section_header", has_match=bool(plan["proposed_d"]),
            warnings=len(plan["warnings"]),
        )
        # Claude refinement for section headers (brand/model decomposition)
        plan = _maybe_refine_with_claude(plan, row, role_info)
        return plan

    if role in ("item", "sub_item"):
        match = find_text_match_in_pdf(pdf_path, b)
        if not match:
            plan["warnings"].append("ไม่เจอข้อความ Col B ใน catalog (อาจเป็น commitment)")
            plan["proposed_d"] = "ยินดีปฏิบัติตามข้อกำหนด"
            plan["fallback"] = "commitment"
            plan["confidence"] = learning.confidence_score(
                generator=plan["generator"], provenance=plan["provenance"],
                role=role, has_match=False, warnings=len(plan["warnings"]),
            )
            # Claude refinement — text search missed but Claude may know better
            plan = _maybe_refine_with_claude(plan, row, role_info)
            return plan
        page = match["page"]
        plan["match"] = match
        plan["proposed_d"] = make_col_d_for_row(row, role_info, pdf_path, page)

        content_rect = merge_rects_to_content(match["rects"])
        try:
            doc = fitz.open(pdf_path)
            pg = doc[page - 1]
            page_w, page_h = pg.rect.width, pg.rect.height
        except Exception:
            page_w, page_h = 595, 842
        label_rect = compute_label_rect(content_rect, page_w, page_h)
        if role == "sub_item" and role_info.get("parent_item") is not None:
            label = f"{role_info['section']} ข้อ {role_info['parent_item']}) ข้อย่อย {role_info['sub_num']}."
        elif role == "sub_item":
            label = f"{role_info['section']} ข้อย่อย {role_info['sub_num']}."
        else:
            label = f"{role_info['section']} ข้อ {role_info['item_num']})"

        plan["annotations"] = [
            {"page": page, "type": "Square",
             "rect": [round(v, 2) for v in content_rect], "contents": ""},
            {"page": page, "type": "FreeText",
             "rect": [round(v, 2) for v in label_rect], "contents": label},
        ]
        plan["confidence"] = learning.confidence_score(
            generator=plan["generator"], provenance=plan["provenance"],
            role=role, has_match=True, warnings=len(plan["warnings"]),
        )
        # Claude refinement (if installed)
        plan = _maybe_refine_with_claude(plan, row, role_info)
        return plan

    plan["warnings"].append(f"row role '{role}' — ไม่รองรับการ auto-annotate")
    plan["ok"] = False
    plan["confidence"] = 0.0
    return plan


def _maybe_refine_with_claude(plan: dict, row: dict, role_info: dict) -> dict:
    """If a Claude provider is installed AND the rule-based plan has low
    confidence, ask Claude for a structured proposal and merge into the plan.

    Strategy:
      • plan.confidence ≥ 0.85 → keep rules (no spend)
      • plan.confidence < 0.85 → ask Claude, override col_d/col_c if Claude
        returns a propose_col_d tool call with higher confidence
      • Always record the Claude call's audit for analytics

    Failure modes are non-fatal — Claude errors fall through to rule output.
    """
    try:
        from app import anthropic_provider as ap
    except Exception:
        return plan

    provider = ap.get_provider()
    if provider is None:
        return plan

    if plan.get("confidence", 0) >= 0.85:
        # rules already confident — skip the spend
        plan["llm"] = {"skipped": "high_confidence_rules"}
        return plan

    # Find a few similar past corrections for in-context learning
    few_shot: list[dict] = []
    try:
        few_shot = _claude_few_shot_for(row, role_info, limit=4)
    except Exception as e:
        sys.stderr.write(f"[claude-refine] few_shot failed: {e}\n")

    row_context = {
        "row": row.get("row"),
        "section": row.get("section_inferred"),
        "role": role_info.get("role"),
        "col_a": row.get("A"),
        "col_b": row.get("B") or "",
        "col_c_current": row.get("C") or "",
        "col_d_current": row.get("D") or "",
        "col_e": row.get("E") or "",
        "pdf_rel": row.get("pdf_rel"),
        "pdf_filename": Path(row["pdf_rel"]).name if row.get("pdf_rel") else None,
        "tor_excerpt": _claude_tor_excerpt(row),
        "rule_proposal": {
            "proposed_d": plan.get("proposed_d", ""),
            "generator": plan.get("generator", "rules"),
            "confidence": plan.get("confidence", 0),
        },
        "_few_shot": few_shot,
    }

    try:
        result = provider.propose(row_context=row_context, few_shot=few_shot)
    except ap.BudgetExceededError as e:
        plan["llm"] = {"error": "budget_exceeded", "msg": str(e)}
        plan["warnings"].append(f"Claude skipped: {e}")
        return plan
    except Exception as e:
        plan["llm"] = {"error": str(e)}
        sys.stderr.write(f"[claude-refine] {e}\n")
        return plan

    if not result.get("ok"):
        plan["llm"] = {"error": result.get("error", "unknown")}
        return plan

    # Merge Claude's proposal into plan
    plan["llm"] = {
        "model": result.get("model"),
        "tokens": result.get("usage"),
        "cost_usd": round(result.get("cost_usd", 0), 6),
        "elapsed_ms": result.get("elapsed_ms"),
        "tool_calls": result.get("tool_calls", []),
        "rationale": "",
        "claude_confidence": 0.0,
        "escalation": None,
    }

    for tc in result.get("tool_calls", []):
        name = tc.get("name")
        inp = tc.get("input", {}) or {}
        if name == "propose_col_d":
            claude_d = (inp.get("col_d_text") or "").strip()
            claude_c = (inp.get("col_c_proposed") or "").strip()
            claude_conf = float(inp.get("confidence", 0))
            plan["llm"]["rationale"] = inp.get("rationale", "")
            plan["llm"]["claude_confidence"] = claude_conf
            plan["llm"]["pattern"] = inp.get("pattern", "")
            # Override only if Claude's confidence beats the rule
            if claude_d and claude_conf > plan.get("confidence", 0):
                plan["proposed_d"] = claude_d
                plan["generator"] = f"claude+{plan.get('generator', 'rules')}"
                plan["confidence"] = max(plan.get("confidence", 0), claude_conf)
                plan["provenance"] = {
                    **plan.get("provenance", {}),
                    "claude_pattern": inp.get("pattern", ""),
                    "claude_rationale": inp.get("rationale", ""),
                }
                if claude_c:
                    plan["proposed_c"] = claude_c
                if inp.get("page_in_catalog"):
                    plan["claude_page"] = int(inp["page_in_catalog"])
        elif name == "propose_brand_model":
            brand = (inp.get("brand") or "").strip()
            model = (inp.get("model") or "").strip()
            if brand and brand != "-" and model:
                # For section_header rows, override Col D with Claude's brand+model
                claude_d = f"ยี่ห้อ {brand} รุ่น {model}"
                if float(inp.get("confidence", 0)) > plan.get("confidence", 0):
                    plan["proposed_d"] = claude_d
                    plan["generator"] = "claude+brand_model"
                    plan["confidence"] = float(inp.get("confidence", 0))
                    plan["provenance"]["claude_brand"] = brand
                    plan["provenance"]["claude_model"] = model
            elif brand == "-" and model:
                plan["proposed_d"] = f"ยี่ห้อ - รุ่น {model}"
                plan["generator"] = "claude+brand_model_fabricate"
                plan["confidence"] = float(inp.get("confidence", 0))
            plan["llm"]["rationale"] = inp.get("rationale", "")
            plan["llm"]["claude_confidence"] = float(inp.get("confidence", 0))
        elif name == "escalate_to_user":
            plan["llm"]["escalation"] = {
                "question": inp.get("question", ""),
                "options": inp.get("options", []),
                "context": inp.get("context", ""),
            }
            plan["warnings"].append(
                f"Claude escalated: {inp.get('question', '')[:100]}"
            )

    return plan


def _claude_few_shot_for(row: dict, role_info: dict, limit: int = 4) -> list[dict]:
    """Pull the most-similar past corrections from learning_feedback to use
    as in-context examples. Heuristic: same section root + same role +
    user_action='edited' (i.e. the user actually fixed something useful)."""
    section = row.get("section_inferred") or ""
    sec_root = ".".join(section.split(".")[:2]) if section else ""
    role = role_info.get("role") or ""
    out: list[dict] = []
    try:
        with db.conn() as c:
            rows = c.execute(
                """SELECT input_b, final_d, generator, correction_kind
                   FROM learning_feedback
                   WHERE user_action = 'edited'
                     AND section LIKE ?
                     AND input_role = ?
                   ORDER BY ts DESC LIMIT ?""",
                (sec_root + ".%", role, limit),
            ).fetchall()
            for r in rows:
                out.append({
                    "input_b": r["input_b"],
                    "final_d": r["final_d"],
                    "generator": r["generator"],
                    "correction_kind": r["correction_kind"],
                })
    except Exception as e:
        sys.stderr.write(f"[claude-fewshot] {e}\n")
    return out


def _claude_tor_excerpt(row: dict, max_chars: int = 800) -> str:
    """Pull the TOR text snippet for the row's section, if available."""
    try:
        section = row.get("section_inferred") or ""
        if not section:
            return ""
        with db.conn() as c:
            r = c.execute(
                """SELECT page_text FROM tor_pages
                   JOIN tor_sections ON tor_sections.page_num = tor_pages.page_num
                   WHERE tor_sections.section = ? LIMIT 1""",
                (section,),
            ).fetchone()
            if r and r["page_text"]:
                return r["page_text"][:max_chars]
    except Exception:
        pass
    return ""


def apply_auto_annotate_plan(plan: dict, write_pdf: bool = True,
                              write_xlsx: bool = True) -> dict:
    """Commit a plan: snapshot first, then write Col C/D in xlsx and add
    annotations to the catalog PDF."""
    if not plan.get("ok"):
        return {"ok": False, "error": "plan is not OK"}

    result = {"ok": True, "row": plan.get("row"),
              "xlsx_written": False, "pdf_written": False, "snapshot": None}

    # Take a quick snap first (xlsx only — fast)
    snap = _run_version_cmd(["snap", f"pre-auto-row-{plan.get('row')}"], timeout=60)
    if snap.get("ok"):
        result["snapshot"] = "pre-snap created"

    # Write xlsx
    if write_xlsx and plan.get("proposed_d"):
        try:
            wb = openpyxl.load_workbook(XLSX_PATH)
            ws = wb.active
            r = plan["row"]
            if plan.get("proposed_d"):
                ws.cell(r, 4).value = plan["proposed_d"]
            if plan.get("proposed_c") and not (ws.cell(r, 3).value or "").strip():
                ws.cell(r, 3).value = plan["proposed_c"]
            # Save with O_TRUNC workaround for Google Drive
            tmp = XLSX_PATH.with_suffix(".xlsx.tmp")
            wb.save(str(tmp))
            with open(tmp, "rb") as f:
                data = f.read()
            fd = os.open(str(XLSX_PATH), os.O_WRONLY | os.O_TRUNC)
            try:
                os.write(fd, data)
            finally:
                os.close(fd)
            tmp.unlink(missing_ok=True)
            result["xlsx_written"] = True
        except Exception as e:
            result["ok"] = False
            result["xlsx_error"] = str(e)

    # Write PDF annotations
    if write_pdf and plan.get("annotations") and plan.get("pdf_rel"):
        try:
            pdf_path = OUTPUT / plan["pdf_rel"]
            edits = []
            for a in plan["annotations"]:
                edits.append({
                    "action": "create",
                    "page": a["page"],
                    "type": a["type"],
                    "rect": a["rect"],
                    "contents": a["contents"],
                    "fontsize": 9,
                })
            r = apply_pdf_edits(pdf_path, edits)
            result["pdf_written"] = (r.get("applied", 0) > 0)
            result["pdf_result"] = r
        except Exception as e:
            result["ok"] = False
            result["pdf_error"] = str(e)

    return result


@app.route("/api/auto_annotate/preview")
def api_auto_annotate_preview():
    row_num = int(request.args.get("row", 0))
    if not row_num:
        return jsonify({"ok": False, "error": "row required"}), 400
    return jsonify(auto_annotate_plan(row_num))


@app.route("/api/auto_annotate/apply", methods=["POST"])
def api_auto_annotate_apply():
    data = request.get_json(silent=True) or {}
    row_num = int(data.get("row", 0))
    if not row_num:
        return jsonify({"ok": False, "error": "row required"}), 400
    plan = auto_annotate_plan(row_num)
    if not plan.get("ok"):
        return jsonify({"ok": False, "error": "no plan", "plan": plan}), 400
    write_pdf = bool(data.get("write_pdf", True))
    write_xlsx = bool(data.get("write_xlsx", True))
    # Record the plan first so even failed applies leave a trace
    try:
        plan_id = db.record_plan(row_num, plan)
    except Exception:
        plan_id = None
    result = apply_auto_annotate_plan(plan, write_pdf=write_pdf, write_xlsx=write_xlsx)
    if result.get("ok"):
        load_rows()
        TOR_CACHE.clear()
        sync_db_from_memory()
        if plan_id:
            try: db.mark_plan_applied(plan_id, result)
            except Exception: pass
        # HITL: record feedback. The user can override the suggested Col D
        # by passing `final_d` in the body; otherwise we treat the apply
        # as "accepted as proposed".
        try:
            final_d = data.get("final_d") if "final_d" in data else plan.get("proposed_d")
            final_c = data.get("final_c") if "final_c" in data else plan.get("proposed_c")
            user_action = data.get("user_action") or (
                "edited" if final_d != plan.get("proposed_d") else "accepted"
            )
            learning.record_feedback(
                row_num=row_num,
                section=plan.get("section"),
                input_b=(next((r["B"] for r in ROWS if r["row"] == row_num), "") or ""),
                input_pdf_rel=plan.get("pdf_rel"),
                input_role=(plan.get("role") or {}).get("role") or "",
                input_filename=plan.get("pdf_filename"),
                suggested_c=plan.get("proposed_c") or "",
                suggested_d=plan.get("proposed_d") or "",
                suggested_annots=plan.get("annotations") or [],
                confidence=plan.get("confidence", 0),
                generator=plan.get("generator", "rules"),
                provenance=plan.get("provenance"),
                user_action=user_action,
                final_c=final_c or "",
                final_d=final_d or "",
                final_annots=plan.get("annotations") or [],
            )
        except Exception as e:
            sys.stderr.write(f"[learning] record_feedback: {e}\n")
        try:
            db.log_audit(action="auto_annotate_apply",
                         target_type="row", target_id=str(row_num),
                         before=plan.get("current_d") or "",
                         after=plan.get("proposed_d") or "",
                         details={"role": (plan.get("role") or {}).get("role"),
                                  "annotations": len(plan.get("annotations") or []),
                                  "write_pdf": write_pdf,
                                  "write_xlsx": write_xlsx,
                                  "plan_id": plan_id,
                                  "confidence": plan.get("confidence")},
                         actor="user")
        except Exception: pass
    return jsonify({"ok": result.get("ok"), "result": result, "plan": plan})


@app.route("/api/auto_annotate/batch_preview")
def api_auto_annotate_batch_preview():
    """Return plans for all rows that look auto-annotateable (have a
    catalog PDF + needs Col D fill)."""
    plans = []
    for r in ROWS:
        if not r.get("needs_col_d") and (r.get("D") or "").strip():
            continue
        if not r.get("pdf_rel"):
            continue
        p = auto_annotate_plan(r["row"])
        plans.append(p)
    return jsonify({
        "ok": True,
        "count": len(plans),
        "by_role": {role: sum(1 for p in plans if (p.get("role") or {}).get("role") == role)
                    for role in ("section_header", "item", "sub_item", "unknown")},
        "plans": plans,
    })


@app.route("/api/auto_annotate/batch_apply", methods=["POST"])
def api_auto_annotate_batch_apply():
    data = request.get_json(silent=True) or {}
    rows = data.get("rows") or []
    write_pdf = bool(data.get("write_pdf", True))
    write_xlsx = bool(data.get("write_xlsx", True))
    if not isinstance(rows, list) or not rows:
        return jsonify({"ok": False, "error": "rows[] required"}), 400

    # One snapshot at start of the batch
    _run_version_cmd(["snap", f"pre-batch-auto-{len(rows)}-rows"], timeout=60)

    results = []
    for rn in rows:
        try:
            plan = auto_annotate_plan(int(rn))
        except Exception as e:
            results.append({"row": rn, "ok": False, "error": str(e)})
            continue
        if not plan.get("ok"):
            results.append({"row": rn, "ok": False, "warnings": plan.get("warnings", [])})
            continue
        r = apply_auto_annotate_plan(plan, write_pdf=write_pdf, write_xlsx=write_xlsx)
        results.append({"row": rn, "ok": r.get("ok"),
                        "proposed_d": plan.get("proposed_d"),
                        "annotations_added": len(plan.get("annotations") or [])
                                              if r.get("pdf_written") else 0})
    # refresh once at end
    load_rows()
    TOR_CACHE.clear()
    return jsonify({"ok": True, "count": len(results),
                    "applied_ok": sum(1 for x in results if x.get("ok")),
                    "results": results})


# ---------------------------------------------------------------------------
# Manual-annotate workflow loop
#
# When the AI couldn't find Col B content in the catalog and fell back to
# "ยินดีปฏิบัติตามข้อกำหนด", but the user notices the spec IS in the catalog,
# they enter manual-annotate mode:
#   1. /api/manual_annotate/context returns the row's role + a suggested
#      FreeText label content (using SKILL.md format).
#   2. User draws a rect on the catalog PDF (via existing edit-mode SVG
#      overlay). The frontend auto-pairs a draggable FreeText label.
#   3. /api/manual_annotate/save accepts the rect + label positions, writes
#      both annotations into the PDF, computes the proper Col D string from
#      the page where the rect sits, updates the xlsx, and records the
#      manual override as strong learning feedback.
# ---------------------------------------------------------------------------

def _label_for_row(role_info: dict) -> str:
    """SKILL.md label for a Square's FreeText partner."""
    role = role_info.get("role")
    section = role_info.get("section") or ""
    if role == "sub_item":
        if role_info.get("parent_item") is not None:
            return (f"{section} ข้อ {role_info['parent_item']}) "
                    f"ข้อย่อย {role_info['sub_num']}.")
        return f"{section} ข้อย่อย {role_info['sub_num']}."
    if role == "item":
        return f"{section} ข้อ {role_info['item_num']})"
    return section


def _candidate_pdfs_for_section(section: str | None) -> list[Path]:
    """Find catalog PDFs that could correspond to a section. Used by the
    manual-annotate flow to populate a picker for commitment rows whose
    pdf_rel wasn't resolved.

    Wildcard matches are sorted by trailing numeric suffix so the parent
    rack catalog (-1) is picked first, not whichever sub-catalog Python
    happened to insert first in the dict.
    """
    if not section:
        return []
    out: list[Path] = []
    seen: set[Path] = set()

    def _add(paths):
        for p in paths:
            if p not in seen:
                seen.add(p); out.append(p)

    def _sort_key(k: str) -> tuple:
        # Sort "5.1.1-1", "5.1.1-2", ... → 1, 2, ... so rack parent (-1) wins
        m = re.search(r"-(\d+)$", k)
        return (int(m.group(1)) if m else 0, k)

    # 1. Direct keys + dot/dash translation (highest priority)
    for k in _ref_to_folder_keys(section):
        if k in PDF_INDEX: _add(PDF_INDEX[k])
        if k in SECTION_INDEX: _add(SECTION_INDEX[k])

    # 2. Wildcard: any sub-catalog under this section, sorted ascending so the
    #    parent rack (-1) comes first (e.g. for 5.1.1 → 5.1.1-1 before -2..-7)
    matches = sorted(
        [k for k in PDF_INDEX if k.startswith(section + "-") or k == section],
        key=_sort_key,
    )
    for k in matches:
        _add(PDF_INDEX[k])

    # 3. Parent section fallback (e.g. 5.1.6.1 → 5.1.6) so a row referring to
    #    a whole sub-section can find its parent rack catalog.
    parts = section.split(".")
    if len(parts) >= 3:
        parent = ".".join(parts[:-1])
        for k in _ref_to_folder_keys(parent):
            if k in PDF_INDEX: _add(PDF_INDEX[k])
            if k in SECTION_INDEX: _add(SECTION_INDEX[k])
        parent_matches = sorted(
            [k for k in PDF_INDEX if k.startswith(parent + "-")],
            key=_sort_key,
        )
        for k in parent_matches:
            _add(PDF_INDEX[k])
    return out


@app.route("/api/manual_annotate/context")
def api_manual_context():
    row_num = int(request.args.get("row", 0))
    row = next((r for r in ROWS if r["row"] == row_num), None)
    if not row:
        return jsonify({"ok": False, "error": "row not found"}), 404

    role_info = detect_row_role(row_num)
    pdf_rel = row.get("pdf_rel")
    section = row.get("section_inferred")

    # If row didn't auto-resolve a PDF (e.g. commitment row), surface a
    # picker of candidate catalogs from the row's section so the user can
    # choose where to mark.
    candidates: list[dict] = []
    if not pdf_rel and section:
        for p in _candidate_pdfs_for_section(section):
            candidates.append({
                "rel": str(p.relative_to(OUTPUT)),
                "name": p.name,
                "folder": p.parent.name,
            })
        if candidates:
            pdf_rel = candidates[0]["rel"]

    def _meta(rel: str) -> dict | None:
        try:
            doc = fitz.open(OUTPUT / rel)
            return {
                "rel": rel,
                "pages": len(doc),
                "page_sizes": [(round(doc[i].rect.width, 1),
                                round(doc[i].rect.height, 1))
                               for i in range(len(doc))],
            }
        except Exception:
            return None

    return jsonify({
        "ok": True,
        "row": row_num,
        "section": section,
        "role": role_info,
        "col_b": row.get("B") or "",
        "col_d_current": row.get("D") or "",
        "is_commitment": (row.get("D") or "").strip().startswith("ยินดีปฏิบัติ"),
        "pdf_rel": pdf_rel,
        "pdf_meta": _meta(pdf_rel) if pdf_rel else None,
        "candidates": candidates,
        "suggested_label": _label_for_row(role_info),
    })


@app.route("/api/manual_annotate/save", methods=["POST"])
def api_manual_save():
    """Body: {row, page, content_rect:[x0,y0,x1,y1], label_rect:[...],
             label_text, pdf_rel?}"""
    data = request.get_json(silent=True) or {}
    try:
        row_num = int(data["row"])
        page = int(data["page"])
        content_rect = list(map(float, data["content_rect"]))
        label_rect = list(map(float, data["label_rect"]))
    except Exception as e:
        return jsonify({"ok": False, "error": f"bad payload: {e}"}), 400

    row = next((r for r in ROWS if r["row"] == row_num), None)
    if not row:
        return jsonify({"ok": False, "error": "row not found"}), 404

    pdf_rel = data.get("pdf_rel") or row.get("pdf_rel")
    if not pdf_rel:
        return jsonify({"ok": False, "error": "no PDF for row"}), 404
    pdf_path = (OUTPUT / pdf_rel).resolve()
    if not str(pdf_path).startswith(str(OUTPUT.resolve())) or not pdf_path.exists():
        return jsonify({"ok": False, "error": "PDF not found"}), 404

    role_info = detect_row_role(row_num)
    label_text = (data.get("label_text") or _label_for_row(role_info)).strip()

    # Pre-snapshot
    _run_version_cmd(["snap", f"pre-manual-row-{row_num}"], timeout=60)

    # Apply PDF edits — Square + FreeText pair. fontsize is left to
    # apply_pdf_edits which clamps from rect height (matches SVG overlay).
    edits = [
        {"action": "create", "page": page, "type": "Square",
         "rect": content_rect, "contents": ""},
        {"action": "create", "page": page, "type": "FreeText",
         "rect": label_rect, "contents": label_text},
    ]
    pdf_result = apply_pdf_edits(pdf_path, edits)

    # Compute new Col D using the same generator as auto-annotate so format
    # stays consistent with the rest of the spreadsheet.
    new_d = make_col_d_for_row(row, role_info, pdf_path, page)
    if not new_d:
        # Fallback (should rarely hit since make_col_d_for_row covers most)
        new_d = (f"เทียบเท่าข้อกำหนด เอกสาร {pdf_path.stem} หน้า {page} "
                 f"ข้อ {role_info.get('section', '')}")

    old_d = row.get("D") or ""

    # Update xlsx Col D (Google Drive O_TRUNC workaround)
    try:
        wb = openpyxl.load_workbook(XLSX_PATH)
        ws = wb.active
        ws.cell(row_num, 4).value = new_d
        # If Col C is still the raw Col B, tighten it via make_col_c_from_b
        cur_c = (ws.cell(row_num, 3).value or "").strip()
        cur_b = (ws.cell(row_num, 2).value or "").strip()
        if not cur_c or cur_c == cur_b:
            ws.cell(row_num, 3).value = make_col_c_from_b(cur_b)
        tmp = XLSX_PATH.with_suffix(".xlsx.tmp")
        wb.save(str(tmp))
        with open(tmp, "rb") as f:
            data_bytes = f.read()
        fd = os.open(str(XLSX_PATH), os.O_WRONLY | os.O_TRUNC)
        try:
            os.write(fd, data_bytes)
        finally:
            os.close(fd)
        tmp.unlink(missing_ok=True)
    except Exception as e:
        return jsonify({"ok": False, "error": f"xlsx write failed: {e}",
                        "pdf_result": pdf_result}), 500

    # Refresh internal caches
    try:
        load_rows()
        TOR_CACHE.clear()
        sync_db_from_memory()
    except Exception as e:
        sys.stderr.write(f"[manual_save] refresh: {e}\n")

    # Audit + learning feedback (strong "edited" signal — manual override)
    try:
        db.log_audit(action="manual_annotate",
                     target_type="row", target_id=str(row_num),
                     before=old_d, after=new_d,
                     details={"page": page, "pdf_rel": pdf_rel,
                              "annots_applied": pdf_result.get("applied", 0),
                              "annots_errors": pdf_result.get("errors", 0)},
                     actor="user")
    except Exception: pass
    try:
        learning.record_feedback(
            row_num=row_num,
            section=row.get("section_inferred"),
            input_b=row.get("B", "") or "",
            input_pdf_rel=pdf_rel,
            input_role=role_info.get("role", ""),
            input_filename=pdf_path.name,
            suggested_c=row.get("C", "") or "",
            suggested_d=old_d,
            suggested_annots=[],
            confidence=0.0,
            generator="commitment_fallback",
            provenance={},
            user_action="edited",
            final_c=row.get("C", "") or "",
            final_d=new_d,
            final_annots=edits,
        )
    except Exception as e:
        sys.stderr.write(f"[manual_save] feedback: {e}\n")

    return jsonify({"ok": True, "row": row_num, "page": page,
                    "old_d": old_d, "new_d": new_d,
                    "label_text": label_text,
                    "pdf_result": pdf_result})


# ---------------------------------------------------------------------------
# Re-annotate Wizard
#
# Extends the manual-annotate flow into a stepwise wizard that:
#   • for brand_model section_header rows, expects 2 rect+label pairs
#     (ยี่ห้อ, รุ่น) per SKILL.md §"Label DA / DS standard (ยี่ห้อ/รุ่น)"
#   • for item / sub_item rows, expects 1 pair (existing behaviour)
#   • lets the user switch the catalog PDF before annotating
#   • optionally deletes existing annotations whose label starts with the
#     same prefix(es) before writing the new ones — a clean re-annotate
# ---------------------------------------------------------------------------

def _expected_steps_for_row(row: dict, role_info: dict) -> list[dict]:
    """Plan the step labels the user must produce. Pairs the SKILL.md
    label convention with the row's structural role."""
    role = role_info.get("role")
    section = role_info.get("section") or ""
    parsed = row.get("parsed") or {}

    # Section header that's brand_model → 2 steps (ยี่ห้อ + รุ่น)
    if role == "section_header":
        # parsed.type is the cheapest signal; fallback to filename inspection
        is_bm = (parsed.get("type") == "brand_model")
        if not is_bm:
            pdf_rel = row.get("pdf_rel") or ""
            if pdf_rel:
                stem = Path(pdf_rel).stem
                b, m = parse_brand_model_from_filename(stem)
                is_bm = bool(b and m)
        if is_bm:
            return [
                {"label": "ยี่ห้อ",
                 "hint": "ลาก rect รอบ <strong>โลโก้ยี่ห้อ</strong> ใน catalog",
                 "kind": "brand"},
                {"label": "รุ่น",
                 "hint": "ลาก rect รอบ <strong>ชื่อรุ่น</strong> (model number) ใน catalog",
                 "kind": "model"},
            ]
        return [
            {"label": section,
             "hint": "ลาก rect รอบเนื้อหาที่เกี่ยวกับ section นี้",
             "kind": "section"},
        ]

    # Item / sub_item — same single-step flow as manual-annotate
    label = _label_for_row(role_info)
    if role == "sub_item":
        hint = (f"ลาก rect รอบเนื้อหาของ <strong>ข้อย่อย "
                f"{role_info.get('sub_num', '?')}</strong>")
    elif role == "item":
        hint = (f"ลาก rect รอบเนื้อหาของ <strong>ข้อ "
                f"{role_info.get('item_num', '?')})</strong>")
    else:
        hint = "ลาก rect รอบเนื้อหาที่เกี่ยวกับ row นี้"
    return [{"label": label, "hint": hint, "kind": role or "row"}]


def _delete_label_prefixes_for_row(row: dict, role_info: dict) -> list[str]:
    """Prefixes whose existing FreeText annotations should be wiped before
    we re-annotate. Conservative — only deletes labels that are *clearly*
    this row's prior work."""
    role = role_info.get("role")
    section = role_info.get("section") or ""
    parsed = row.get("parsed") or {}

    if role == "section_header":
        if parsed.get("type") == "brand_model":
            return ["ยี่ห้อ", "รุ่น"]
        if section:
            return [section]   # exact section header label
        return []

    if role == "item":
        n = role_info.get("item_num")
        if section and n is not None:
            # Match "5.1.1 ข้อ 1)" but NOT "5.1.1 ข้อ 1) ข้อย่อย 1." — sub-items
            # share the parent prefix, so we anchor on a closing ')' boundary.
            return [f"{section} ข้อ {n})"]
        return []

    if role == "sub_item":
        sub = role_info.get("sub_num")
        parent = role_info.get("parent_item")
        if section and sub is not None:
            if parent is not None:
                return [f"{section} ข้อ {parent}) ข้อย่อย {sub}."]
            return [f"{section} ข้อย่อย {sub}."]
        return []

    return []


def _annots_to_delete_by_prefix(pdf_path: Path, page_num: int,
                                 prefixes: list[str]) -> list[int]:
    """Return xrefs of FreeText annotations whose contents begin with any
    of the given prefixes, plus the xref of any nearby Square that looks
    paired with them (within 12pt edge-to-edge on any side).

    Conservative: a Square is only paired if it shares a side with the
    label rect — that's how SKILL.md and our generator place them.
    """
    if not prefixes or not pdf_path.exists():
        return []
    try:
        doc = fitz.open(pdf_path)
    except Exception:
        return []
    if page_num < 1 or page_num > len(doc):
        doc.close()
        return []

    page = doc[page_num - 1]
    annots = safe_iter_annots(page)
    label_to_xref: list[tuple[float, float, float, float, int]] = []  # x0,y0,x1,y1,xref
    square_xrefs: list[tuple[float, float, float, float, int]] = []
    for ann in annots:
        try:
            kind = ann.type[1] if isinstance(ann.type, tuple) else str(ann.type)
        except Exception:
            kind = ""
        try:
            r = ann.rect
            xref = ann.xref
            contents = (ann.info.get("content") or "").strip()
        except Exception:
            continue
        if kind == "FreeText":
            for pre in prefixes:
                if contents.startswith(pre):
                    label_to_xref.append((r.x0, r.y0, r.x1, r.y1, int(xref)))
                    break
        elif kind == "Square":
            square_xrefs.append((r.x0, r.y0, r.x1, r.y1, int(xref)))

    out: set[int] = set()
    for (lx0, ly0, lx1, ly1, lxref) in label_to_xref:
        out.add(lxref)
        # Find the closest Square that shares a side within 12pt
        best_xref = None
        best_dist = 1e9
        for (sx0, sy0, sx1, sy1, sxref) in square_xrefs:
            # Right side of Square ↔ Left side of label
            dx = abs(sx1 - lx0)
            dy = abs((sy0 + sy1) / 2 - (ly0 + ly1) / 2)
            d = dx + dy
            if dx <= 12 and dy <= 30 and d < best_dist:
                best_dist = d; best_xref = sxref
            # Below Square ↔ above label
            dy2 = abs(sy1 - ly0)
            dx2 = abs((sx0 + sx1) / 2 - (lx0 + lx1) / 2)
            d2 = dy2 + dx2
            if dy2 <= 12 and dx2 <= 80 and d2 < best_dist:
                best_dist = d2; best_xref = sxref
            # Left side of Square ↔ Right side of label (label on left)
            dx3 = abs(sx0 - lx1)
            d3 = dx3 + dy
            if dx3 <= 12 and dy <= 30 and d3 < best_dist:
                best_dist = d3; best_xref = sxref
            # Above Square ↔ below label (label above)
            dy4 = abs(sy0 - ly1)
            d4 = dy4 + dx2
            if dy4 <= 12 and dx2 <= 80 and d4 < best_dist:
                best_dist = d4; best_xref = sxref
        if best_xref is not None:
            out.add(best_xref)

    doc.close()
    return sorted(out)


@app.route("/api/reannotate/context")
def api_reannotate_context():
    """Returns the wizard plan for a row.

    Response: {ok, row, section, role, col_b, col_d_current, parsed_type,
               pdf_rel, pdf_meta, candidates, steps, delete_label_prefixes,
               brand_hint, model_hint}
    """
    try:
        row_num = int(request.args.get("row", 0))
    except Exception:
        return jsonify({"ok": False, "error": "bad row"}), 400
    row = next((r for r in ROWS if r["row"] == row_num), None)
    if not row:
        return jsonify({"ok": False, "error": "row not found"}), 404

    role_info = detect_row_role(row_num)
    section = row.get("section_inferred")
    pdf_rel = row.get("pdf_rel")

    # Always surface candidates so the user can swap PDFs even when one is
    # already assigned (key request from the user).
    candidates: list[dict] = []
    cand_paths = _candidate_pdfs_for_section(section)
    for p in cand_paths:
        try:
            rel = str(p.relative_to(OUTPUT))
        except Exception:
            continue
        candidates.append({
            "rel": rel,
            "name": p.name,
            "folder": p.parent.name,
            "is_current": (rel == pdf_rel),
        })

    # If the row had no PDF, default to the first candidate
    if not pdf_rel and candidates:
        pdf_rel = candidates[0]["rel"]

    def _meta(rel: str) -> dict | None:
        try:
            doc = fitz.open(OUTPUT / rel)
            return {
                "rel": rel,
                "pages": len(doc),
                "page_sizes": [(round(doc[i].rect.width, 1),
                                round(doc[i].rect.height, 1))
                               for i in range(len(doc))],
            }
        except Exception:
            return None

    # Filename-derived brand/model hints (defaults the wizard pre-fills)
    brand_hint, model_hint = ("", "")
    if pdf_rel:
        stem = Path(pdf_rel).stem
        brand_hint, model_hint = parse_brand_model_from_filename(stem)

    return jsonify({
        "ok": True,
        "row": row_num,
        "section": section,
        "role": role_info,
        "col_b": row.get("B") or "",
        "col_d_current": row.get("D") or "",
        "parsed_type": (row.get("parsed") or {}).get("type"),
        "pdf_rel": pdf_rel,
        "pdf_meta": _meta(pdf_rel) if pdf_rel else None,
        "candidates": candidates,
        "steps": _expected_steps_for_row(row, role_info),
        "delete_label_prefixes": _delete_label_prefixes_for_row(row, role_info),
        "brand_hint": brand_hint,
        "model_hint": model_hint,
    })


@app.route("/api/reannotate/save", methods=["POST"])
def api_reannotate_save():
    """Body:
      {row, pdf_rel, page, steps:[{content_rect, label_rect, label_text}, ...],
       delete_existing: bool, delete_prefixes: [str, ...],
       col_d_override: str|null}
    """
    data = request.get_json(silent=True) or {}
    try:
        row_num = int(data["row"])
        pdf_rel = str(data["pdf_rel"])
        page = int(data["page"])
        steps = list(data.get("steps") or [])
    except Exception as e:
        return jsonify({"ok": False, "error": f"bad payload: {e}"}), 400

    if not steps:
        return jsonify({"ok": False, "error": "no steps to save"}), 400

    row = next((r for r in ROWS if r["row"] == row_num), None)
    if not row:
        return jsonify({"ok": False, "error": "row not found"}), 404

    pdf_path = (OUTPUT / pdf_rel).resolve()
    if (not str(pdf_path).startswith(str(OUTPUT.resolve())) or
            not pdf_path.exists()):
        return jsonify({"ok": False, "error": "PDF not found"}), 404

    role_info = detect_row_role(row_num)

    # Pre-snapshot — `pre-reannotate-row-N` shows up in the snapshots list
    _run_version_cmd(["snap", f"pre-reannotate-row-{row_num}"], timeout=60)

    # 1. Compute deletions (only if requested AND prefixes provided)
    edits: list[dict] = []
    deleted_xrefs: list[int] = []
    if data.get("delete_existing"):
        prefixes = list(data.get("delete_prefixes") or [])
        if prefixes:
            deleted_xrefs = _annots_to_delete_by_prefix(pdf_path, page, prefixes)
            for xref in deleted_xrefs:
                edits.append({"action": "delete", "xref": xref})

    # 2. Append create edits (Square + FreeText pairs)
    for st in steps:
        try:
            crect = list(map(float, st["content_rect"]))
            lrect = list(map(float, st["label_rect"]))
            ltext = (st.get("label_text") or "").strip()
        except Exception as e:
            return jsonify({"ok": False, "error": f"bad step: {e}"}), 400
        edits.append({"action": "create", "page": page, "type": "Square",
                      "rect": crect, "contents": ""})
        edits.append({"action": "create", "page": page, "type": "FreeText",
                      "rect": lrect, "contents": ltext})

    # 3. Apply
    pdf_result = apply_pdf_edits(pdf_path, edits)

    # 4. Compute new Col D
    col_d_override = (data.get("col_d_override") or "").strip()
    if col_d_override:
        new_d = col_d_override
    else:
        # If we're reannotating a brand_model section header AND the user
        # passed brand/model strings via override, use those; otherwise fall
        # through to the file-name-based generator.
        brand = (data.get("brand") or "").strip()
        model = (data.get("model") or "").strip()
        if brand and model:
            new_d = f"ยี่ห้อ {brand} รุ่น {model}"
        elif brand:
            new_d = f"ยี่ห้อ {brand}"
        else:
            new_d = make_col_d_for_row(row, role_info, pdf_path, page)
            if not new_d:
                new_d = (f"เทียบเท่าข้อกำหนด เอกสาร {pdf_path.stem} หน้า {page} "
                         f"ข้อ {role_info.get('section', '')}")

    old_d = row.get("D") or ""
    old_pdf_rel = row.get("pdf_rel") or ""

    # 5. Persist Col D + (when PDF changed) update xlsx
    try:
        wb = openpyxl.load_workbook(XLSX_PATH)
        ws = wb.active
        ws.cell(row_num, 4).value = new_d
        cur_c = (ws.cell(row_num, 3).value or "").strip()
        cur_b = (ws.cell(row_num, 2).value or "").strip()
        if not cur_c or cur_c == cur_b:
            ws.cell(row_num, 3).value = make_col_c_from_b(cur_b)
        tmp = XLSX_PATH.with_suffix(".xlsx.tmp")
        wb.save(str(tmp))
        with open(tmp, "rb") as f:
            data_bytes = f.read()
        fd = os.open(str(XLSX_PATH), os.O_WRONLY | os.O_TRUNC)
        try:
            os.write(fd, data_bytes)
        finally:
            os.close(fd)
        tmp.unlink(missing_ok=True)
    except Exception as e:
        return jsonify({"ok": False, "error": f"xlsx write failed: {e}",
                        "pdf_result": pdf_result}), 500

    # 6. Refresh + audit + learning feedback
    try:
        load_rows()
        TOR_CACHE.clear()
        sync_db_from_memory()
    except Exception as e:
        sys.stderr.write(f"[reannotate] refresh: {e}\n")

    try:
        db.log_audit(action="reannotate",
                     target_type="row", target_id=str(row_num),
                     before=old_d, after=new_d,
                     details={"page": page, "pdf_rel": pdf_rel,
                              "old_pdf_rel": old_pdf_rel,
                              "n_steps": len(steps),
                              "deleted_xrefs": deleted_xrefs,
                              "annots_applied": pdf_result.get("applied", 0),
                              "annots_errors": pdf_result.get("errors", 0)},
                     actor="user")
    except Exception: pass

    try:
        learning.record_feedback(
            row_num=row_num,
            section=row.get("section_inferred"),
            input_b=row.get("B", "") or "",
            input_pdf_rel=pdf_rel,
            input_role=role_info.get("role", ""),
            input_filename=pdf_path.name,
            suggested_c=row.get("C", "") or "",
            suggested_d=old_d,
            suggested_annots=[],
            confidence=0.0,
            generator="reannotate_wizard",
            provenance={"old_pdf_rel": old_pdf_rel,
                        "n_steps": len(steps),
                        "n_deleted": len(deleted_xrefs)},
            user_action="edited",
            final_c=row.get("C", "") or "",
            final_d=new_d,
            final_annots=edits,
        )
    except Exception as e:
        sys.stderr.write(f"[reannotate] feedback: {e}\n")

    return jsonify({
        "ok": True, "row": row_num, "page": page,
        "old_d": old_d, "new_d": new_d,
        "pdf_rel": pdf_rel,
        "pdf_changed": pdf_rel != old_pdf_rel,
        "deleted_xrefs": deleted_xrefs,
        "pdf_result": pdf_result,
    })


# ---------------------------------------------------------------------------
# Project-level versioning — wraps scripts/version.py
#
# version.py is the canonical snapshot tool (snap, snap-full, list, restore,
# diff, prune, auto-snap). Rather than reimplementing it, we shell out via
# subprocess so behavior stays in lock-step with the CLI tool.
# ---------------------------------------------------------------------------

def _version_script_available() -> bool:
    return VERSION_SCRIPT.exists()


def _read_snapshot(snap_dir: Path) -> dict | None:
    m_path = snap_dir / "manifest.json"
    if not m_path.exists():
        return None
    try:
        m = json.loads(m_path.read_text(encoding="utf-8"))
    except Exception:
        return None
    try:
        size = sum(p.stat().st_size for p in snap_dir.rglob("*") if p.is_file())
    except Exception:
        size = 0
    return {
        "id": m.get("id", snap_dir.name),
        "tag": m.get("tag", ""),
        "kind": m.get("kind", "?"),
        "timestamp": m.get("timestamp", ""),
        "size": size,
        "n_tracked": len(m.get("files", {})),
        "n_output": len(m.get("output_tree", [])),
        "has_tarball": (snap_dir / "output.tar.gz").exists(),
        "files": [
            {"path": k, "size": v.get("size", 0), "kind": v.get("kind", "")}
            for k, v in (m.get("files", {}) or {}).items()
        ],
    }


def _run_version_cmd(args: list[str], timeout: int = 600) -> dict:
    """Invoke version.py with the given args and return a JSON-serialisable
    summary."""
    if not _version_script_available():
        return {"ok": False, "error": "scripts/version.py not found",
                "stdout": "", "stderr": ""}
    cmd = [sys.executable, str(VERSION_SCRIPT)] + args
    try:
        r = subprocess.run(cmd, cwd=str(ROOT),
                           capture_output=True, text=True, timeout=timeout)
    except subprocess.TimeoutExpired:
        return {"ok": False, "error": "timeout", "stdout": "", "stderr": ""}
    except Exception as e:
        return {"ok": False, "error": str(e), "stdout": "", "stderr": ""}
    return {
        "ok": r.returncode == 0,
        "returncode": r.returncode,
        "stdout": r.stdout,
        "stderr": r.stderr,
    }


@app.route("/api/versions")
def api_versions():
    """List project snapshots from _versions/snapshots/."""
    snaps = []
    if SNAPS_DIR.exists():
        for d in sorted(SNAPS_DIR.iterdir(), reverse=True):
            if not d.is_dir():
                continue
            info = _read_snapshot(d)
            if info:
                snaps.append(info)
    return jsonify({
        "available": _version_script_available(),
        "root": str(ROOT),
        "snapshots": snaps,
    })


@app.route("/api/versions/snap", methods=["POST"])
def api_versions_snap():
    """Create a new snapshot. Body: {tag?: str, full?: bool}."""
    data = request.get_json(silent=True) or {}
    tag = (data.get("tag") or "").strip()
    full = bool(data.get("full"))
    args = ["snap-full" if full else "snap"]
    if tag:
        args.append(tag)
    result = _run_version_cmd(args, timeout=600 if full else 60)
    if result.get("ok"):
        try:
            db.log_audit(action="snapshot",
                         target_type="project", target_id=tag or "(no tag)",
                         details={"kind": "full" if full else "quick", "tag": tag},
                         actor="user")
        except Exception: pass
    return jsonify(result)


@app.route("/api/versions/restore", methods=["POST"])
def api_versions_restore():
    """Restore from snapshot. Body: {id: str, full?: bool}.

    Always takes an auto-snap of the current state first (handled by version.py
    only when --pre-snap is used in some setups; here we explicitly create one
    via auto-snap to be safe).
    """
    data = request.get_json(silent=True) or {}
    snap_id = (data.get("id") or "").strip()
    full = bool(data.get("full"))
    if not snap_id:
        return jsonify({"ok": False, "error": "id required"}), 400
    # Take a safety snapshot of the current state first
    _run_version_cmd(["snap", "pre-restore"], timeout=60)
    result = _run_version_cmd(
        ["restore-full" if full else "restore", snap_id, "-y"],
        timeout=600 if full else 60,
    )
    if result.get("ok"):
        # Refresh internal caches after restore
        try:
            build_pdf_index()
            load_rows()
            collect_extra_refs()
            build_tor_section_index()
            index_tor_text()
            TOR_CACHE.clear()
            sync_db_from_memory()
        except Exception as e:
            sys.stderr.write(f"[restore refresh] {e}\n")
        try:
            db.log_audit(action="restore",
                         target_type="project", target_id=snap_id,
                         details={"full": full}, actor="user")
        except Exception: pass
    return jsonify(result)


@app.route("/api/versions/diff")
def api_versions_diff():
    id1 = (request.args.get("id1") or "").strip()
    id2 = (request.args.get("id2") or "").strip()
    if not id1 or not id2:
        return jsonify({"ok": False, "error": "id1 + id2 required"}), 400
    result = _run_version_cmd(["diff", id1, id2], timeout=60)
    return jsonify(result)


@app.route("/api/versions/auto-snap", methods=["POST"])
def api_versions_auto_snap():
    """Snap only if the xlsx changed since the most recent snapshot."""
    result = _run_version_cmd(["auto-snap"], timeout=60)
    return jsonify(result)


@app.route("/api/versions/prune", methods=["POST"])
def api_versions_prune():
    data = request.get_json(silent=True) or {}
    keep = max(1, int(data.get("keep", 10)))
    result = _run_version_cmd(["prune", "--keep", str(keep), "-y"], timeout=60)
    return jsonify(result)


@app.route("/api/versions/show")
def api_versions_show():
    snap_id = (request.args.get("id") or "").strip()
    if not snap_id:
        return jsonify({"ok": False, "error": "id required"}), 400
    # Find snap by exact ID first, then by prefix
    candidates = []
    if SNAPS_DIR.exists():
        for d in SNAPS_DIR.iterdir():
            if d.is_dir() and (d.name == snap_id or d.name.startswith(snap_id)):
                candidates.append(d)
    if len(candidates) != 1:
        return jsonify({"ok": False, "error": f"{len(candidates)} matches"}), 404
    info = _read_snapshot(candidates[0])
    return jsonify({"ok": True, **(info or {})})


# ---------------------------------------------------------------------------
# "Always work on latest" — sync the live working files with the most recent
# snapshot
# ---------------------------------------------------------------------------

def _file_sha256(path: Path) -> str:
    h = hashlib.sha256()
    try:
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(65536), b""):
                h.update(chunk)
    except Exception:
        return ""
    return h.hexdigest()


def _find_latest_snapshot() -> tuple[Path, dict] | None:
    if not SNAPS_DIR.exists():
        return None
    snaps = sorted(
        [d for d in SNAPS_DIR.iterdir() if d.is_dir()],
        reverse=True,
    )
    for snap_dir in snaps:
        m_path = snap_dir / "manifest.json"
        if not m_path.exists():
            continue
        try:
            return snap_dir, json.loads(m_path.read_text(encoding="utf-8"))
        except Exception:
            continue
    return None


def get_version_sync_status() -> dict:
    """Compare the live working files against the most recent valid snapshot.

    Returns a state describing the relationship:

    - ``no_snapshots``      — versioning system is empty
    - ``in_sync``           — every tracked file matches the latest snapshot
    - ``working_ahead``     — files differ AND working mtimes ≥ snapshot ts
                              (the user has been editing since last snapshot;
                              the boot hook auto-snaps to bring "latest"
                              forward to match)
    - ``working_behind``    — files differ AND working mtimes < snapshot ts
                              (something rolled the working dir back; UI
                              prompts the user to restore)
    - ``divergent``         — mixed (some newer, some older) — UI prompts
    - ``incomplete_local``  — a tracked file is missing from the working dir
    """
    out: dict = {
        "has_snapshots": False,
        "state": "no_snapshots",
        "latest": None,
        "files": [],
    }
    found = _find_latest_snapshot()
    if found is None:
        return out
    snap_dir, m = found

    out["has_snapshots"] = True
    out["latest"] = {
        "id": m.get("id", snap_dir.name),
        "tag": m.get("tag", ""),
        "kind": m.get("kind", "?"),
        "timestamp": m.get("timestamp", ""),
    }

    snap_ts = 0.0
    if m.get("timestamp"):
        try:
            snap_ts = datetime.fromisoformat(m["timestamp"]).timestamp()
        except Exception:
            pass

    all_match = True
    any_newer = False
    any_older = False
    any_missing = False

    for rel, info in (m.get("files") or {}).items():
        if info.get("kind") == "tarball":
            continue
        full = ROOT / rel
        if not full.exists():
            all_match = False
            any_missing = True
            out["files"].append({"path": rel, "status": "missing_local"})
            continue
        cur_hash = _file_sha256(full)
        snap_hash = info.get("sha256", "")
        match = (cur_hash == snap_hash)
        cur_mtime = full.stat().st_mtime
        if not match:
            all_match = False
            if cur_mtime >= snap_ts:
                any_newer = True
            else:
                any_older = True
        out["files"].append({
            "path": rel,
            "status": "match" if match else "differ",
            "current_size": full.stat().st_size,
            "snapshot_size": info.get("size", 0),
            "current_mtime": cur_mtime,
            "snapshot_mtime": snap_ts,
        })

    if all_match:
        out["state"] = "in_sync"
    elif any_missing:
        out["state"] = "incomplete_local"
    elif any_newer and not any_older:
        out["state"] = "working_ahead"
    elif any_older and not any_newer:
        out["state"] = "working_behind"
    else:
        out["state"] = "divergent"
    return out


@app.route("/api/versions/sync")
def api_versions_sync():
    return jsonify(get_version_sync_status())


def boot_sync_check() -> dict:
    """Run at server start. Tries to keep the invariant ``latest = current``:

    - If no snapshot exists yet → take a baseline snap so versioning is
      bootstrapped.
    - If working state is **ahead** of latest (the common case after a work
      session) → take an auto-snap so latest reflects the new state.
    - If working state is **behind** or **divergent** → DO NOTHING
      automatically (would risk overwriting in-progress work). Surface the
      condition through the UI so the user can decide.

    Set env ``COMPLY_NO_BOOT_SNAP=1`` to disable boot auto-snapping.
    """
    result = {"performed": None, "status": None}
    if not _version_script_available():
        result["status"] = "version_script_missing"
        return result
    status = get_version_sync_status()
    result["status"] = status["state"]

    if os.environ.get("COMPLY_NO_BOOT_SNAP") == "1":
        return result

    state = status["state"]
    if state == "no_snapshots":
        r = _run_version_cmd(["snap", "boot-baseline"], timeout=60)
        result["performed"] = "baseline_snap" if r.get("ok") else f"failed: {r.get('error')}"
    elif state == "working_ahead":
        # auto-snap creates a tagged snapshot only if xlsx actually changed
        # (it's idempotent)
        r = _run_version_cmd(["snap", "boot-auto"], timeout=60)
        result["performed"] = "boot_auto_snap" if r.get("ok") else f"failed: {r.get('error')}"
    elif state in ("working_behind", "divergent", "incomplete_local"):
        result["performed"] = "skipped (manual review required)"
    else:
        result["performed"] = "no-op (already in sync)"

    return result


# ---------------------------------------------------------------------------
# DB sync + DB-backed API endpoints
# ---------------------------------------------------------------------------

def _build_pdf_records_for_db() -> list[dict]:
    """Snapshot the PDF index into dict records for the DB.

    Annotations are NOT pre-listed — they used to be (via list_pdf_annots
    per PDF), which made boot ~60-90 sec on Google Drive filesystems
    because PyMuPDF opens each of 101 PDFs and walks every annot. The
    pdf_annotations table is only ever WRITTEN (and counted), never
    queried — annotations are fetched live via /api/pdf_meta when a row
    is selected. Keeping it empty saves ~minute on every boot.
    """
    seen: dict[Path, dict] = {}
    # Combine both indexes
    for ref_key, paths in PDF_INDEX.items():
        for p in paths:
            if p in seen:
                continue
            seen[p] = {"folder_key": ref_key}
    for sec_key, paths in SECTION_INDEX.items():
        for p in paths:
            seen.setdefault(p, {"folder_key": None,
                                "section_prefix": sec_key})
            seen[p].setdefault("section_prefix", sec_key)

    records = []
    for p, meta in seen.items():
        try:
            stat = p.stat()
            rel = str(p.relative_to(OUTPUT))
        except Exception:
            continue
        # Try detect brand/model from filename (fast, no PDF open)
        try:
            brand, model = parse_brand_model_from_filename(p.stem)
        except Exception:
            brand, model = ("", "")
        records.append({
            "rel_path": rel,
            "folder_key": meta.get("folder_key"),
            "section_prefix": meta.get("section_prefix"),
            "size": stat.st_size,
            "mtime": stat.st_mtime,
            "num_pages": None,        # filled lazily via /api/pdf_meta
            "brand": brand or None,
            "model": model or None,
            "annotations": [],         # empty — see docstring
        })
    return records


def sync_db_from_memory() -> None:
    """Mirror the in-memory rows + PDFs + TOR + snapshots into the DB."""
    if db.db_path() is None:
        return
    try:
        db.sync_rows(ROWS)
    except Exception as e:
        sys.stderr.write(f"[db] sync_rows: {e}\n")
    try:
        db.sync_pdfs(_build_pdf_records_for_db())
    except Exception as e:
        sys.stderr.write(f"[db] sync_pdfs: {e}\n")
    try:
        db.sync_tor(TOR_SECTION_INDEX, TOR_PAGE_TEXTS)
    except Exception as e:
        sys.stderr.write(f"[db] sync_tor: {e}\n")
    # Snapshots — read from filesystem
    try:
        snaps = []
        if SNAPS_DIR.exists():
            for d in SNAPS_DIR.iterdir():
                if d.is_dir():
                    info = _read_snapshot(d)
                    if info:
                        snaps.append(info)
        db.sync_snapshots(snaps)
    except Exception as e:
        sys.stderr.write(f"[db] sync_snapshots: {e}\n")
    # PDF history (output/_pdf_history/<flat>/*.pdf)
    try:
        records = []
        if PDF_HISTORY.exists():
            for flat_dir in PDF_HISTORY.iterdir():
                if not flat_dir.is_dir():
                    continue
                for snap in flat_dir.glob("*.pdf"):
                    st = snap.stat()
                    records.append({
                        "pdf_rel": flat_dir.name,
                        "snapshot_filename": snap.name,
                        "ts": datetime.fromtimestamp(st.st_mtime).isoformat(timespec="seconds"),
                        "tag": "",
                        "size": st.st_size,
                    })
        db.sync_pdf_history(records)
    except Exception as e:
        sys.stderr.write(f"[db] sync_pdf_history: {e}\n")


@app.route("/api/db/stats")
def api_db_stats():
    return jsonify(db.stats_summary())


@app.route("/api/db/audit")
def api_db_audit():
    limit = int(request.args.get("limit", 100))
    action = request.args.get("action")
    return jsonify({
        "entries": db.recent_audit(limit=min(500, max(1, limit)),
                                   action_filter=action),
    })


@app.route("/api/db/search")
def api_db_search():
    q = request.args.get("q", "").strip()
    if not q:
        return jsonify({"results": []})
    return jsonify({"results": db.fts_search(q, limit=80)})


@app.route("/api/db/section_progress")
def api_db_section_progress():
    return jsonify({"sections": db.section_progress()})


# ---------------------------------------------------------------------------
# HITL learning endpoints
# ---------------------------------------------------------------------------

@app.route("/api/learn/stats")
def api_learn_stats():
    days = int(request.args.get("days", 30))
    return jsonify(learning.feedback_stats(window_days=max(1, min(365, days))))


@app.route("/api/learn/patterns")
def api_learn_patterns():
    ptype = request.args.get("type")
    return jsonify({"patterns": learning.list_patterns(pattern_type=ptype)})


@app.route("/api/learn/retrain", methods=["POST"])
def api_learn_retrain():
    result = learning.retrain_patterns()
    try:
        db.log_audit(action="retrain", target_type="learning",
                     details=result, actor="user")
    except Exception: pass
    return jsonify({"ok": True, **result})


@app.route("/api/learn/pin_pattern", methods=["POST"])
def api_learn_pin_pattern():
    """Force-learn a pattern from a single row (Sprint S2.1).

    User clicks 'Pin as template' on a row they've just verified. We
    register the row's (filename, section, role, final Col D) tuple as
    learned_patterns with confidence=1.0 and samples_total=1, regardless
    of the usual PROMOTION_THRESHOLD=2. Future rows matching the trigger
    will use this template directly.
    """
    data = request.get_json(silent=True) or {}
    try:
        row_num = int(data.get("row", 0))
    except Exception:
        return jsonify({"ok": False, "error": "bad row"}), 400
    row = next((r for r in ROWS if r["row"] == row_num), None)
    if not row:
        return jsonify({"ok": False, "error": "row not found"}), 404

    final_d = (row.get("D") or "").strip()
    if not final_d:
        return jsonify({"ok": False, "error": "row has no Col D to pin"}), 400

    section = row.get("section_inferred")
    pdf_rel = row.get("pdf_rel")
    role_info = detect_row_role(row_num)
    role = role_info.get("role", "unknown")

    pinned = []
    with db.conn() as c:
        # 1. filename_brand pattern (if Col D is brand_model)
        if final_d.startswith("ยี่ห้อ ") and pdf_rel:
            stem = Path(pdf_rel).stem
            tokens = re.findall(r"[A-Za-z][A-Za-z0-9]{2,}", stem)
            m = re.match(r"ยี่ห้อ\s+(\S+)\s+รุ่น\s+", final_d)
            if m and tokens:
                trigger = next((t for t in tokens if len(t) >= 3), tokens[0]).lower()
                brand = m.group(1)
                _upsert_pattern(c, "filename_brand", trigger, None, brand,
                                samples=1, correct=1, confidence=1.0,
                                note=f"pinned from R{row_num}")
                pinned.append(("filename_brand", trigger, brand))

        # 2. row_format_d pattern (always — captures shape preference)
        sec_root = ".".join(section.split(".")[:2]) if section else "unknown"
        shape = learning._shape_of_col_d(final_d)
        _upsert_pattern(c, "row_format_d", role, json.dumps([sec_root]),
                        shape, samples=1, correct=1, confidence=1.0,
                        note=f"pinned from R{row_num}")
        pinned.append(("row_format_d", f"{role}/{sec_root}", shape))

    try:
        db.log_audit(action="pin_pattern", target_type="row",
                     target_id=str(row_num),
                     details={"pinned": pinned, "final_d": final_d},
                     actor="user")
    except Exception: pass

    return jsonify({"ok": True, "row": row_num, "pinned": pinned,
                    "n_patterns": len(pinned)})


def _upsert_pattern(c, ptype, trigger, trigger_extra, value,
                     samples=1, correct=1, confidence=1.0, note=None):
    """Helper for pin_pattern — insert or update a learned_patterns row."""
    existing = c.execute(
        """SELECT pattern_id, samples_total, samples_correct
           FROM learned_patterns
           WHERE pattern_type=? AND trigger_key=?
             AND ifnull(trigger_extra,'')=ifnull(?, '')""",
        (ptype, trigger, trigger_extra),
    ).fetchone()
    if existing:
        c.execute(
            """UPDATE learned_patterns
               SET output_value=?, samples_total=samples_total+?,
                   samples_correct=samples_correct+?,
                   confidence=?, enabled=1,
                   note=COALESCE(?, note),
                   last_used_at=CURRENT_TIMESTAMP
               WHERE pattern_id=?""",
            (value, samples, correct, confidence, note, existing["pattern_id"]),
        )
    else:
        c.execute(
            """INSERT INTO learned_patterns
               (pattern_type, trigger_key, trigger_extra, output_value,
                samples_total, samples_correct, confidence, enabled, note)
               VALUES (?, ?, ?, ?, ?, ?, ?, 1, ?)""",
            (ptype, trigger, trigger_extra, value, samples, correct, confidence, note),
        )


@app.route("/api/settings/api_key", methods=["POST"])
def api_settings_save_api_key():
    """Save the Anthropic API key from the frontend.

    The key is written to ROOT/.env (gitignored), the Anthropic provider
    singleton is reset, and we reinstall it into the learning hook so
    subsequent auto-annotate calls go through Claude immediately — no
    restart needed.
    """
    data = request.get_json(silent=True) or {}
    key = (data.get("api_key") or "").strip()
    model = (data.get("model") or "").strip()
    budget = data.get("budget_usd_per_day")

    if key and not key.startswith("sk-ant-"):
        return jsonify({"ok": False, "error": "API key must start with 'sk-ant-'"}), 400

    env_path = ROOT / ".env"
    # Read existing .env (if present) and replace/append the keys
    lines: list[str] = []
    have = {"ANTHROPIC_API_KEY": False, "COMPLY_LLM": False,
            "COMPLY_LLM_MODEL": False, "COMPLY_LLM_BUDGET_USD_PER_DAY": False}
    try:
        if env_path.exists():
            for raw in env_path.read_text(encoding="utf-8").splitlines():
                stripped = raw.strip()
                if stripped.startswith("#") or "=" not in stripped:
                    lines.append(raw); continue
                k = stripped.split("=", 1)[0].strip()
                if k == "ANTHROPIC_API_KEY":
                    if key:
                        lines.append(f"ANTHROPIC_API_KEY={key}")
                        have[k] = True
                    # else: drop the line (user cleared the key)
                elif k == "COMPLY_LLM":
                    lines.append("COMPLY_LLM=anthropic" if key else "COMPLY_LLM=")
                    have[k] = True
                elif k == "COMPLY_LLM_MODEL":
                    if model:
                        lines.append(f"COMPLY_LLM_MODEL={model}"); have[k] = True
                    else:
                        lines.append(raw); have[k] = True
                elif k == "COMPLY_LLM_BUDGET_USD_PER_DAY":
                    if budget is not None:
                        lines.append(f"COMPLY_LLM_BUDGET_USD_PER_DAY={float(budget):.2f}")
                        have[k] = True
                    else:
                        lines.append(raw); have[k] = True
                else:
                    lines.append(raw)
    except Exception as e:
        return jsonify({"ok": False, "error": f"failed to read .env: {e}"}), 500

    # Append any missing keys
    if key and not have["ANTHROPIC_API_KEY"]:
        lines.append(f"ANTHROPIC_API_KEY={key}")
    if not have["COMPLY_LLM"]:
        lines.append("COMPLY_LLM=anthropic" if key else "COMPLY_LLM=")
    if not have["COMPLY_LLM_MODEL"] and model:
        lines.append(f"COMPLY_LLM_MODEL={model}")
    if not have["COMPLY_LLM_BUDGET_USD_PER_DAY"] and budget is not None:
        lines.append(f"COMPLY_LLM_BUDGET_USD_PER_DAY={float(budget):.2f}")

    # Write back atomically
    try:
        tmp = env_path.with_suffix(".env.tmp")
        tmp.write_text("\n".join(lines) + "\n", encoding="utf-8")
        tmp.replace(env_path)
        # Also restrict perms (key is sensitive)
        try: env_path.chmod(0o600)
        except Exception: pass
    except Exception as e:
        return jsonify({"ok": False, "error": f"failed to write .env: {e}"}), 500

    # Live-reload env vars + provider
    if key:
        os.environ["ANTHROPIC_API_KEY"] = key
        os.environ["COMPLY_LLM"] = "anthropic"
        if model: os.environ["COMPLY_LLM_MODEL"] = model
        if budget is not None: os.environ["COMPLY_LLM_BUDGET_USD_PER_DAY"] = str(float(budget))
    else:
        os.environ.pop("ANTHROPIC_API_KEY", None)
        os.environ["COMPLY_LLM"] = ""

    # Reset and reinstall the provider singleton
    try:
        from app import anthropic_provider as ap
        ap._provider = None
        installed = ap.install_into_learning() if key else False
        status = ap.get_provider().budget_status() if installed else {"available": False}
    except Exception as e:
        return jsonify({"ok": False, "error": f"provider init: {e}"}), 500

    try:
        db.log_audit(action="set_api_key", target_type="settings",
                     details={"model": model, "budget": budget,
                              "cleared": not bool(key)},
                     actor="user")
    except Exception: pass

    return jsonify({"ok": True, "installed": bool(key) and installed,
                    "status": status})


@app.route("/api/learn/llm_status")
def api_learn_llm_status():
    """LLM provider info + budget snapshot (Sprint S3 — surfaced in Settings).

    Phase 1 (Claude Code as core): try Claude Code provider FIRST, fall back
    to Anthropic API provider. Returns a ``provider_kind`` field so the UI
    can render the right copy ('Claude Max OAuth' vs 'API key + budget').
    """
    base = learning.llm_status()
    base["provider_kind"] = "off"
    # Try Claude Code first
    try:
        from app import claude_code_provider as cp
        p = cp.get_provider()
        if p:
            base.update(p.budget_status())
            base["provider_kind"] = "claude_code"
    except Exception as e:
        base["error_claude_code"] = str(e)
    # Fall back to Anthropic API
    if base.get("provider_kind") == "off":
        try:
            from app import anthropic_provider as ap
            p = ap.get_provider()
            if p:
                base.update(p.budget_status())
                base["provider_kind"] = "anthropic_api"
        except Exception as e:
            base["error_anthropic"] = str(e)
    # Today's call count + last call info (works for both providers)
    if base.get("provider_kind") != "off":
        try:
            with db.conn() as c:
                row = c.execute(
                    """SELECT COUNT(*) AS n,
                              COALESCE(SUM(input_tokens),0) AS in_tok,
                              COALESCE(SUM(output_tokens),0) AS out_tok,
                              COALESCE(SUM(cache_read_tokens),0) AS cache_read,
                              COALESCE(MAX(ts),'') AS last_ts
                       FROM llm_calls
                       WHERE substr(ts, 1, 10) = strftime('%Y-%m-%d', 'now')""",
                ).fetchone()
                if row:
                    base.update({
                        "calls_today": row["n"],
                        "tokens_in_today": row["in_tok"],
                        "tokens_out_today": row["out_tok"],
                        "tokens_cache_read_today": row["cache_read"],
                        "last_call_ts": row["last_ts"],
                    })
        except Exception as e:
            base["error_calls"] = str(e)
    return jsonify(base)


@app.route("/api/claude/stream")
def api_claude_stream():
    """Phase 1 (Claude Code as core): stream Claude's reasoning live via SSE.

    Frontend opens an EventSource on this endpoint per row. Each line is a
    JSON event from ``ClaudeCodeProvider.propose_streaming``:
      • {type:"thinking", text:...}
      • {type:"tool_use",  name:..., input:{...}}     ← Read/Grep/propose_*
      • {type:"tool_result", name:..., text:...}
      • {type:"text",      content:...}
      • {type:"result",    proposal:{...}, usage:{...}, cost_usd, elapsed_ms}
      • {type:"error",     error:...}

    The endpoint runs the async generator on a private event loop in this
    request's thread (Flask is sync). For one user that's plenty; if we
    ever scale, swap to Hypercorn/Quart.
    """
    try:
        row_num = int(request.args.get("row", "0"))
    except ValueError:
        return jsonify({"ok": False, "error": "row required"}), 400
    row = next((r for r in ROWS if r["row"] == row_num), None)
    if not row:
        return jsonify({"ok": False, "error": "row not found"}), 404

    # Lazy-import to keep boot fast when SDK isn't installed
    try:
        from app import claude_code_provider as cp
    except ImportError:
        return jsonify({"ok": False, "error": "claude-agent-sdk not installed"}), 503
    p = cp.get_provider()
    if p is None:
        return jsonify({
            "ok": False,
            "error": "Claude Code provider unavailable",
            "hint": "Set COMPLY_LLM=claude_code and ensure 'claude' CLI is logged in",
        }), 503

    # Build row_context the same shape AnthropicProvider expects
    role_info = detect_row_role(row_num)
    pdf_filename = None
    if row.get("pdf_rel"):
        pdf_filename = Path(row["pdf_rel"]).name
    # Lightweight rule-based proposal as a hint for Claude
    rule_plan = None
    try:
        rp = auto_annotate_plan(row_num)
        if rp.get("ok"):
            rule_plan = {
                "proposed_d": rp.get("proposed_d", ""),
                "generator": rp.get("generator", ""),
                "confidence": rp.get("confidence", 0),
            }
    except Exception:
        pass
    row_context = {
        "row": row_num,
        "section": row.get("section_inferred"),
        "role": role_info.get("role", "unknown"),
        "col_a": row.get("A"),
        "col_b": row.get("B", ""),
        "col_c_current": row.get("C", ""),
        "col_d_current": row.get("D", ""),
        "col_e": row.get("E", ""),
        "pdf_rel": row.get("pdf_rel"),
        "pdf_filename": pdf_filename,
        "rule_proposal": rule_plan,
    }

    import asyncio as _aio

    def event_stream():
        loop = _aio.new_event_loop()
        try:
            agen = p.propose_streaming(row_context=row_context)
            agen_iter = agen.__aiter__()
            while True:
                try:
                    event = loop.run_until_complete(agen_iter.__anext__())
                except StopAsyncIteration:
                    break
                yield f"data: {json.dumps(event, ensure_ascii=False)}\n\n"
            yield "event: done\ndata: {}\n\n"
        except Exception as e:
            err = {"type": "error", "error": str(e)}
            yield f"data: {json.dumps(err, ensure_ascii=False)}\n\n"
        finally:
            try:
                loop.close()
            except Exception:
                pass

    headers = {
        "Cache-Control": "no-cache",
        "X-Accel-Buffering": "no",       # disable nginx buffering if proxied
        "Connection": "keep-alive",
    }
    return Response(event_stream(), mimetype="text/event-stream", headers=headers)


# (Catalog/companies/projects routes moved to app/routes/catalog_api.py — T2.2, 2026-05-10)


# ─── Apply catalog → row (writes Col D + records link) ───────────────

@app.route("/api/row/apply_catalog", methods=["POST"])
def api_row_apply_catalog():
    """Bind a catalog (and optional page) to a project row, generating Col D.

    Body: {row, catalog_id, page?, col_d_text?}

    If ``col_d_text`` is omitted we synthesize one using the same convention
    as the existing rule-based generator (``make_col_d_for_row``).
    """
    data = request.get_json(silent=True) or {}
    try:
        row_num = int(data["row"])
        cat_id = int(data["catalog_id"])
    except (KeyError, ValueError):
        return jsonify({"ok": False, "error": "row, catalog_id required"}), 400

    row = next((r for r in ROWS if r["row"] == row_num), None)
    if not row:
        return jsonify({"ok": False, "error": "row not found"}), 404
    cat = catalog.get_catalog(cat_id)
    if not cat:
        return jsonify({"ok": False, "error": "catalog not found"}), 404

    page = int(data["page"]) if data.get("page") else None
    col_d = (data.get("col_d_text") or "").strip()
    if not col_d:
        # Synthesize: use existing generator if possible
        try:
            role_info = detect_row_role(row_num)
            pdf_path = OUTPUT / cat["pdf_rel"]
            col_d = make_col_d_for_row(row, role_info, pdf_path, page) or ""
        except Exception as e:
            sys.stderr.write(f"[apply_catalog] synth col_d failed: {e}\n")
            col_d = f"เอกสาร {cat.get('section_hint', '?')} {cat.get('brand') or ''} {cat.get('model') or ''}".strip()
            if page:
                col_d += f" หน้า {page}"

    proj = catalog.get_active_project()
    if not proj:
        return jsonify({"ok": False, "error": "no active project"}), 500

    # Pre-snap, write xlsx Col D, then refresh + record link
    _run_version_cmd(["snap", f"pre-apply-catalog-row-{row_num}"], timeout=60)

    try:
        wb = openpyxl.load_workbook(XLSX_PATH)
        ws = wb.active
        original = (ws.cell(row_num, 4).value or "").strip()
        ws.cell(row_num, 4).value = col_d
        tmp = XLSX_PATH.with_suffix(".xlsx.tmp")
        wb.save(str(tmp))
        with open(tmp, "rb") as f:
            buf = f.read()
        fd = os.open(str(XLSX_PATH), os.O_WRONLY | os.O_TRUNC)
        try:
            os.write(fd, buf)
        finally:
            os.close(fd)
        tmp.unlink(missing_ok=True)
    except Exception as e:
        return jsonify({"ok": False, "error": f"xlsx write failed: {e}"}), 500

    # Update memory + DB mirror
    load_rows()
    sync_db_from_memory()

    # Record link
    catalog.bind_row_to_catalog(
        project_id=int(proj["project_id"]),
        row_num=row_num, catalog_id=cat_id,
        page=page, col_d_text=col_d,
    )
    db.log_audit(action="apply_catalog", target_type="row",
                 target_id=str(row_num),
                 before=original, after=col_d,
                 details={"catalog_id": cat_id, "page": page}, actor="user")

    return jsonify({"ok": True, "row": row_num, "col_d": col_d,
                    "catalog_id": cat_id, "page": page})


# (Export routes moved to app/routes/export_api.py — T2.2, 2026-05-10)


# (Continuity routes moved to app/routes/continuity_api.py — T2.2, 2026-05-10)

@app.route("/api/learn/feedback", methods=["POST"])
def api_learn_feedback():
    """Direct feedback recorder (e.g., when user edits Col D outside the
    auto-annotate flow but still wants the system to learn)."""
    data = request.get_json(silent=True) or {}
    required = ("row_num", "user_action", "final_d")
    if any(k not in data for k in required):
        return jsonify({"ok": False, "error": f"missing {required}"}), 400
    fb_id = learning.record_feedback(
        row_num=int(data["row_num"]),
        section=data.get("section"),
        input_b=data.get("input_b", ""),
        input_pdf_rel=data.get("input_pdf_rel"),
        input_role=data.get("input_role", ""),
        input_filename=data.get("input_filename"),
        suggested_c=data.get("suggested_c", ""),
        suggested_d=data.get("suggested_d", ""),
        suggested_annots=data.get("suggested_annots") or [],
        confidence=float(data.get("confidence", 0)),
        generator=data.get("generator", "manual"),
        provenance=data.get("provenance"),
        user_action=data["user_action"],
        final_c=data.get("final_c", ""),
        final_d=data["final_d"],
        final_annots=data.get("final_annots") or [],
    )
    return jsonify({"ok": True, "fb_id": fb_id})


@app.route("/api/learn/pattern/<int:pattern_id>", methods=["POST"])
def api_learn_toggle_pattern(pattern_id: int):
    """Enable/disable a learned pattern (or override its output)."""
    data = request.get_json(silent=True) or {}
    fields = []
    params: list = []
    if "enabled" in data:
        fields.append("enabled = ?"); params.append(1 if data["enabled"] else 0)
    if "output_value" in data:
        fields.append("output_value = ?"); params.append(str(data["output_value"]))
    if "note" in data:
        fields.append("note = ?"); params.append(str(data["note"]))
    if not fields:
        return jsonify({"ok": False, "error": "nothing to update"}), 400
    params.append(pattern_id)
    with db.conn() as c:
        c.execute(f"UPDATE learned_patterns SET {', '.join(fields)} WHERE pattern_id = ?",
                  params)
    return jsonify({"ok": True})


# ---------------------------------------------------------------------------
# HTML template
# ---------------------------------------------------------------------------

# (INDEX_HTML moved to app/server/templates/index.html in T2.1 — 2026-05-10)




# ---------------------------------------------------------------------------
# Boot
# ---------------------------------------------------------------------------

def boot() -> None:
    print(f"[boot] root={ROOT}")

    # Bring up the SQLite layer first so subsequent index/sync calls have
    # somewhere to mirror to.
    print(f"[boot] opening database at {DB_PATH.relative_to(ROOT)}…")
    db.init_db(DB_PATH)
    # One-time migration: import legacy verification_status.json if the DB
    # table is empty
    try:
        existing = db.get_all_status()
    except Exception:
        existing = {}
    if not existing and STATUS_PATH.exists():
        n = db.import_status_from_json(STATUS_PATH)
        if n:
            print(f"[boot] migrated {n} verification_status entries → DB")

    print(f"[boot] indexing PDFs in {OUTPUT}…")
    build_pdf_index()
    n_indexed = len({p for v in PDF_INDEX.values() for p in v} |
                    {p for v in SECTION_INDEX.values() for p in v})
    print(f"[boot] {len(PDF_INDEX)} ref-keys + {len(SECTION_INDEX)} section-keys → "
          f"{n_indexed} unique PDFs indexed")
    print(f"[boot] loading {XLSX_PATH.name}…")
    load_rows()
    print(f"[boot] {len(ROWS)} rows")
    collect_extra_refs()
    print("[boot] indexing TOR sections + page text…")
    build_tor_section_index()
    index_tor_text()
    ch5 = sum(1 for s in TOR_SECTION_INDEX if s.startswith("5"))
    print(f"[boot] {len(TOR_SECTION_INDEX)} TOR sections ({ch5} in chapter 5), "
          f"{len(TOR_PAGE_TEXTS)} pages indexed for text-search")

    # Mirror everything into the DB so queries can hit it instead of the
    # in-memory caches.
    sync_db_from_memory()

    # Phase 2: catalog library bootstrap (idempotent)
    # Ensure default company + active project exist so row_catalog_links
    # has a project_id to point at, then ingest output/ PDFs as catalogs.
    try:
        cm_id = catalog.upsert_company(name="Smart Solution", code="SMART")
        proj = catalog.get_active_project()
        if not proj:
            pid = catalog.upsert_project(
                company_id=cm_id, name="Smart Plant 1", code="SP1",
                xlsx_rel=str(XLSX_PATH.relative_to(ROOT)),
                output_rel=str(OUTPUT.relative_to(ROOT)),
            )
            catalog.set_active_project(pid)
            proj = catalog.get_active_project()
        result = catalog.ingest_output_dir(OUTPUT)
        if result.get("ok"):
            print(f"[boot] catalog library: {result['scanned']} PDFs scanned "
                  f"({result['inserted']} new, {result['updated']} updated, "
                  f"{result['skipped']} unchanged) · "
                  f"active project={proj.get('name') if proj else '?'}")
    except Exception as e:
        sys.stderr.write(f"[boot] catalog ingest failed: {e}\n")

    # Load .env file (if present) so ANTHROPIC_API_KEY etc. are available
    _load_dotenv(ROOT / ".env")

    # Install Claude provider — Claude Code (Agent SDK) takes priority over
    # API direct, since Phase 1 uses Claude Max OAuth (no metered API costs).
    # Set COMPLY_LLM=claude_code (default) or =anthropic for the legacy path.
    _llm_mode = (os.environ.get("COMPLY_LLM") or "claude_code").lower()
    if _llm_mode in ("claude_code", "claude-code"):
        try:
            from app import claude_code_provider as cp
            if cp.install_into_learning():
                p = cp.get_provider()
                bs = p.budget_status() if p else {}
                print(f"[boot] Claude Code provider installed: model={bs.get('model')} "
                      f"auth={bs.get('auth_mode')} (Max = unlimited)")
            else:
                if not cp.SDK_AVAILABLE:
                    print("[boot] Claude Code provider OFF (claude-agent-sdk not installed)")
                else:
                    print("[boot] Claude Code provider OFF (set COMPLY_LLM=claude_code to enable)")
        except Exception as e:
            sys.stderr.write(f"[boot] Claude Code install failed: {e}\n")
    elif _llm_mode == "anthropic":
        try:
            from app import anthropic_provider as ap
            if ap.install_into_learning():
                p = ap.get_provider()
                bs = p.budget_status() if p else {}
                print(f"[boot] Anthropic API provider installed: model={bs.get('model')} "
                      f"budget=${bs.get('budget_usd_per_day'):.2f}/day "
                      f"spent_today=${bs.get('spent_today_usd', 0):.2f}")
            else:
                print("[boot] Anthropic API provider OFF (set ANTHROPIC_API_KEY to enable)")
        except Exception as e:
            sys.stderr.write(f"[boot] Anthropic install failed: {e}\n")

    # Always-load-latest invariant: ensure the latest snapshot reflects the
    # current working state, OR surface a divergence the user must resolve.
    if _version_script_available():
        sync = boot_sync_check()
        latest = get_version_sync_status().get("latest")
        latest_id = latest["id"] if latest else "(none)"
        action = sync.get("performed") or "—"
        state = sync.get("status") or "?"
        print(f"[boot] version sync: state={state}, latest={latest_id}, action={action}")
        if state in ("working_behind", "divergent", "incomplete_local"):
            print("[boot] ⚠ working dir differs from latest snapshot — review in 📚 Versions")


def main() -> None:
    boot()
    host = "127.0.0.1"
    port = 5173
    url = f"http://{host}:{port}"
    print("\n  ✦ Comply Verify GUI ✦")
    print(f"  → open {url} in browser\n")

    def _open():
        time.sleep(0.6)
        try:
            webbrowser.open(url)
        except Exception:
            pass
    threading.Thread(target=_open, daemon=True).start()

    app.run(host=host, port=port, debug=False, threaded=True)


if __name__ == "__main__":
    main()
