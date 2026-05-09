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
from flask import Flask, abort, jsonify, render_template_string, request, send_file

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

from app import database as db
from app import learning
from app import core


ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "output"

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

app = Flask(__name__)
app.config["JSON_AS_ASCII"] = False


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
    """Iterate annotations skipping ones with broken appearance streams."""
    out = []
    try:
        gen = page.annots()
        if gen is None:
            return out
        # iterate manually so a single bad annot doesn't kill the whole loop
        while True:
            try:
                ann = next(gen)
            except StopIteration:
                break
            except Exception:
                # bad annot — try to continue past it
                continue
            if ann is None:
                continue
            out.append(ann)
    except Exception:
        pass
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
    for ann in parse_inline_annots(doc, page):
        r = [round(c, 2) for c in ann["rect"]]
        sig = (ann["type"], tuple(r), ann["contents"])
        if sig in seen_signatures:
            continue
        out.append({
            "xref": 0,            # 0 = inline (not editable via load_annot)
            "page": page_num,
            "type": ann["type"],
            "rect": r,
            "contents": ann["contents"],
            "_inline": True,
        })
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
    _pages: dict[int, "fitz.Page"] = {}

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
                    a = page.add_freetext_annot(
                        rect,
                        str(edit.get("contents", "")),
                        fontsize=int(edit.get("fontsize", 9)),
                        text_color=(1, 0, 0),
                        align=0,
                    )
                    # Force red bold appearance string
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
    return render_template_string(INDEX_HTML)


@app.route("/api/index")
def api_index():
    rows_payload = []
    for r in ROWS:
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
    if not rel:
        abort(400, "rel required")
    p = (OUTPUT / rel).resolve()
    if not str(p).startswith(str(OUTPUT.resolve())):
        abort(403)
    if not p.exists():
        abort(404)
    png, info = render_pdf_page_png(p, page, dpi=dpi,
                                    highlight=highlight,
                                    no_annots=edit_mode)
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
        return plan

    plan["warnings"].append(f"row role '{role}' — ไม่รองรับการ auto-annotate")
    plan["ok"] = False
    plan["confidence"] = 0.0
    return plan


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

    # Apply PDF edits — Square + FreeText pair
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
    """Snapshot the PDF index + their annotations into dict records that can
    go straight into the DB."""
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
        # Try detect brand/model from filename
        try:
            brand, model = parse_brand_model_from_filename(p.stem)
        except Exception:
            brand, model = ("", "")
        # Annotations
        annots = []
        try:
            annots = list_pdf_annots(p)
        except Exception:
            pass
        try:
            doc = fitz.open(p)
            n_pages = len(doc)
        except Exception:
            n_pages = None
        records.append({
            "rel_path": rel,
            "folder_key": meta.get("folder_key"),
            "section_prefix": meta.get("section_prefix"),
            "size": stat.st_size,
            "mtime": stat.st_mtime,
            "num_pages": n_pages,
            "brand": brand or None,
            "model": model or None,
            "annotations": annots,
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

INDEX_HTML = r"""<!doctype html>
<html lang="th">
<head>
<meta charset="utf-8">
<title>Comply Verify GUI — Smart Plant 1</title>
<style>
/* =====================================================================
 * Comply Verify — Production UI v2
 * Design tokens · components · layout · dark mode · motion
 * Class names referenced by JS are preserved verbatim.
 * ================================================================== */

/* ── Tokens ─────────────────────────────────────────────────────── */
:root {
  /* Brand & semantic palette (light) */
  --c-bg:           #f6f7f9;
  --c-bg-subtle:    #eceef2;
  --c-surface:      #ffffff;
  --c-surface-2:    #f9fafb;
  --c-surface-3:    #f3f4f6;
  --c-overlay:      rgba(15, 23, 42, 0.55);

  --c-border:       #e5e7eb;
  --c-border-strong:#d1d5db;
  --c-divider:      #f1f2f4;

  --c-text:         #0f172a;
  --c-text-muted:   #475569;
  --c-text-soft:    #64748b;
  --c-text-faint:   #94a3b8;
  --c-text-invert:  #ffffff;

  --c-primary:      #2563eb;
  --c-primary-hover:#1d4ed8;
  --c-primary-soft: #dbeafe;
  --c-primary-text: #1e3a8a;

  --c-success:      #10b981;
  --c-success-hover:#059669;
  --c-success-soft: #d1fae5;
  --c-success-text: #065f46;

  --c-warn:         #f59e0b;
  --c-warn-hover:   #d97706;
  --c-warn-soft:    #fef3c7;
  --c-warn-text:    #92400e;

  --c-danger:       #ef4444;
  --c-danger-hover: #dc2626;
  --c-danger-soft:  #fee2e2;
  --c-danger-text:  #991b1b;

  --c-info:         #3b82f6;
  --c-info-soft:    #dbeafe;
  --c-info-text:    #1e40af;

  --c-purple:       #a855f7;
  --c-purple-soft:  #f3e8ff;
  --c-purple-text:  #6b21a8;

  --c-teal:         #14b8a6;
  --c-teal-soft:    #ccfbf1;
  --c-teal-text:    #115e59;

  --c-pink-soft:    #fce7f3;
  --c-pink-text:    #9d174d;

  /* PDF viewer chrome */
  --c-canvas-bg:    #2c2e33;
  --c-canvas-edit:  #3a3c42;

  /* Focus ring */
  --c-focus:        #3b82f6;
  --c-focus-ring:   0 0 0 3px rgba(59, 130, 246, 0.30);

  /* Spacing scale (4px grid) */
  --s-0: 0;
  --s-1: 2px;
  --s-2: 4px;
  --s-3: 6px;
  --s-4: 8px;
  --s-5: 10px;
  --s-6: 12px;
  --s-7: 16px;
  --s-8: 20px;
  --s-9: 24px;
  --s-10: 32px;
  --s-12: 48px;

  /* Radius */
  --r-sm: 4px;
  --r-md: 6px;
  --r-lg: 8px;
  --r-xl: 12px;
  --r-pill: 999px;

  /* Elevation */
  --e-0: none;
  --e-1: 0 1px 2px rgba(15, 23, 42, 0.06), 0 1px 1px rgba(15, 23, 42, 0.04);
  --e-2: 0 2px 6px rgba(15, 23, 42, 0.08), 0 1px 2px rgba(15, 23, 42, 0.04);
  --e-3: 0 6px 16px rgba(15, 23, 42, 0.10), 0 1px 3px rgba(15, 23, 42, 0.06);
  --e-4: 0 12px 28px rgba(15, 23, 42, 0.14), 0 2px 6px rgba(15, 23, 42, 0.06);
  --e-5: 0 24px 48px rgba(15, 23, 42, 0.20), 0 4px 12px rgba(15, 23, 42, 0.08);

  /* Type */
  --f-sans: ui-sans-serif, -apple-system, BlinkMacSystemFont, "Inter", "Segoe UI",
            "Sukhumvit Set", "Noto Sans Thai", "Thonburi", system-ui, sans-serif;
  --f-mono: ui-monospace, "SF Mono", Menlo, Consolas, "Roboto Mono", monospace;

  --t-xs:   10px;
  --t-sm:   11px;
  --t-base: 12px;
  --t-md:   13px;
  --t-lg:   14px;
  --t-xl:   16px;
  --t-2xl:  20px;
  --t-3xl:  24px;

  --lh-tight:  1.25;
  --lh-normal: 1.45;
  --lh-loose:  1.6;

  /* Motion */
  --ease-out:  cubic-bezier(0.16, 1, 0.3, 1);
  --ease-in:   cubic-bezier(0.4, 0, 1, 1);
  --ease-std:  cubic-bezier(0.4, 0, 0.2, 1);

  --d-fast:    120ms;
  --d-base:    180ms;
  --d-slow:    260ms;

  /* Layout */
  --topbar-h:    44px;
  --pane-head-h: 36px;
  --tree-w:      300px;
  --tree-w-md:   260px;
  --action-bar-w: 720px;

  color-scheme: light;
}

/* Dark mode — auto + manual */
@media (prefers-color-scheme: dark) {
  :root:not([data-theme="light"]) {
    --c-bg:           #0b0d12;
    --c-bg-subtle:    #11141b;
    --c-surface:      #161a22;
    --c-surface-2:    #1c2029;
    --c-surface-3:    #232834;
    --c-overlay:      rgba(0, 0, 0, 0.7);

    --c-border:       #2a2f3a;
    --c-border-strong:#3a4150;
    --c-divider:      #1f2430;

    --c-text:         #e8ecf3;
    --c-text-muted:   #b6bdca;
    --c-text-soft:    #8a93a4;
    --c-text-faint:   #5a6373;

    --c-primary:      #60a5fa;
    --c-primary-hover:#3b82f6;
    --c-primary-soft: rgba(59, 130, 246, 0.18);
    --c-primary-text: #bfdbfe;

    --c-success-soft: rgba(16, 185, 129, 0.18);
    --c-success-text: #6ee7b7;

    --c-warn-soft:    rgba(245, 158, 11, 0.18);
    --c-warn-text:    #fcd34d;

    --c-danger-soft:  rgba(239, 68, 68, 0.18);
    --c-danger-text:  #fca5a5;

    --c-info-soft:    rgba(59, 130, 246, 0.18);
    --c-info-text:    #bfdbfe;

    --c-purple-soft:  rgba(168, 85, 247, 0.18);
    --c-purple-text:  #d8b4fe;

    --c-teal-soft:    rgba(20, 184, 166, 0.18);
    --c-teal-text:    #5eead4;

    --c-pink-soft:    rgba(236, 72, 153, 0.18);
    --c-pink-text:    #fbcfe8;

    --c-canvas-bg:    #06080c;
    --c-canvas-edit:  #0e1118;

    --e-1: 0 1px 2px rgba(0, 0, 0, 0.40);
    --e-2: 0 2px 6px rgba(0, 0, 0, 0.45);
    --e-3: 0 6px 16px rgba(0, 0, 0, 0.55);
    --e-4: 0 12px 28px rgba(0, 0, 0, 0.65);
    --e-5: 0 24px 48px rgba(0, 0, 0, 0.75);

    color-scheme: dark;
  }
}
:root[data-theme="dark"] {
  --c-bg:           #0b0d12;
  --c-bg-subtle:    #11141b;
  --c-surface:      #161a22;
  --c-surface-2:    #1c2029;
  --c-surface-3:    #232834;
  --c-overlay:      rgba(0, 0, 0, 0.7);
  --c-border:       #2a2f3a;
  --c-border-strong:#3a4150;
  --c-divider:      #1f2430;
  --c-text:         #e8ecf3;
  --c-text-muted:   #b6bdca;
  --c-text-soft:    #8a93a4;
  --c-text-faint:   #5a6373;
  --c-primary:      #60a5fa;
  --c-primary-hover:#3b82f6;
  --c-primary-soft: rgba(59, 130, 246, 0.18);
  --c-primary-text: #bfdbfe;
  --c-success-soft: rgba(16, 185, 129, 0.18);
  --c-success-text: #6ee7b7;
  --c-warn-soft:    rgba(245, 158, 11, 0.18);
  --c-warn-text:    #fcd34d;
  --c-danger-soft:  rgba(239, 68, 68, 0.18);
  --c-danger-text:  #fca5a5;
  --c-info-soft:    rgba(59, 130, 246, 0.18);
  --c-info-text:    #bfdbfe;
  --c-purple-soft:  rgba(168, 85, 247, 0.18);
  --c-purple-text:  #d8b4fe;
  --c-teal-soft:    rgba(20, 184, 166, 0.18);
  --c-teal-text:    #5eead4;
  --c-pink-soft:    rgba(236, 72, 153, 0.18);
  --c-pink-text:    #fbcfe8;
  --c-canvas-bg:    #06080c;
  --c-canvas-edit:  #0e1118;
  --e-1: 0 1px 2px rgba(0, 0, 0, 0.40);
  --e-2: 0 2px 6px rgba(0, 0, 0, 0.45);
  --e-3: 0 6px 16px rgba(0, 0, 0, 0.55);
  --e-4: 0 12px 28px rgba(0, 0, 0, 0.65);
  --e-5: 0 24px 48px rgba(0, 0, 0, 0.75);
  color-scheme: dark;
}

/* ── Reset ──────────────────────────────────────────────────────── */
*, *::before, *::after { box-sizing: border-box; }
html, body { margin: 0; padding: 0; height: 100%; }
html { -webkit-font-smoothing: antialiased; -moz-osx-font-smoothing: grayscale; text-size-adjust: 100%; }
body {
  font-family: var(--f-sans);
  font-size: var(--t-md);
  line-height: var(--lh-normal);
  color: var(--c-text);
  background: var(--c-bg);
  overflow: hidden;
}
button, input, select, textarea { font-family: inherit; color: inherit; }
button { cursor: pointer; }
button:disabled { cursor: not-allowed; }
img { display: block; max-width: 100%; }
a { color: var(--c-primary); text-decoration: none; }
a:hover { text-decoration: underline; }
::selection { background: var(--c-primary-soft); color: var(--c-primary-text); }

/* Custom scrollbars (subtle, persistent) */
@supports (scrollbar-width: thin) {
  * { scrollbar-width: thin; scrollbar-color: var(--c-border-strong) transparent; }
}
::-webkit-scrollbar { width: 10px; height: 10px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb {
  background: var(--c-border-strong);
  border: 2px solid transparent; background-clip: padding-box;
  border-radius: 999px;
}
::-webkit-scrollbar-thumb:hover { background: var(--c-text-faint); background-clip: padding-box; border: 2px solid transparent; }

/* Focus-visible ring (universal) */
:focus { outline: none; }
:focus-visible {
  outline: 2px solid var(--c-focus);
  outline-offset: 2px;
  border-radius: var(--r-sm);
}
button:focus-visible, .btn:focus-visible, .ab-btn:focus-visible,
.versions-btn:focus-visible, .canvas-toolbar button:focus-visible,
.edit-toolbar button:focus-visible, .edit-toggle-btn:focus-visible,
.col-d-menu button:focus-visible, .mobile-tabs button:focus-visible {
  outline: 2px solid var(--c-focus);
  outline-offset: 2px;
}
input:focus-visible, select:focus-visible, textarea:focus-visible {
  outline: none; border-color: var(--c-primary);
  box-shadow: var(--c-focus-ring);
}

/* Visually-hidden helper for SR-only labels */
.sr-only {
  position: absolute; width: 1px; height: 1px; padding: 0; margin: -1px;
  overflow: hidden; clip: rect(0, 0, 0, 0); white-space: nowrap; border: 0;
}

/* ── App shell ──────────────────────────────────────────────────── */
#app {
  display: grid;
  grid-template-rows: var(--topbar-h) 1fr;
  grid-template-columns: minmax(var(--tree-w-md), var(--tree-w)) 1fr 1fr;
  grid-template-areas:
    "topbar topbar topbar"
    "tree   center pdf";
  height: 100vh;
  background: var(--c-bg);
}
#app.has-sync-banner { padding-top: 32px; }

.topbar {
  grid-area: topbar;
  display: flex; align-items: center; gap: var(--s-6);
  padding: 0 var(--s-7);
  background: var(--c-surface);
  border-bottom: 1px solid var(--c-border);
  z-index: 10;
}
.topbar-brand {
  display: flex; align-items: center; gap: var(--s-4);
  font-weight: 700; font-size: var(--t-lg); letter-spacing: -0.01em;
  color: var(--c-text);
}
.topbar-brand .logo {
  width: 24px; height: 24px; border-radius: var(--r-md);
  background: linear-gradient(135deg, var(--c-primary), var(--c-purple));
  display: inline-flex; align-items: center; justify-content: center;
  color: white; font-weight: 800; font-size: 12px;
  box-shadow: var(--e-1);
}
.topbar-brand .sub { font-weight: 400; color: var(--c-text-soft); font-size: var(--t-base); }
.topbar-spacer { flex: 1; }
.topbar-actions { display: flex; gap: var(--s-3); align-items: center; }
.topbar-actions .stats-pill {
  display: inline-flex; align-items: center; gap: var(--s-3);
  padding: 4px var(--s-5);
  background: var(--c-surface-2); border: 1px solid var(--c-border);
  border-radius: var(--r-pill);
  font-size: var(--t-sm); color: var(--c-text-muted);
}
.topbar-actions .stats-pill strong { color: var(--c-text); font-weight: 700; }
.topbar-actions .stats-pill .sep { color: var(--c-text-faint); }

.pane {
  overflow: hidden;
  display: flex; flex-direction: column;
  background: var(--c-surface);
  border-right: 1px solid var(--c-border);
  min-width: 0; min-height: 0;
}
.pane:last-child { border-right: none; }
.tree-pane   { grid-area: tree; }
.center-pane { grid-area: center; }
.pdf-pane    { grid-area: pdf; }

/* Pane headers */
.pane > h2 {
  margin: 0;
  padding: 0 var(--s-7);
  height: var(--pane-head-h);
  flex-shrink: 0;
  font-size: var(--t-sm); font-weight: 600;
  background: var(--c-surface);
  border-bottom: 1px solid var(--c-border);
  color: var(--c-text);
  display: flex; justify-content: space-between; align-items: center;
  letter-spacing: 0.02em;
  gap: var(--s-4);
}
.pane > h2 .pane-title { display: inline-flex; align-items: center; gap: var(--s-3); }
.pane > h2 .pane-title .icon { font-size: 14px; }

/* ── Buttons (canonical) ────────────────────────────────────────── */
.btn {
  display: inline-flex; align-items: center; justify-content: center;
  gap: var(--s-3);
  padding: 6px 12px;
  border: 1px solid var(--c-border-strong);
  background: var(--c-surface);
  color: var(--c-text);
  border-radius: var(--r-md);
  font-size: var(--t-base);
  font-weight: 500;
  line-height: 1.4;
  transition: background var(--d-fast) var(--ease-std),
              border-color var(--d-fast) var(--ease-std),
              color var(--d-fast) var(--ease-std),
              transform var(--d-fast) var(--ease-out),
              box-shadow var(--d-fast) var(--ease-std);
  white-space: nowrap;
}
.btn:hover { background: var(--c-surface-2); border-color: var(--c-text-faint); }
.btn:active { transform: translateY(1px); }
.btn:disabled, .btn.is-disabled {
  opacity: 0.5; pointer-events: none;
}
.btn.btn-primary {
  background: var(--c-primary); color: var(--c-text-invert); border-color: var(--c-primary);
}
.btn.btn-primary:hover { background: var(--c-primary-hover); border-color: var(--c-primary-hover); }
.btn.btn-success, .btn.save-btn {
  background: var(--c-success); color: white; border-color: var(--c-success);
  font-weight: 600;
}
.btn.btn-success:hover, .btn.save-btn:hover:not(:disabled) {
  background: var(--c-success-hover); border-color: var(--c-success-hover);
}
.btn.save-btn:disabled { background: var(--c-border-strong); color: var(--c-text-soft); border-color: var(--c-border-strong); }
.btn.btn-ghost { background: transparent; border-color: transparent; color: var(--c-text-muted); }
.btn.btn-ghost:hover { background: var(--c-surface-2); color: var(--c-text); }
.btn.btn-danger { color: var(--c-danger-text); border-color: var(--c-danger); background: var(--c-surface); }
.btn.btn-danger:hover { background: var(--c-danger-soft); }

/* Toolbar buttons (compact) */
.canvas-toolbar {
  padding: var(--s-3) var(--s-7);
  background: var(--c-surface);
  border-bottom: 1px solid var(--c-border);
  display: flex; gap: var(--s-3); align-items: center;
  font-size: var(--t-base); flex-wrap: wrap;
  min-height: 36px;
}
.canvas-toolbar button {
  display: inline-flex; align-items: center; justify-content: center;
  min-width: 26px; height: 26px;
  padding: 0 var(--s-4);
  border: 1px solid var(--c-border);
  background: var(--c-surface);
  color: var(--c-text);
  border-radius: var(--r-sm);
  font-size: var(--t-sm);
  transition: background var(--d-fast) var(--ease-std), border-color var(--d-fast) var(--ease-std);
}
.canvas-toolbar button:hover { background: var(--c-surface-2); border-color: var(--c-border-strong); }
.canvas-toolbar .info { color: var(--c-text-soft); font-size: var(--t-sm); }
.canvas-toolbar label {
  display: inline-flex; align-items: center; gap: 4px;
  font-size: var(--t-sm); color: var(--c-text-muted);
  cursor: pointer; user-select: none;
}

/* ── Inputs ─────────────────────────────────────────────────────── */
input[type="text"], input[type="search"], select, textarea {
  font-family: inherit; font-size: var(--t-base);
  padding: 6px var(--s-5);
  border: 1px solid var(--c-border-strong);
  background: var(--c-surface);
  color: var(--c-text);
  border-radius: var(--r-md);
  transition: border-color var(--d-fast) var(--ease-std), box-shadow var(--d-fast) var(--ease-std);
  width: 100%;
}
input[type="checkbox"] {
  accent-color: var(--c-primary);
  cursor: pointer;
}
input::placeholder, textarea::placeholder { color: var(--c-text-faint); }
select { cursor: pointer; }

/* ── kbd ────────────────────────────────────────────────────────── */
.kbd {
  display: inline-block;
  padding: 0 5px; min-width: 14px;
  font-family: var(--f-mono); font-size: var(--t-xs);
  background: var(--c-surface-3); color: var(--c-text-muted);
  border: 1px solid var(--c-border);
  border-bottom-width: 2px;
  border-radius: var(--r-sm);
  line-height: 16px; text-align: center;
}

/* ── Mobile-tab nav ─────────────────────────────────────────────── */
.mobile-tabs {
  display: none;
  background: var(--c-surface); border-bottom: 1px solid var(--c-border);
  padding: 4px;
  grid-area: topbar;
}
.mobile-tabs button {
  flex: 1; padding: 8px 4px; font-size: var(--t-base); font-weight: 600;
  border: 0; background: transparent; cursor: pointer;
  color: var(--c-text-soft);
  border-radius: var(--r-md);
  transition: background var(--d-fast) var(--ease-std), color var(--d-fast) var(--ease-std);
}
.mobile-tabs button:hover { background: var(--c-surface-2); color: var(--c-text); }
.mobile-tabs button.active {
  background: var(--c-primary-soft); color: var(--c-primary-text);
}

/* ── Stats bar (top of tree) ────────────────────────────────────── */
.stats-bar {
  padding: var(--s-3) var(--s-7);
  background: var(--c-surface-2);
  font-size: var(--t-sm);
  color: var(--c-text-muted);
  display: flex; gap: var(--s-6);
  border-bottom: 1px solid var(--c-border);
  flex-wrap: wrap;
}
.stats-bar strong { color: var(--c-text); font-weight: 700; }

/* ── Tree pane ──────────────────────────────────────────────────── */
.tree-pane .toolbar {
  padding: var(--s-3) var(--s-5);
  border-bottom: 1px solid var(--c-border);
  background: var(--c-surface);
  display: flex; flex-direction: column; gap: var(--s-3);
}
.tree-pane .toolbar input,
.tree-pane .toolbar select {
  font-size: var(--t-base);
  padding: 5px var(--s-4);
}
.tree-pane .toolbar input[type="text"]::-webkit-search-cancel-button { cursor: pointer; }

.filter-row { display: flex; gap: var(--s-3); align-items: center; }
.filter-row select { font-size: var(--t-sm); padding: 4px var(--s-3); flex: 1; min-width: 0; }
.filter-row label {
  display: inline-flex; align-items: center; gap: 4px;
  font-size: var(--t-sm); color: var(--c-text-muted); flex: 1; cursor: pointer;
  user-select: none;
}
.filter-row .mini-btn {
  font-size: var(--t-sm); padding: 3px var(--s-4);
  background: var(--c-surface); border: 1px solid var(--c-border);
  border-radius: var(--r-sm); color: var(--c-text-muted);
  transition: background var(--d-fast) var(--ease-std), color var(--d-fast) var(--ease-std);
}
.filter-row .mini-btn:hover { background: var(--c-surface-2); color: var(--c-text); }

.tree-scroll { flex: 1; overflow: auto; padding: var(--s-3) 0; font-size: var(--t-base); }
.tree-node { user-select: none; }
.tree-row {
  display: flex; align-items: center;
  padding: 3px var(--s-3) 3px 0;
  cursor: pointer; gap: var(--s-1);
  line-height: 1.4; min-height: 24px;
  position: relative;
  transition: background var(--d-fast) var(--ease-std);
}
.tree-row:hover { background: var(--c-surface-2); }
.tree-row.selected {
  background: var(--c-primary-soft);
  box-shadow: inset 2px 0 0 var(--c-primary);
}
.tree-row.selected .tree-label { color: var(--c-primary-text); font-weight: 600; }
.tree-chev {
  width: 16px; flex-shrink: 0; text-align: center; color: var(--c-text-faint);
  font-size: 9px;
  transition: transform var(--d-fast) var(--ease-out);
}
.tree-chev.empty { visibility: hidden; }
.tree-node.expanded > .tree-row .tree-chev:not(.empty) { transform: rotate(0deg); }
.tree-icon { width: 16px; flex-shrink: 0; font-size: 12px; text-align: center; color: var(--c-text-faint); }
.tree-icon.section { color: var(--c-primary); }
.tree-icon.item    { color: var(--c-warn); }
.tree-icon.sub     { color: var(--c-success); }
.tree-label {
  flex: 1; min-width: 0;
  overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
  font-size: var(--t-base); color: var(--c-text);
}
.tree-status { width: 14px; flex-shrink: 0; font-size: var(--t-sm); text-align: center; }
.tree-flags  { font-size: var(--t-xs); color: var(--c-warn-text); }
.tree-children { display: none; }
.tree-node.expanded > .tree-children { display: block; }
.tree-row.match { background: var(--c-warn-soft); }
.tree-row.path-match .tree-label { font-weight: 600; }

/* Confidence dots */
.tree-row .conf-dot {
  width: 8px; height: 8px; border-radius: 50%;
  display: inline-block; flex-shrink: 0; margin-left: 2px;
  box-shadow: 0 0 0 2px var(--c-surface);
}
.conf-dot.high { background: var(--c-success); }
.conf-dot.med  { background: var(--c-warn); }
.conf-dot.low  { background: var(--c-danger); }
.conf-dot.none { background: var(--c-border-strong); }

/* ── Center pane (TOR / xlsx) ───────────────────────────────────── */
.center-pane { display: flex; flex-direction: column; }
.split-top, .split-bot { display: flex; flex-direction: column; min-height: 0; overflow: hidden; }
.split-top { flex: 1 1 50%; }
.split-bot { flex: 1 1 50%; border-top: 1px solid var(--c-border); position: relative; }
.split-bot::before {
  content: ''; position: absolute; top: -2px; left: 0; right: 0; height: 4px;
  cursor: ns-resize; z-index: 2;
  background: transparent;
  transition: background var(--d-base) var(--ease-std);
}
.split-bot:hover::before { background: var(--c-primary-soft); }

.tor-canvas, .pdf-canvas {
  flex: 1; overflow: auto;
  background: var(--c-canvas-bg);
  padding: var(--s-7);
  text-align: center;
  scroll-behavior: smooth;
}
.tor-canvas img, .pdf-canvas img {
  max-width: 100%;
  box-shadow: 0 4px 16px rgba(0,0,0,0.45);
  background: white;
  border-radius: 2px;
}
.empty-canvas {
  color: var(--c-text-faint);
  padding: var(--s-12) var(--s-7);
  text-align: center;
  font-size: var(--t-md);
  display: flex; flex-direction: column; align-items: center; gap: var(--s-4);
}
.empty-canvas::before {
  content: '◌';
  font-size: 32px; opacity: 0.5;
}
.empty-canvas.loading::before {
  content: '';
  width: 28px; height: 28px;
  border: 3px solid var(--c-border);
  border-top-color: var(--c-primary);
  border-radius: 50%;
  animation: spin 0.8s linear infinite;
}

/* ── XLSX preview ───────────────────────────────────────────────── */
.xlsx-wrap { flex: 1; overflow: auto; background: var(--c-bg); }
.xlsx-table { width: 100%; border-collapse: separate; border-spacing: 0; font-size: var(--t-sm); }
.xlsx-table thead th {
  position: sticky; top: 0; z-index: 1;
  background: var(--c-surface-3);
  color: var(--c-text);
  padding: var(--s-3) var(--s-5);
  border-bottom: 1px solid var(--c-border-strong);
  border-right: 1px solid var(--c-border);
  font-size: var(--t-xs);
  font-weight: 700;
  text-align: left; white-space: nowrap;
  letter-spacing: 0.04em; text-transform: uppercase;
  color: var(--c-text-muted);
}
.xlsx-table thead th:last-child { border-right: none; }
.xlsx-table td {
  border-bottom: 1px solid var(--c-divider);
  border-right: 1px solid var(--c-divider);
  padding: var(--s-4) var(--s-5);
  vertical-align: top; white-space: pre-wrap; word-break: break-word;
  background: var(--c-surface);
  transition: background var(--d-fast) var(--ease-std);
}
.xlsx-table td:last-child { border-right: none; }
.xlsx-table .col-A { width: 56px; font-family: var(--f-mono); color: var(--c-text-soft); }
.xlsx-table .col-B { width: 28%; }
.xlsx-table .col-C { width: 28%; }
.xlsx-table .col-D { width: 24%; }
.xlsx-table .col-E { width: 80px; font-size: var(--t-xs); }
.xlsx-table .col-F { width: 110px; font-size: var(--t-xs); color: var(--c-text-soft); }
.xlsx-table .row-num {
  width: 42px; font-family: var(--f-mono); font-size: var(--t-xs);
  color: var(--c-text-faint); background: var(--c-surface-2);
}
.xlsx-table tr:hover td { background: var(--c-surface-2); }
.xlsx-table tr.target td {
  background: var(--c-warn-soft) !important;
  font-weight: 500;
}
.xlsx-table tr.target .col-A,
.xlsx-table tr.target .row-num { color: var(--c-text); font-weight: 700; }
.xlsx-table tr.target:hover td { background: #fde68a !important; }

.xlsx-table td.col-D.editable {
  cursor: pointer; position: relative; padding-right: 22px !important;
}
.xlsx-table td.col-D.editable:hover { background: var(--c-purple-soft); }
.xlsx-table td.col-D.editable:hover .d-caret { opacity: 1; transform: translateY(-50%); }
.xlsx-table td.col-D.editable .d-caret {
  position: absolute; right: 6px; top: 50%; transform: translateY(-50%) translateX(2px);
  font-size: var(--t-xs); color: var(--c-text-faint); opacity: 0.5;
  transition: opacity var(--d-fast) var(--ease-std), transform var(--d-fast) var(--ease-out);
}
.xlsx-table td.col-D.commitment .d-text { color: var(--c-text-soft); font-style: italic; }
.xlsx-table td.col-D.commitment::before {
  content: '⚠ '; color: var(--c-warn); font-style: normal;
}
.xlsx-table td.col-D.editing {
  background: var(--c-purple-soft) !important;
  outline: 2px solid var(--c-purple);
  outline-offset: -2px;
}
.xlsx-table td.col-D.editing .d-caret { display: none; }

/* Vendor tags */
.vendor-tag {
  display: inline-flex; align-items: center;
  padding: 1px var(--s-4);
  border-radius: var(--r-sm);
  font-size: var(--t-xs); font-weight: 700;
  letter-spacing: 0.04em;
}
.vendor-tag.SMART { background: var(--c-primary-soft); color: var(--c-primary-text); }
.vendor-tag.TRIO  { background: var(--c-success-soft); color: var(--c-success-text); }
.vendor-tag.SR    { background: var(--c-pink-soft); color: var(--c-pink-text); }

/* ── Col D dropdown menu ────────────────────────────────────────── */
.col-d-menu {
  position: fixed; z-index: 250;
  background: var(--c-surface);
  border-radius: var(--r-lg);
  box-shadow: var(--e-4);
  border: 1px solid var(--c-border);
  padding: 5px; min-width: 260px;
  font-size: var(--t-base);
  animation: menu-pop var(--d-fast) var(--ease-out);
}
@keyframes menu-pop {
  from { opacity: 0; transform: scale(0.96) translateY(-4px); }
  to   { opacity: 1; transform: scale(1) translateY(0); }
}
.col-d-menu .menu-header {
  padding: var(--s-4) var(--s-5) var(--s-3);
  font-size: var(--t-xs);
  color: var(--c-text-faint);
  text-transform: uppercase; letter-spacing: 0.06em; font-weight: 700;
  border-bottom: 1px solid var(--c-divider);
  margin-bottom: var(--s-2);
}
.col-d-menu button {
  display: flex; align-items: center; gap: var(--s-4);
  width: 100%; padding: 7px var(--s-5);
  border: 0; background: transparent;
  border-radius: var(--r-sm);
  text-align: left;
  font-size: var(--t-base);
  color: var(--c-text);
  transition: background var(--d-fast) var(--ease-std);
}
.col-d-menu button:hover { background: var(--c-surface-2); }
.col-d-menu button .icon { font-size: 14px; width: 18px; text-align: center; }
.col-d-menu button .label { flex: 1; }
.col-d-menu button .hint { font-size: var(--t-xs); color: var(--c-text-faint); }
.col-d-menu button.danger { color: var(--c-danger-text); }
.col-d-menu button.danger:hover { background: var(--c-danger-soft); }
.col-d-menu button.primary { color: var(--c-success-text); font-weight: 600; }
.col-d-menu button.primary:hover { background: var(--c-success-soft); }
.col-d-menu .sep { height: 1px; background: var(--c-divider); margin: var(--s-2) 0; }

/* ── Catalog pane ───────────────────────────────────────────────── */
.pdf-pane > h2 { gap: var(--s-4); }
.pdf-pane > h2 .filename {
  font-weight: 400;
  text-transform: none;
  letter-spacing: 0;
  color: var(--c-text-soft);
  font-size: var(--t-sm);
  font-family: var(--f-mono);
  max-width: 320px;
  overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
  flex: 1;
}
.edit-toggle-btn {
  font-size: var(--t-sm); padding: 4px 10px;
  background: var(--c-surface); border: 1px solid var(--c-border-strong);
  color: var(--c-text-muted);
  border-radius: var(--r-md);
  text-transform: none; letter-spacing: 0; font-weight: 600;
  display: inline-flex; align-items: center; gap: 4px;
  transition: all var(--d-fast) var(--ease-std);
}
.edit-toggle-btn:hover { background: var(--c-surface-2); color: var(--c-text); }
.edit-toggle-btn.active {
  background: var(--c-warn); color: white; border-color: var(--c-warn);
  box-shadow: var(--e-2);
}

/* Edit toolbar */
.edit-toolbar {
  padding: var(--s-3) var(--s-7);
  background: linear-gradient(180deg, var(--c-warn-soft) 0%, transparent 200%);
  border-bottom: 1px solid var(--c-warn);
  display: flex; gap: var(--s-2);
  align-items: center; flex-wrap: wrap;
  min-height: 36px;
}
.edit-toolbar button {
  display: inline-flex; align-items: center; justify-content: center;
  font-size: var(--t-md); padding: 4px var(--s-4);
  border: 1px solid var(--c-border-strong);
  background: var(--c-surface);
  color: var(--c-text);
  border-radius: var(--r-md);
  min-width: 30px; height: 28px;
  transition: all var(--d-fast) var(--ease-std);
}
.edit-toolbar button:hover:not(:disabled) {
  background: var(--c-surface-2); border-color: var(--c-warn);
}
.edit-toolbar button:disabled { opacity: 0.4; }
.edit-toolbar button.tool.active {
  background: var(--c-warn); color: white; border-color: var(--c-warn);
  box-shadow: var(--e-1);
}
.edit-toolbar .sep {
  width: 1px; background: var(--c-border-strong); height: 18px; margin: 0 var(--s-3);
}
.edit-toolbar .save-btn {
  background: var(--c-success); color: white; border-color: var(--c-success);
  font-weight: 600;
}
.edit-toolbar .save-btn:disabled {
  background: var(--c-border-strong); color: var(--c-text-soft);
  border-color: var(--c-border-strong);
}
.edit-toolbar .dirty-indicator {
  color: var(--c-warn-hover); font-size: var(--t-sm); font-weight: 600;
  padding: 0 var(--s-3);
}

/* PDF canvas wrapper for SVG overlay */
.pdf-page-host { position: relative; display: inline-block; }
.pdf-page-host img.pdf-page-img {
  max-width: 100%;
  box-shadow: 0 4px 16px rgba(0,0,0,0.5);
  display: block; background: white;
  border-radius: 2px;
}
.pdf-overlay { position: absolute; top: 0; left: 0; width: 100%; height: 100%; pointer-events: none; }
.pdf-overlay.editable { pointer-events: all; }
.pdf-overlay g.annot { pointer-events: all; cursor: pointer; }
.pdf-overlay g.annot.selected { cursor: move; }
.pdf-overlay rect.ann-rect { fill: transparent; stroke: rgba(255,0,0,0.85); stroke-width: 1; }
.pdf-overlay rect.ann-rect.freetext { stroke: rgba(255,0,0,0.4); stroke-dasharray: 3 2; }
.pdf-overlay g.annot.selected rect.ann-rect { stroke: var(--c-warn); stroke-width: 2; }
.pdf-overlay text.ann-text { fill: red; font-weight: bold; user-select: none; pointer-events: none; }
.pdf-overlay rect.ann-hit { fill: rgba(0,0,0,0); stroke: none; }
.pdf-overlay rect.ann-hit:hover { fill: rgba(245,158,11,0.18); }
.pdf-overlay g.annot.selected rect.ann-hit { fill: rgba(245,158,11,0.10); }
.pdf-overlay .handle { fill: white; stroke: var(--c-warn); stroke-width: 1.5; cursor: nwse-resize; }
.pdf-overlay .handle.h-n, .pdf-overlay .handle.h-s { cursor: ns-resize; }
.pdf-overlay .handle.h-e, .pdf-overlay .handle.h-w { cursor: ew-resize; }
.pdf-overlay .handle.h-ne, .pdf-overlay .handle.h-sw { cursor: nesw-resize; }
.pdf-overlay rect.draw-preview {
  fill: rgba(245,158,11,0.20); stroke: var(--c-warn);
  stroke-width: 2; stroke-dasharray: 4 2; pointer-events: none;
}
.pdf-canvas.edit-mode { background: var(--c-canvas-edit); }
.pdf-canvas.edit-mode.tool-drawRect { cursor: crosshair; }
.pdf-canvas.edit-mode.tool-addText { cursor: text; }

/* Inline text editor (for FreeText) */
.text-editor {
  position: absolute; box-sizing: border-box;
  background: white; color: red; font-weight: bold;
  border: 2px solid var(--c-warn); border-radius: 2px;
  padding: 2px 4px; outline: none; resize: none;
  z-index: 50; font-family: inherit;
  box-shadow: var(--e-3);
}

.pdf-annots {
  max-height: 140px; overflow: auto;
  background: var(--c-surface-2);
  border-top: 1px solid var(--c-border);
  padding: var(--s-3) var(--s-7);
  font-size: var(--t-xs);
  font-family: var(--f-mono);
  color: var(--c-text-muted);
}
.pdf-annots .ann-row { padding: 1px 0; }
.pdf-annots .ann-row.matched {
  background: var(--c-warn-soft);
  font-weight: 600;
  padding-left: var(--s-3);
  border-left: 2px solid var(--c-warn);
  color: var(--c-warn-text);
  border-radius: 0 var(--r-sm) var(--r-sm) 0;
}

/* Manual-annotate banner */
.manual-mode-banner {
  display: none;
  padding: var(--s-3) var(--s-7);
  background: linear-gradient(90deg, var(--c-success-soft), rgba(167,243,208,0.5));
  border-bottom: 2px solid var(--c-success);
  font-size: var(--t-base);
  align-items: center; gap: var(--s-4);
}
.manual-mode-banner.show { display: flex; animation: slide-down var(--d-base) var(--ease-out); }
@keyframes slide-down {
  from { opacity: 0; transform: translateY(-6px); }
  to   { opacity: 1; transform: translateY(0); }
}
.manual-mode-banner .target-info { flex: 1; color: var(--c-success-text); }
.manual-mode-banner .target-info strong { font-weight: 700; }
.manual-mode-banner .target-info code {
  background: rgba(255,255,255,0.7);
  padding: 1px 4px; border-radius: var(--r-sm);
  font-size: var(--t-sm); font-family: var(--f-mono);
}
.manual-mode-banner button {
  padding: 5px var(--s-5);
  border: 1px solid var(--c-success); background: var(--c-surface);
  color: var(--c-success-text);
  border-radius: var(--r-md); cursor: pointer;
  font-size: var(--t-sm); font-weight: 600;
  transition: all var(--d-fast) var(--ease-std);
}
.manual-mode-banner button:hover { background: var(--c-success-soft); }
.manual-mode-banner button.save {
  background: var(--c-success); color: white; border-color: var(--c-success);
}
.manual-mode-banner button.save:hover:not(:disabled) {
  background: var(--c-success-hover); border-color: var(--c-success-hover);
}
.manual-mode-banner button.save:disabled {
  background: var(--c-border-strong); border-color: var(--c-border-strong);
  color: var(--c-text-soft); cursor: not-allowed;
}
.manual-mode-banner button.cancel {
  color: var(--c-danger-text); border-color: var(--c-danger);
}
.manual-mode-banner button.cancel:hover { background: var(--c-danger-soft); }

/* ── Floating action bar ────────────────────────────────────────── */
.action-bar {
  position: fixed; bottom: var(--s-7); left: 50%; transform: translateX(-50%);
  background: var(--c-surface);
  border-radius: var(--r-xl);
  box-shadow: var(--e-4);
  padding: var(--s-5) var(--s-7);
  min-width: var(--action-bar-w); max-width: 92vw;
  z-index: 100;
  border: 1px solid var(--c-border);
  backdrop-filter: blur(8px);
  animation: ab-rise var(--d-slow) var(--ease-out);
}
@keyframes ab-rise {
  from { opacity: 0; transform: translate(-50%, 12px); }
  to   { opacity: 1; transform: translate(-50%, 0); }
}
.ab-top { display: flex; gap: var(--s-6); align-items: center; flex-wrap: wrap; }
.ab-row-info {
  font-weight: 600; font-size: var(--t-md); color: var(--c-text);
  display: inline-flex; align-items: center; gap: var(--s-3);
}
.ab-row-info .row-num {
  font-family: var(--f-mono);
  background: var(--c-primary-soft);
  color: var(--c-primary-text);
  padding: 2px var(--s-4);
  border-radius: var(--r-sm);
  font-size: var(--t-base); font-weight: 700;
}
.ab-row-info .section { color: var(--c-text-soft); font-weight: 400; font-size: var(--t-sm); }
.ab-flags {
  font-size: var(--t-sm); color: var(--c-warn-hover);
  background: var(--c-warn-soft); padding: 2px var(--s-4);
  border-radius: var(--r-sm);
}
.ab-spacer { flex: 1; }
.ab-buttons { display: flex; gap: 4px; align-items: center; }
.ab-btn {
  display: inline-flex; align-items: center; gap: var(--s-3);
  padding: 7px var(--s-6);
  border: 1px solid var(--c-border-strong);
  background: var(--c-surface);
  color: var(--c-text);
  border-radius: var(--r-md);
  font-size: var(--t-base); font-weight: 500;
  transition: background var(--d-fast) var(--ease-std),
              color var(--d-fast) var(--ease-std),
              border-color var(--d-fast) var(--ease-std),
              transform var(--d-fast) var(--ease-out);
  white-space: nowrap;
}
.ab-btn:hover { background: var(--c-surface-2); }
.ab-btn:active { transform: translateY(1px); }
.ab-btn.pass { border-color: var(--c-success); color: var(--c-success-text); }
.ab-btn.pass:hover { background: var(--c-success-soft); }
.ab-btn.pass.active {
  background: var(--c-success); color: white; border-color: var(--c-success);
  box-shadow: var(--e-2);
}
.ab-btn.fail { border-color: var(--c-danger); color: var(--c-danger-text); }
.ab-btn.fail:hover { background: var(--c-danger-soft); }
.ab-btn.fail.active {
  background: var(--c-danger); color: white; border-color: var(--c-danger);
  box-shadow: var(--e-2);
}
.ab-btn.fix  { border-color: var(--c-warn); color: var(--c-warn-text); }
.ab-btn.fix:hover { background: var(--c-warn-soft); }
.ab-btn.fix.active  {
  background: var(--c-warn); color: white; border-color: var(--c-warn);
  box-shadow: var(--e-2);
}
.ab-btn.skip { border-color: var(--c-border-strong); color: var(--c-text-muted); }
.ab-btn.skip:hover { background: var(--c-surface-3); }
.ab-btn.skip.active {
  background: var(--c-text-soft); color: white; border-color: var(--c-text-soft);
  box-shadow: var(--e-1);
}
.ab-btn.auto { border-color: var(--c-purple); color: var(--c-purple-text); }
.ab-btn.auto:hover { background: var(--c-purple-soft); }
.ab-btn.mark { border-color: var(--c-teal); color: var(--c-teal-text); }
.ab-btn.mark:hover { background: var(--c-teal-soft); }
.ab-btn.mark.commitment {
  background: var(--c-warn-soft); border-color: var(--c-warn); color: var(--c-warn-text);
  animation: pulse-warn 2s ease-in-out infinite;
}
.ab-btn .kbd {
  background: rgba(0,0,0,0.06);
  padding: 0 4px; border-radius: var(--r-sm);
  font-family: var(--f-mono); font-size: var(--t-xs);
  border: none; color: inherit;
}
.ab-btn.active .kbd { background: rgba(255,255,255,0.25); color: inherit; border: none; }

.ab-bottom { display: flex; gap: var(--s-4); align-items: center; margin-top: var(--s-4); }
.ab-bottom textarea {
  flex: 1; resize: none; min-height: 32px; max-height: 100px;
  padding: 6px var(--s-5); font: inherit; font-size: var(--t-base);
  border: 1px solid var(--c-border-strong); border-radius: var(--r-md);
  background: var(--c-surface);
}
.ab-bottom .reset-btn {
  font-size: var(--t-sm); color: var(--c-text-soft);
  background: none; border: none;
  padding: 4px var(--s-4);
  border-radius: var(--r-sm);
  transition: color var(--d-fast) var(--ease-std), background var(--d-fast) var(--ease-std);
}
.ab-bottom .reset-btn:hover { color: var(--c-text); background: var(--c-surface-2); }

/* ── kbd-help (bottom-right hint) ───────────────────────────────── */
.kbd-help {
  position: fixed; bottom: var(--s-7); right: var(--s-7);
  font-size: var(--t-xs); color: var(--c-text-soft);
  background: var(--c-surface);
  padding: var(--s-3) var(--s-5);
  border-radius: var(--r-md);
  border: 1px solid var(--c-border);
  box-shadow: var(--e-1);
  display: inline-flex; gap: var(--s-4); align-items: center;
  opacity: 0.85;
  transition: opacity var(--d-base) var(--ease-std);
  z-index: 50;
}
.kbd-help:hover { opacity: 1; }
.kbd-help .kbd-group { display: inline-flex; gap: 3px; align-items: center; }

/* ── Sync banner ────────────────────────────────────────────────── */
.sync-banner {
  position: fixed; top: 0; left: 0; right: 0;
  z-index: 150;
  padding: var(--s-3) var(--s-7);
  font-size: var(--t-base); text-align: center;
  display: none;
  border-bottom: 1px solid;
}
.sync-banner.show { display: block; animation: slide-down var(--d-base) var(--ease-out); }
.sync-banner.warn   { background: var(--c-warn-soft); color: var(--c-warn-text); border-color: var(--c-warn); }
.sync-banner.danger { background: var(--c-danger-soft); color: var(--c-danger-text); border-color: var(--c-danger); }
.sync-banner button {
  margin-left: var(--s-6);
  padding: 3px var(--s-5);
  border: 1px solid currentColor; background: var(--c-surface);
  border-radius: var(--r-sm); cursor: pointer; font-weight: 600;
  color: inherit;
}
.sync-banner button:hover { background: rgba(0,0,0,0.05); }
.sync-banner .close-x {
  float: right; cursor: pointer;
  padding: 0 var(--s-4); opacity: 0.7;
  transition: opacity var(--d-fast) var(--ease-std);
}
.sync-banner .close-x:hover { opacity: 1; }

/* Versions toolbar buttons */
.versions-btn {
  font-size: var(--t-sm); padding: 4px var(--s-5);
  border: 1px solid var(--c-border);
  background: var(--c-surface);
  border-radius: var(--r-md);
  cursor: pointer;
  text-transform: none; letter-spacing: 0; font-weight: 600;
  color: var(--c-text-muted);
  display: inline-flex; align-items: center; gap: 4px;
  transition: all var(--d-fast) var(--ease-std);
}
.versions-btn:hover {
  background: var(--c-surface-2); color: var(--c-text);
  border-color: var(--c-border-strong);
}
.versions-btn.warn   { border-color: var(--c-warn); background: var(--c-warn-soft); color: var(--c-warn-text); }
.versions-btn.danger {
  border-color: var(--c-danger);
  background: var(--c-danger-soft);
  color: var(--c-danger-text);
  animation: pulse-warn 2s ease-in-out infinite;
}
@keyframes pulse-warn {
  0%, 100% { box-shadow: 0 0 0 0 rgba(239,68,68,0.4); }
  50%      { box-shadow: 0 0 0 4px rgba(239,68,68,0); }
}

/* Sync badge */
.sync-badge {
  display: inline-flex; align-items: center; gap: var(--s-3);
  padding: 4px var(--s-5);
  border-radius: var(--r-pill);
  font-size: var(--t-sm); font-weight: 600;
  margin: var(--s-3) 0 var(--s-5) 0;
}
.sync-badge.in-sync       { background: var(--c-success-soft); color: var(--c-success-text); }
.sync-badge.working-ahead { background: var(--c-info-soft); color: var(--c-info-text); }
.sync-badge.working-behind,
.sync-badge.divergent,
.sync-badge.incomplete-local { background: var(--c-danger-soft); color: var(--c-danger-text); }
.sync-badge.no-snapshots  { background: var(--c-surface-3); color: var(--c-text-soft); }
.sync-badge .latest-id {
  font-family: var(--f-mono); font-weight: 400;
  font-size: var(--t-xs); opacity: 0.85;
}

/* ── Modals ─────────────────────────────────────────────────────── */
.modal-bg {
  position: fixed; inset: 0;
  background: var(--c-overlay);
  display: none; align-items: center; justify-content: center;
  z-index: 200;
  padding: var(--s-7);
  backdrop-filter: blur(4px);
  -webkit-backdrop-filter: blur(4px);
}
.modal-bg[style*="display: flex"], .modal-bg.show {
  display: flex !important;
  animation: modal-bg-in var(--d-base) var(--ease-std);
}
@keyframes modal-bg-in { from { opacity: 0; } to { opacity: 1; } }

.modal {
  background: var(--c-surface);
  padding: var(--s-7) var(--s-8);
  border-radius: var(--r-xl);
  min-width: 480px; max-width: 640px;
  max-height: 85vh; overflow: auto;
  box-shadow: var(--e-5);
  border: 1px solid var(--c-border);
  animation: modal-in var(--d-slow) var(--ease-out);
}
@keyframes modal-in {
  from { opacity: 0; transform: translateY(8px) scale(0.985); }
  to   { opacity: 1; transform: translateY(0) scale(1); }
}
.modal h3 {
  margin: 0 0 var(--s-7);
  font-size: var(--t-xl); font-weight: 700;
  color: var(--c-text);
  display: flex; justify-content: space-between; align-items: center;
  gap: var(--s-4);
  border-bottom: 1px solid var(--c-divider);
  padding-bottom: var(--s-5);
}
.modal h3 .close {
  background: none; border: none; cursor: pointer;
  font-size: var(--t-xl); color: var(--c-text-soft);
  width: 32px; height: 32px;
  display: inline-flex; align-items: center; justify-content: center;
  border-radius: var(--r-md);
  transition: background var(--d-fast) var(--ease-std), color var(--d-fast) var(--ease-std);
}
.modal h3 .close:hover { background: var(--c-surface-2); color: var(--c-text); }

/* History modal */
.history-item {
  padding: var(--s-4) var(--s-5);
  border-radius: var(--r-md);
  margin-bottom: var(--s-2);
  background: var(--c-surface-2);
  display: flex; gap: var(--s-4); align-items: center;
  transition: background var(--d-fast) var(--ease-std);
}
.history-item:hover { background: var(--c-surface-3); }
.history-item .ts {
  font-family: var(--f-mono); font-size: var(--t-sm);
  color: var(--c-text-muted); flex: 1;
}
.history-item .size { font-size: var(--t-xs); color: var(--c-text-faint); }
.history-item button {
  font-size: var(--t-sm); padding: 4px var(--s-4);
  border: 1px solid var(--c-border-strong);
  background: var(--c-surface);
  color: var(--c-text);
  border-radius: var(--r-sm);
}
.history-item button:hover { background: var(--c-surface-3); }

/* ── Toasts ─────────────────────────────────────────────────────── */
.toast-stack {
  position: fixed;
  top: calc(var(--topbar-h) + var(--s-3));
  right: var(--s-7);
  display: flex; flex-direction: column; gap: var(--s-4);
  z-index: 300; pointer-events: none;
  max-width: 380px;
}
@media (max-width: 700px) {
  .toast-stack { top: 56px; }
}
.toast {
  pointer-events: auto;
  background: var(--c-surface);
  border-radius: var(--r-lg);
  padding: var(--s-5) var(--s-6);
  font-size: var(--t-base);
  box-shadow: var(--e-3);
  border-left: 4px solid var(--c-success);
  border-top: 1px solid var(--c-border);
  border-right: 1px solid var(--c-border);
  border-bottom: 1px solid var(--c-border);
  animation: toast-in var(--d-slow) var(--ease-out);
  max-width: 380px;
  position: relative;
  overflow: hidden;
}
.toast.warn  { border-left-color: var(--c-warn); }
.toast.error { border-left-color: var(--c-danger); }
.toast.info  { border-left-color: var(--c-info); }
.toast.learn { border-left-color: var(--c-purple); }
.toast.learn::before {
  content: ''; position: absolute; inset: 0;
  background: linear-gradient(135deg, var(--c-purple-soft) 0%, transparent 60%);
  pointer-events: none;
}
.toast .title {
  font-weight: 700; margin-bottom: 2px;
  color: var(--c-text);
  position: relative;
}
.toast .body {
  color: var(--c-text-muted);
  position: relative;
}
.toast .close {
  float: right; cursor: pointer; opacity: 0.5;
  padding: 0 var(--s-3);
  font-size: var(--t-md);
  transition: opacity var(--d-fast) var(--ease-std);
  position: relative;
}
.toast .close:hover { opacity: 1; }
@keyframes toast-in {
  from { transform: translateX(120%); opacity: 0; }
  to   { transform: translateX(0); opacity: 1; }
}
.toast.fading {
  transition: opacity 0.4s var(--ease-std), transform 0.4s var(--ease-std);
  opacity: 0; transform: translateX(80%);
}

/* ── Spinner ────────────────────────────────────────────────────── */
.spinner {
  display: inline-block;
  width: 24px; height: 24px;
  border: 3px solid var(--c-border);
  border-top-color: var(--c-primary);
  border-radius: 50%;
  animation: spin 0.8s linear infinite;
  margin-right: var(--s-4); vertical-align: middle;
}
@keyframes spin { to { transform: rotate(360deg); } }

/* ── Versions modal ─────────────────────────────────────────────── */
.versions-modal-body { min-width: 720px; max-width: 920px; }
.versions-snap-form {
  display: flex; gap: var(--s-3); align-items: center;
  margin-bottom: var(--s-5);
  padding: var(--s-4) var(--s-5);
  background: var(--c-surface-2);
  border: 1px solid var(--c-border);
  border-radius: var(--r-md);
}
.versions-snap-form input[type=text] {
  flex: 1; padding: 6px var(--s-5);
  border: 1px solid var(--c-border-strong);
  background: var(--c-surface);
  border-radius: var(--r-sm);
  font-size: var(--t-base);
}
.versions-snap-form .snap-quick { background: var(--c-info-soft); color: var(--c-info-text); border-color: var(--c-info-soft); }
.versions-snap-form .snap-quick:hover { background: var(--c-primary-soft); border-color: var(--c-primary); }
.versions-snap-form .snap-full  { background: var(--c-warn-soft); color: var(--c-warn-text); border-color: var(--c-warn-soft); }
.versions-snap-form .snap-full:hover { border-color: var(--c-warn); }
.versions-snap-form .auto-snap  { background: var(--c-success-soft); color: var(--c-success-text); border-color: var(--c-success-soft); }
.versions-snap-form .auto-snap:hover { border-color: var(--c-success); }

.versions-actions {
  display: flex; gap: var(--s-3); align-items: center;
  margin-bottom: var(--s-4);
  font-size: var(--t-sm); color: var(--c-text-soft);
}
.versions-actions .diff-helper { color: var(--c-text-faint); }

.versions-list {
  max-height: 50vh; overflow-y: auto;
  border: 1px solid var(--c-border); border-radius: var(--r-md);
  background: var(--c-surface);
}
.version-item {
  padding: var(--s-4) var(--s-5);
  border-bottom: 1px solid var(--c-divider);
  display: grid;
  grid-template-columns: 28px 1fr auto;
  gap: var(--s-4); align-items: center;
  transition: background var(--d-fast) var(--ease-std);
}
.version-item:last-child { border-bottom: none; }
.version-item:hover { background: var(--c-surface-2); }
.version-item.diff-selected { background: var(--c-warn-soft); }
.version-item .v-check { cursor: pointer; }
.version-item .v-tag {
  font-weight: 600; color: var(--c-text); font-size: var(--t-base);
  margin-bottom: 2px;
  display: flex; align-items: center; gap: var(--s-3);
}
.version-item .v-tag .untagged {
  color: var(--c-text-faint); font-style: italic; font-weight: 400;
}
.version-item .v-meta {
  font-size: var(--t-sm); color: var(--c-text-soft);
  font-family: var(--f-mono);
}
.version-item .v-kind {
  display: inline-block; padding: 1px var(--s-4);
  border-radius: var(--r-sm); margin-right: var(--s-3);
  font-size: 9px; font-weight: 700; text-transform: uppercase;
  letter-spacing: 0.04em;
}
.version-item .v-kind.quick { background: var(--c-info-soft); color: var(--c-info-text); }
.version-item .v-kind.full  { background: var(--c-warn-soft); color: var(--c-warn-text); }
.version-item .v-actions { display: flex; gap: var(--s-2); }
.version-item .v-actions button {
  font-size: var(--t-xs); padding: 4px var(--s-4);
  border: 1px solid var(--c-border-strong);
  background: var(--c-surface);
  color: var(--c-text);
  border-radius: var(--r-sm);
  transition: background var(--d-fast) var(--ease-std);
}
.version-item .v-actions button:hover { background: var(--c-surface-2); }
.version-item .v-actions .restore { color: var(--c-warn-text); border-color: var(--c-warn); }
.version-item .v-actions .restore-full { color: var(--c-danger-text); border-color: var(--c-danger); }

.versions-busy {
  padding: var(--s-7); text-align: center; color: var(--c-text-muted);
}
.versions-output {
  margin-top: var(--s-5);
  padding: var(--s-5);
  background: #111827; color: #e5e7eb;
  border-radius: var(--r-md);
  font-size: var(--t-sm);
  font-family: var(--f-mono);
  max-height: 300px; overflow: auto;
  white-space: pre-wrap; word-break: break-all;
  border: 1px solid #1f2937;
}

/* ── Learning modal ─────────────────────────────────────────────── */
.learn-pattern-list {
  max-height: 45vh; overflow: auto;
  border: 1px solid var(--c-border); border-radius: var(--r-md);
  background: var(--c-surface);
}
.learn-pattern-row {
  padding: var(--s-3) var(--s-5);
  border-bottom: 1px solid var(--c-divider);
  display: grid;
  grid-template-columns: 110px 140px 1fr 70px auto;
  gap: var(--s-4); align-items: center;
  font-size: var(--t-sm);
  transition: background var(--d-fast) var(--ease-std);
}
.learn-pattern-row:last-child { border-bottom: none; }
.learn-pattern-row:hover { background: var(--c-surface-2); }
.learn-pattern-row.disabled { opacity: 0.5; }
.learn-pattern-row .pt {
  display: inline-block; padding: 2px var(--s-4);
  border-radius: var(--r-sm);
  font-size: 9px; font-weight: 700; text-transform: uppercase;
  letter-spacing: 0.04em;
}
.learn-pattern-row .pt.filename_brand { background: var(--c-warn-soft); color: var(--c-warn-text); }
.learn-pattern-row .pt.section_vendor { background: var(--c-info-soft); color: var(--c-info-text); }
.learn-pattern-row .pt.row_format_d   { background: var(--c-success-soft); color: var(--c-success-text); }
.learn-pattern-row .trigger {
  font-family: var(--f-mono); font-size: var(--t-xs);
  background: var(--c-surface-3); padding: 2px var(--s-3);
  border-radius: var(--r-sm); color: var(--c-text-muted);
}
.learn-pattern-row .arrow { color: var(--c-text-faint); }
.learn-pattern-row .output { font-weight: 600; color: var(--c-text); }
.learn-pattern-row .conf {
  font-size: var(--t-xs); color: var(--c-text-soft);
  font-family: var(--f-mono);
}
.learn-pattern-row .conf .high { color: var(--c-success-text); font-weight: 700; }
.learn-pattern-row .conf .low { color: var(--c-danger-text); }

/* Confidence badge for plans */
.conf-badge {
  display: inline-block; padding: 2px var(--s-5);
  border-radius: var(--r-pill);
  font-size: var(--t-sm); font-weight: 700;
  margin-left: var(--s-3);
}
.conf-badge.high { background: var(--c-success-soft); color: var(--c-success-text); }
.conf-badge.med  { background: var(--c-warn-soft); color: var(--c-warn-text); }
.conf-badge.low  { background: var(--c-danger-soft); color: var(--c-danger-text); }
.conf-badge .src { font-size: 9px; opacity: 0.8; margin-left: 4px; }

/* ── Audit / DB modal ───────────────────────────────────────────── */
.audit-stats {
  display: grid; grid-template-columns: repeat(4, 1fr);
  gap: var(--s-4); margin-bottom: var(--s-6);
}
.audit-stats .stat-card {
  padding: var(--s-5);
  background: var(--c-surface-2);
  border-radius: var(--r-md);
  border: 1px solid var(--c-border);
  transition: transform var(--d-fast) var(--ease-out);
}
.audit-stats .stat-card:hover { transform: translateY(-1px); border-color: var(--c-border-strong); }
.audit-stats .stat-card .v {
  font-size: var(--t-2xl); font-weight: 700;
  color: var(--c-text); line-height: 1;
}
.audit-stats .stat-card .l {
  font-size: var(--t-xs); color: var(--c-text-soft);
  text-transform: uppercase; letter-spacing: 0.04em;
  margin-top: 2px;
}
.audit-search-row {
  display: flex; gap: var(--s-3); margin-bottom: var(--s-4);
}
.audit-search-row input {
  flex: 1;
}
.audit-search-row select {
  width: auto; min-width: 140px;
  padding: 6px var(--s-5);
}
.audit-list {
  max-height: 50vh; overflow: auto;
  border: 1px solid var(--c-border); border-radius: var(--r-md);
  background: var(--c-surface);
}
.audit-row {
  padding: var(--s-3) var(--s-5);
  border-bottom: 1px solid var(--c-divider);
  display: grid;
  grid-template-columns: 140px 130px 70px 1fr;
  gap: var(--s-4); align-items: center;
  font-size: var(--t-sm);
  transition: background var(--d-fast) var(--ease-std);
}
.audit-row:last-child { border-bottom: none; }
.audit-row:hover { background: var(--c-surface-2); }
.audit-row .ts {
  font-family: var(--f-mono); color: var(--c-text-soft); font-size: var(--t-xs);
}
.audit-row .ac { font-weight: 600; }
.audit-row .ac.status_change       { color: var(--c-success-text); }
.audit-row .ac.notes_update         { color: var(--c-info-text); }
.audit-row .ac.auto_annotate_apply  { color: var(--c-purple-text); }
.audit-row .ac.pdf_edit             { color: var(--c-warn-text); }
.audit-row .ac.snapshot             { color: var(--c-info-text); }
.audit-row .ac.restore              { color: var(--c-danger-text); }
.audit-row .tg {
  font-family: var(--f-mono); font-size: var(--t-xs);
  color: var(--c-text-muted); cursor: pointer;
}
.audit-row .tg:hover { color: var(--c-text); text-decoration: underline; }
.audit-row .ba { color: var(--c-text-soft); font-size: var(--t-xs); }
.audit-row .ba .arrow { color: var(--c-text-faint); margin: 0 var(--s-2); }
.audit-row .ba .before { color: var(--c-danger-text); }
.audit-row .ba .after  { color: var(--c-success-text); }
.audit-search-results .ar-row {
  padding: var(--s-3) var(--s-5);
  border-bottom: 1px solid var(--c-divider);
  font-size: var(--t-sm); cursor: pointer;
  transition: background var(--d-fast) var(--ease-std);
}
.audit-search-results .ar-row:hover { background: var(--c-warn-soft); }
.audit-search-results mark {
  background: var(--c-warn-soft); color: var(--c-warn-text);
  font-weight: 600; padding: 0 2px; border-radius: 2px;
}

/* ── Auto-annotate modal ────────────────────────────────────────── */
.auto-block { margin-bottom: var(--s-5); }
.auto-label {
  font-size: var(--t-xs); font-weight: 700;
  color: var(--c-text-soft);
  text-transform: uppercase; margin-bottom: 4px;
  letter-spacing: 0.06em;
}
.auto-pre {
  margin: 0; padding: var(--s-4) var(--s-5);
  background: var(--c-surface-2);
  border: 1px solid var(--c-border);
  border-radius: var(--r-md);
  font-size: var(--t-base);
  max-height: 100px; overflow: auto;
  white-space: pre-wrap; word-break: break-word;
  font-family: inherit;
  color: var(--c-text);
}
.auto-pre.empty { color: var(--c-text-faint); font-style: italic; }
.auto-batch-list {
  max-height: 50vh; overflow: auto;
  border: 1px solid var(--c-border); border-radius: var(--r-md);
}
.auto-batch-row {
  display: grid; grid-template-columns: 28px 56px 110px 1fr auto;
  gap: var(--s-3); align-items: center;
  padding: var(--s-3) var(--s-4);
  border-bottom: 1px solid var(--c-divider);
  font-size: var(--t-sm);
}
.auto-batch-row:last-child { border-bottom: none; }
.auto-batch-row .b-row { font-family: var(--f-mono); color: var(--c-text-soft); }
.auto-batch-row .b-role {
  display: inline-block; padding: 1px var(--s-3);
  border-radius: var(--r-sm);
  font-size: 9px; font-weight: 700; text-transform: uppercase;
  letter-spacing: 0.04em;
}
.auto-batch-row .b-role.section_header { background: var(--c-info-soft); color: var(--c-info-text); }
.auto-batch-row .b-role.item            { background: var(--c-warn-soft); color: var(--c-warn-text); }
.auto-batch-row .b-role.sub_item        { background: var(--c-success-soft); color: var(--c-success-text); }
.auto-batch-row .b-role.unknown         { background: var(--c-surface-3); color: var(--c-text-soft); }
.auto-batch-row .b-d {
  font-size: var(--t-xs); color: var(--c-text-muted);
  white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
}
.auto-batch-row.fail { background: var(--c-danger-soft); }
.auto-batch-row.warn { background: var(--c-warn-soft); }

/* ── Reduced motion ─────────────────────────────────────────────── */
@media (prefers-reduced-motion: reduce) {
  *, *::before, *::after {
    animation-duration: 0.01ms !important;
    animation-iteration-count: 1 !important;
    transition-duration: 0.01ms !important;
    scroll-behavior: auto !important;
  }
}

/* ── Responsive ─────────────────────────────────────────────────── */
@media (max-width: 1280px) {
  #app {
    grid-template-columns: minmax(var(--tree-w-md), 280px) 1fr;
    grid-template-areas:
      "topbar topbar"
      "tree   center";
  }
  .pdf-pane { display: none; }
  .pdf-pane.expand { display: flex; grid-column: 2; grid-area: center; }
}

@media (max-width: 900px) {
  #app {
    grid-template-columns: 1fr;
    grid-template-rows: var(--topbar-h) 240px 1fr;
    grid-template-areas:
      "topbar"
      "tree"
      "center";
  }
  .tree-pane { border-right: none; border-bottom: 1px solid var(--c-border); }
}

@media (max-width: 700px) {
  .topbar { display: none; }
  .mobile-tabs { display: flex; }
  #app {
    grid-template-columns: 1fr;
    grid-template-rows: 44px 1fr;
    grid-template-areas:
      "topbar"
      "tree";
  }
  #app > .pane { display: none; grid-area: tree; }
  #app > .pane.mobile-active { display: flex; }
  .action-bar {
    min-width: auto; left: var(--s-4); right: var(--s-4);
    transform: none; bottom: var(--s-4);
    padding: var(--s-4) var(--s-5);
    border-radius: var(--r-lg);
  }
  .ab-top { flex-wrap: wrap; gap: var(--s-3); }
  .ab-buttons { flex-wrap: wrap; }
  .ab-btn { padding: 5px var(--s-4); font-size: var(--t-sm); }
  .modal { min-width: auto !important; max-width: calc(100vw - 24px); padding: var(--s-6); }
  .audit-stats { grid-template-columns: repeat(2, 1fr) !important; }
  .versions-modal-body { min-width: auto !important; }
  .kbd-help { display: none; }
}

@media (max-width: 480px) {
  .audit-stats { grid-template-columns: 1fr !important; }
  .audit-row { grid-template-columns: 1fr !important; gap: 2px !important; }
  .topbar-actions .stats-pill { display: none; }
}

/* ── Print (subtle) ─────────────────────────────────────────────── */
@media print {
  .topbar, .action-bar, .mobile-tabs, .toast-stack, .kbd-help, .canvas-toolbar { display: none !important; }
  #app { display: block; height: auto; }
  .pane { border: none; height: auto; }
  body { background: white; }
}
</style>
</head>
<body>

<!-- Sync banner (shown when working dir differs from latest snapshot) -->
<div id="sync-banner" class="sync-banner" role="status" aria-live="polite">
  <span id="sync-banner-msg"></span>
  <span id="sync-banner-actions"></span>
  <span class="close-x" onclick="dismissSyncBanner()" role="button" tabindex="0" aria-label="ปิด banner">✕</span>
</div>

<div id="app">

  <!-- ───── TOPBAR ──────────────────────────────────────────────── -->
  <header class="topbar" role="banner">
    <div class="topbar-brand">
      <span class="logo" aria-hidden="true">SP</span>
      <span>Comply Verify <span class="sub">Smart Plant 1</span></span>
    </div>
    <div class="topbar-spacer"></div>
    <div class="topbar-actions">
      <span class="stats-pill" id="stats-pill" role="status" aria-label="overall progress">
        <span id="stats-pill-progress">—</span>
      </span>
      <button class="versions-btn" onclick="toggleTheme()" title="สลับธีม" aria-label="สลับธีม light/dark" id="theme-toggle">🌓</button>
      <button class="versions-btn" onclick="showLearning()" title="HITL learning loop / patterns">🧠 Learn</button>
      <button class="versions-btn" onclick="showAudit()" title="Audit log + DB stats">📊 Audit</button>
      <button class="versions-btn" onclick="showVersions()" title="โปรเจกต์ snapshots / versioning">📚 Versions</button>
    </div>
  </header>

  <!-- Mobile tab navigation (replaces topbar at <700px) -->
  <nav class="mobile-tabs" role="tablist" aria-label="pane navigation">
    <button class="active" data-tab="tree"   onclick="setMobileTab('tree')"   role="tab" aria-selected="true">🌲 Tree</button>
    <button              data-tab="center" onclick="setMobileTab('center')" role="tab" aria-selected="false">📄 TOR/xlsx</button>
    <button              data-tab="pdf"    onclick="setMobileTab('pdf')"    role="tab" aria-selected="false">📑 Catalog</button>
  </nav>

  <!-- ───── LEFT: Tree ──────────────────────────────────────────── -->
  <section class="pane tree-pane" aria-label="row tree">
    <h2>
      <span class="pane-title"><span class="icon" aria-hidden="true">🗂</span><span>โครงสร้าง</span> <span id="tree-count" style="font-weight:400;color:var(--c-text-faint)"></span></span>
    </h2>
    <div class="stats-bar" id="stats" aria-live="polite"></div>
    <div class="toolbar" role="search">
      <input type="search" id="tree-search" placeholder="🔍 ค้น section / row / spec…" aria-label="ค้นหา row">
      <div class="filter-row">
        <select id="filter-status" aria-label="กรอง status">
          <option value="">ทุกสถานะ</option>
          <option value="unverified">ยังไม่ตรวจ</option>
          <option value="pass">ผ่าน</option>
          <option value="fail">ไม่ผ่าน</option>
          <option value="need_fix">ต้องแก้</option>
          <option value="skip">ข้าม</option>
        </select>
        <select id="filter-vendor" aria-label="กรอง vendor">
          <option value="">ทุก vendor</option>
          <option value="SMART">SMART</option>
          <option value="TRIO">TRIO</option>
          <option value="SR">SR</option>
        </select>
      </div>
      <div class="filter-row">
        <label><input type="checkbox" id="filter-flags"> มี flag</label>
        <label><input type="checkbox" id="filter-haspdf"> มี catalog</label>
        <button class="mini-btn" onclick="expandAll()" title="แตกทุก section">⊞</button>
        <button class="mini-btn" onclick="collapseAll()" title="หุบทุก section">⊟</button>
      </div>
    </div>
    <div class="tree-scroll" id="tree" role="tree" aria-label="rows"></div>
  </section>

  <!-- ───── CENTER: TOR top + xlsx bot ──────────────────────────── -->
  <section class="pane center-pane" aria-label="TOR and spreadsheet">
    <!-- Top: TOR -->
    <div class="split-top">
      <h2>
        <span class="pane-title"><span class="icon" aria-hidden="true">📄</span><span>TOR</span> <span style="font-weight:400;color:var(--c-text-faint)" id="tor-info"></span></span>
      </h2>
      <div class="canvas-toolbar" role="toolbar" aria-label="TOR navigation">
        <button onclick="torPrev()" title="หน้าก่อน" aria-label="หน้าก่อน">◀</button>
        <span class="info" id="tor-page-info">— / —</span>
        <button onclick="torNext()" title="หน้าถัดไป" aria-label="หน้าถัดไป">▶</button>
        <span class="info" id="tor-status">เลือก row เพื่อดู</span>
        <span style="flex:1"></span>
        <button onclick="torZoom(-1)" title="ย่อ" aria-label="ย่อ">−</button>
        <button onclick="torZoom(1)" title="ขยาย" aria-label="ขยาย">＋</button>
        <button onclick="torJumpToMatch()" title="ไปยังหน้าที่เจอ">⌖</button>
      </div>
      <div class="tor-canvas" id="tor-canvas">
        <div class="empty-canvas">เลือก row เพื่อดู TOR</div>
      </div>
    </div>

    <!-- Bottom: xlsx preview -->
    <div class="split-bot">
      <h2>
        <span class="pane-title"><span class="icon" aria-hidden="true">📊</span><span>Comply.xlsx</span> <span style="font-weight:400;color:var(--c-text-faint)" id="xlsx-info"></span></span>
      </h2>
      <div class="canvas-toolbar" role="toolbar" aria-label="spreadsheet context">
        <span class="info">บริบท ±<span id="ctx-radius">6</span> rows</span>
        <button onclick="ctxRadius(-2)" title="ลด context">−2</button>
        <button onclick="ctxRadius(2)" title="เพิ่ม context">+2</button>
        <span style="flex:1"></span>
        <span class="info">คลิกแถวเพื่อ select</span>
      </div>
      <div class="xlsx-wrap" id="xlsx-wrap">
        <div class="empty-canvas">เลือก row จากต้นไม้</div>
      </div>
    </div>
  </section>

  <!-- ───── RIGHT: Catalog PDF ──────────────────────────────────── -->
  <section class="pane pdf-pane" aria-label="catalog PDF">
    <h2>
      <span class="pane-title"><span class="icon" aria-hidden="true">📑</span><span>Catalog</span></span>
      <span class="filename" id="pdf-filename">(ไม่มี)</span>
      <button class="edit-toggle-btn" id="edit-toggle-btn" onclick="toggleEditMode()" title="Edit annotations" aria-pressed="false">✏ Edit</button>
    </h2>
    <div class="canvas-toolbar" role="toolbar" aria-label="catalog navigation">
      <button onclick="pdfPrev()" title="หน้าก่อน" aria-label="หน้าก่อน">◀</button>
      <span class="info" id="pdf-page-info">— / —</span>
      <button onclick="pdfNext()" title="หน้าถัดไป" aria-label="หน้าถัดไป">▶</button>
      <span style="flex:1"></span>
      <button onclick="pdfZoom(-1)" title="ย่อ" aria-label="ย่อ">−</button>
      <button onclick="pdfZoom(1)" title="ขยาย" aria-label="ขยาย">＋</button>
      <label><input type="checkbox" id="hl-toggle" checked onchange="renderPdf()"> highlight</label>
      <button onclick="openInBrowser()" title="เปิดใน browser">⤴</button>
    </div>
    <!-- Edit toolbar (visible only in edit mode) -->
    <div class="edit-toolbar" id="edit-toolbar" style="display:none;">
      <button class="tool active" data-tool="select" onclick="setTool('select')" title="Select / move (V)">⬚</button>
      <button class="tool" data-tool="drawRect" onclick="setTool('drawRect')" title="Draw rectangle (R)">▭</button>
      <button class="tool" data-tool="addText" onclick="setTool('addText')" title="Add text (T)">T</button>
      <button onclick="deleteSelected()" title="Delete selected (Del)">🗑</button>
      <span class="sep"></span>
      <button onclick="undo()" id="undo-btn" disabled title="Undo (⌘Z)">↶</button>
      <button onclick="redo()" id="redo-btn" disabled title="Redo (⇧⌘Z)">↷</button>
      <span class="sep"></span>
      <button onclick="showHistory()" title="Version history">⏰ History</button>
      <span style="flex:1"></span>
      <span class="dirty-indicator" id="dirty-ind"></span>
      <button onclick="saveEdits()" id="save-btn" class="save-btn" disabled title="Save (⌘S)">💾 Save</button>
    </div>
    <!-- Manual-annotate mode banner -->
    <div class="manual-mode-banner" id="manual-banner">
      <div class="target-info" id="manual-target-info">—</div>
      <button class="cancel" onclick="cancelManualAnnotate()">✕ Cancel</button>
      <button class="save" id="manual-save-btn" onclick="saveManualAnnotate()" disabled>✓ Save & update Col D</button>
    </div>
    <div class="pdf-canvas" id="pdf-canvas">
      <div class="empty-canvas">เลือก row ที่อ้างอิง catalog</div>
    </div>
    <div class="pdf-annots" id="pdf-annots" aria-label="annotations on current page"></div>
  </section>
</div>

<!-- Per-PDF history modal (catalog edits) -->
<div class="modal-bg" id="history-modal" role="dialog" aria-modal="true" aria-labelledby="history-modal-title" onclick="if(event.target.id==='history-modal') closeHistory()">
  <div class="modal">
    <h3 id="history-modal-title">📑 Catalog edit history <button class="close" onclick="closeHistory()" aria-label="ปิด">✕</button></h3>
    <div id="history-info" style="font-size:var(--t-sm);color:var(--c-text-soft);margin-bottom:var(--s-4);"></div>
    <ul id="history-list" style="list-style:none;padding:0;margin:0;"></ul>
    <div style="margin-top:var(--s-6);display:flex;gap:var(--s-4);">
      <button onclick="manualSnapshot()" class="btn">📷 Snapshot now</button>
      <button onclick="closeHistory()" class="btn">Close</button>
    </div>
  </div>
</div>

<!-- Project-level versions modal (snap, restore, diff via version.py) -->
<div class="modal-bg" id="versions-modal" role="dialog" aria-modal="true" aria-labelledby="versions-modal-title" onclick="if(event.target.id==='versions-modal') closeVersions()">
  <div class="modal versions-modal-body">
    <h3 id="versions-modal-title">📚 Project versions
      <button class="close" onclick="closeVersions()" aria-label="ปิด">✕</button>
    </h3>
    <div id="versions-sync-badge"></div>
    <div id="versions-info" style="font-size:var(--t-sm);color:var(--c-text-soft);margin-bottom:var(--s-5);"></div>

    <div class="versions-snap-form">
      <input type="text" id="versions-tag" placeholder='tag (เช่น "before-rcbo-edit")' aria-label="snapshot tag">
      <button class="btn snap-quick" onclick="takeProjectSnap(false)">📷 Quick snap</button>
      <button class="btn snap-full"  onclick="takeProjectSnap(true)" title="รวม output/ ทั้งหมด (~200 MB)">📦 Full snap</button>
      <button class="btn auto-snap" onclick="autoSnapNow()" title="snap เฉพาะถ้า xlsx เปลี่ยน">⚡ Auto</button>
    </div>

    <div class="versions-actions" id="versions-bulk-actions">
      <button class="btn" onclick="pruneVersions()">🗑 Prune (keep 10)</button>
      <button class="btn" onclick="loadVersions()">↻ Refresh</button>
      <span style="flex:1"></span>
      <span class="diff-helper" id="diff-helper">เลือก 2 snapshots เพื่อ diff</span>
    </div>

    <div class="versions-list" id="versions-list">
      <div style="color:var(--c-text-faint);padding:var(--s-8);text-align:center;">Loading…</div>
    </div>

    <div id="versions-busy" class="versions-busy" style="display:none;">
      <div class="spinner"></div>
      <div id="versions-busy-msg">working…</div>
    </div>

    <pre id="versions-output" class="versions-output" style="display:none;"></pre>
  </div>
</div>

<!-- HITL Learning modal -->
<div class="modal-bg" id="learn-modal" role="dialog" aria-modal="true" aria-labelledby="learn-modal-title" onclick="if(event.target.id==='learn-modal') closeLearning()">
  <div class="modal" style="min-width:780px;max-width:1000px;">
    <h3 id="learn-modal-title">🧠 HITL Learning
      <button class="close" onclick="closeLearning()" aria-label="ปิด">✕</button>
    </h3>
    <div style="font-size:var(--t-sm);color:var(--c-text-soft);margin-bottom:var(--s-5);">
      Core proposes → user does visual proof → corrections become rules.
      Patterns ที่ user แก้ซ้ำ ≥ 2 ครั้ง จะถูก promote เป็น rule อัตโนมัติ.
    </div>

    <div class="audit-stats" id="learn-stats"></div>

    <div style="display:flex;gap:var(--s-4);align-items:center;margin:var(--s-6) 0 var(--s-3);">
      <strong style="font-size:var(--t-base);">Learned patterns</strong>
      <select id="learn-pattern-filter" onchange="loadLearnPatterns()" style="width:auto;min-width:140px;">
        <option value="">[all types]</option>
        <option value="filename_brand">filename_brand</option>
        <option value="section_vendor">section_vendor</option>
        <option value="row_format_d">row_format_d</option>
      </select>
      <span style="flex:1"></span>
      <button class="btn save-btn" onclick="runRetrain()">🔄 Retrain now</button>
    </div>

    <div class="learn-pattern-list" id="learn-pattern-list"></div>

    <div id="learn-llm-row" style="margin-top:var(--s-6);font-size:var(--t-sm);color:var(--c-text-soft);">
      <strong>LLM provider:</strong> <span id="learn-llm-name">off</span>
      — เปิดใช้ทาง env var <code>COMPLY_LLM</code> (Anthropic / OpenAI / local Ollama)
    </div>

    <pre id="learn-output" class="versions-output" style="display:none;margin-top:var(--s-5);"></pre>
  </div>
</div>

<!-- Audit log + DB stats modal -->
<div class="modal-bg" id="audit-modal" role="dialog" aria-modal="true" aria-labelledby="audit-modal-title" onclick="if(event.target.id==='audit-modal') closeAudit()">
  <div class="modal" style="min-width:780px;max-width:1000px;">
    <h3 id="audit-modal-title">📊 Database &amp; Audit
      <button class="close" onclick="closeAudit()" aria-label="ปิด">✕</button>
    </h3>

    <!-- DB stats summary -->
    <div id="audit-stats" class="audit-stats"></div>

    <!-- Search bar (FTS5 over rows) -->
    <div class="audit-search-row">
      <input type="search" id="audit-search" placeholder="🔎 ค้นหาข้าม Col A/B/C/D/E (FTS5)…" oninput="onAuditSearch()" aria-label="full-text search">
      <select id="audit-action-filter" onchange="loadAudit()" aria-label="กรอง action">
        <option value="">[all actions]</option>
        <option value="status_change">status_change</option>
        <option value="notes_update">notes_update</option>
        <option value="auto_annotate_apply">auto_annotate_apply</option>
        <option value="pdf_edit">pdf_edit</option>
        <option value="snapshot">snapshot</option>
        <option value="restore">restore</option>
        <option value="refresh">refresh</option>
      </select>
    </div>

    <!-- FTS results (only when searching) -->
    <div id="audit-search-results" style="display:none;"></div>

    <!-- Audit log timeline -->
    <div id="audit-list-wrap">
      <h4 style="margin: var(--s-6) 0 var(--s-3); font-size: var(--t-base); color: var(--c-text-muted);">Recent activity</h4>
      <div class="audit-list" id="audit-list"></div>
    </div>
  </div>
</div>

<!-- Auto-annotate modal -->
<div class="modal-bg" id="auto-modal" role="dialog" aria-modal="true" aria-labelledby="auto-modal-title" onclick="if(event.target.id==='auto-modal') closeAuto()">
  <div class="modal" style="min-width:640px;max-width:800px;">
    <h3 id="auto-modal-title">✨ Auto-annotate <span id="auto-title" style="font-weight:400;color:var(--c-text-soft);font-size:var(--t-base);"></span>
      <button class="close" onclick="closeAuto()" aria-label="ปิด">✕</button>
    </h3>

    <!-- Per-row preview -->
    <div id="auto-single">
      <div id="auto-meta" style="font-size:var(--t-sm);color:var(--c-text-soft);margin-bottom:var(--s-4);"></div>
      <div id="auto-warn" class="versions-output" style="display:none;background:var(--c-warn-soft);color:var(--c-warn-text);"></div>

      <div class="auto-block">
        <div class="auto-label">Col C (proposed)</div>
        <pre class="auto-pre" id="auto-c"></pre>
      </div>
      <div class="auto-block">
        <div class="auto-label">Col D (proposed)</div>
        <pre class="auto-pre" id="auto-d"></pre>
      </div>
      <div class="auto-block">
        <div class="auto-label">PDF annotations <span id="auto-ann-count"></span></div>
        <pre class="auto-pre" id="auto-ann"></pre>
      </div>

      <div style="margin-top:var(--s-6);display:flex;gap:var(--s-4);align-items:center;flex-wrap:wrap;">
        <label style="font-size:var(--t-sm);"><input type="checkbox" id="auto-write-xlsx" checked> เขียน Col C/D ลง xlsx</label>
        <label style="font-size:var(--t-sm);"><input type="checkbox" id="auto-write-pdf" checked> เขียน rect+label ลง PDF</label>
        <span style="flex:1"></span>
        <button class="btn" onclick="closeAuto()">Cancel</button>
        <button class="btn" onclick="showBatchAuto()">📚 Batch mode…</button>
        <button class="btn save-btn" onclick="applyAutoSingle()">✓ Apply</button>
      </div>
    </div>

    <!-- Batch preview -->
    <div id="auto-batch" style="display:none;">
      <div id="auto-batch-info" style="font-size:var(--t-sm);color:var(--c-text-soft);margin-bottom:var(--s-3);"></div>
      <div class="auto-batch-list" id="auto-batch-list"></div>
      <div style="margin-top:var(--s-6);display:flex;gap:var(--s-4);align-items:center;flex-wrap:wrap;">
        <button class="btn" onclick="showSingleAuto()">← back to single</button>
        <span style="flex:1"></span>
        <button class="btn" onclick="closeAuto()">Cancel</button>
        <button class="btn save-btn" onclick="applyAutoBatch()" id="auto-batch-apply">✓ Apply selected</button>
      </div>
    </div>

    <pre id="auto-output" class="versions-output" style="display:none;margin-top:var(--s-5);"></pre>
  </div>
</div>

<!-- Floating action bar -->
<div class="action-bar" id="action-bar" style="display:none;" role="toolbar" aria-label="row verdict">
  <div class="ab-top">
    <div class="ab-row-info" id="ab-row-info">—</div>
    <div class="ab-flags" id="ab-flags"></div>
    <span class="ab-spacer"></span>
    <div class="ab-buttons">
      <button class="ab-btn pass" onclick="setStatus('pass')" aria-label="ผ่าน">✓ ผ่าน <span class="kbd" aria-hidden="true">1</span></button>
      <button class="ab-btn fail" onclick="setStatus('fail')" aria-label="ไม่ผ่าน">✗ ไม่ผ่าน <span class="kbd" aria-hidden="true">2</span></button>
      <button class="ab-btn fix"  onclick="setStatus('need_fix')" aria-label="ต้องแก้">⚠ แก้ <span class="kbd" aria-hidden="true">3</span></button>
      <button class="ab-btn skip" onclick="setStatus('skip')" aria-label="ข้าม">⏭ ข้าม <span class="kbd" aria-hidden="true">4</span></button>
      <button class="ab-btn auto" onclick="showAutoAnnotate()" title="Auto-annotate row (preview)">✨ Auto</button>
      <button class="ab-btn mark" id="ab-mark-btn" onclick="startManualAnnotate()" title="ลาก rect ใน catalog เพื่อแก้ Col D ที่เป็น 'ยินดีปฏิบัติ' (เมื่อ AI หาเนื้อหาไม่เจอ)">📍 Mark</button>
    </div>
    <label style="font-size:var(--t-sm);display:flex;align-items:center;gap:4px;cursor:pointer;color:var(--c-text-muted);margin-left:var(--s-4);">
      <input type="checkbox" id="auto-advance" onchange="toggleAutoAdvance()" checked> auto-next
    </label>
  </div>
  <div class="ab-bottom">
    <textarea id="ab-notes" placeholder="บันทึก / Notes…" oninput="saveNotesDebounced()" aria-label="row notes"></textarea>
    <button class="reset-btn" onclick="setStatus('unverified')" title="reset verdict">↺ reset</button>
  </div>
</div>

<!-- Toast notifications (top-right) -->
<div id="toasts" class="toast-stack" role="region" aria-live="polite" aria-label="notifications"></div>

<aside class="kbd-help" aria-label="keyboard shortcuts">
  <span class="kbd-group"><span class="kbd">J</span><span class="kbd">K</span> rows</span>
  <span class="kbd-group"><span class="kbd">N</span> next-uncertain</span>
  <span class="kbd-group"><span class="kbd">1</span>–<span class="kbd">4</span> verdict</span>
  <span class="kbd-group"><span class="kbd">[</span><span class="kbd">]</span> PDF</span>
  <span class="kbd-group"><span class="kbd">,</span><span class="kbd">.</span> TOR</span>
</aside>

<script>
let DATA = null;
let SELECTED_ROW = null;
let ROWS_BY_NUM = {};
let TREE_ROOT = null;
let VISIBLE_ROWS = new Set();          // rows passing filter
let EXPANDED = new Set();              // tree node keys expanded
let CTX_RADIUS = 6;
let TOR_DPI = 110, PDF_DPI = 130;
let TOR_PAGE = 1, TOR_TARGET_PAGE = 1;
let TOR_PAGES = 0, TOR_HITS = 0;
let CURRENT_PDF = null, PDF_PAGE = 1, CURRENT_HIGHLIGHT = null;

async function init() {
  const r = await fetch('/api/index');
  DATA = await r.json();
  ROWS_BY_NUM = Object.fromEntries(DATA.rows.map(r => [r.row, r]));
  TREE_ROOT = DATA.tree;
  buildRowBlobs();

  // expand top-level by default
  for (const c of TREE_ROOT.children) EXPANDED.add(c.key);

  applyFilters();
  renderStats();
  renderTree();

  // Reflect "always work on latest" status on the main UI
  if (DATA.version_sync) refreshSyncIndicators(DATA.version_sync);

  // Periodically refresh sync status (every 60s) so divergence detected
  // by other tools (e.g. Google Drive sync conflicts) shows up.
  setInterval(async () => {
    try {
      const r = await fetch('/api/versions/sync');
      const sync = await r.json();
      refreshSyncIndicators(sync);
    } catch (e) {}
  }, 60000);

  const last = parseInt(localStorage.getItem('lastRow'));
  if (last && ROWS_BY_NUM[last]) selectRow(last, false);
  else if (DATA.rows.length) selectRow(DATA.rows[0].row, false);
}

// ── Search normalization ───────────────────────────────────────
/* Match xlsx/comply text regardless of:
   - precomposed vs decomposed Thai SARA AM (ำ ↔ ํา)
   - case
   - Whitespace differences
   Also support multi-token AND matching ("RCBO Schneider" → both must appear).
*/
function normalizeForSearch(s) {
  if (s == null) return '';
  return String(s).toLowerCase()
    .replace(/ํา/g, 'ำ')   // decomposed → precomposed SARA AM
    .replace(/\s+/g, ' ');
}
function tokenizeQuery(q) {
  // split on whitespace; quote support: "a b" treated as one token
  const tokens = [];
  const re = /"([^"]+)"|(\S+)/g;
  let m;
  while ((m = re.exec(q)) !== null) {
    const t = (m[1] || m[2] || '').trim();
    if (t) tokens.push(normalizeForSearch(t));
  }
  return tokens;
}
const SECTION_QUERY = /^\d+(\.\d+){0,4}\.?$/;

// Per-row search blob cache (rebuilt on init, not on every filter)
let ROW_BLOBS = new Map();
function buildRowBlobs() {
  ROW_BLOBS.clear();
  for (const r of DATA.rows) {
    const parts = [
      `r${r.row}`, r.A||'', r.B||'', r.C||'', r.D||'', r.E||'', r.F||'',
      r.section||'', (r.parsed && r.parsed.brand)||'', (r.parsed && r.parsed.model)||'',
    ];
    ROW_BLOBS.set(r.row, normalizeForSearch(parts.join(' ')));
  }
}

// ── Filtering ──────────────────────────────────────────────────
function applyFilters() {
  const rawQ = document.getElementById('tree-search').value.trim();
  const tokens = tokenizeQuery(rawQ);
  const fStat = document.getElementById('filter-status').value;
  const fVen  = document.getElementById('filter-vendor').value;
  const fFlag = document.getElementById('filter-flags').checked;
  const fPdf  = document.getElementById('filter-haspdf').checked;
  const status = DATA.status || {};

  // Detect "pure section number" query (e.g. "5.1.2") and expand its match
  // semantics: match rows whose section IS or BEGINS WITH this prefix.
  let sectionPrefix = null;
  if (tokens.length === 1 && SECTION_QUERY.test(tokens[0])) {
    sectionPrefix = tokens[0].replace(/\.$/, '');  // strip trailing dot
  }

  VISIBLE_ROWS.clear();
  for (const r of DATA.rows) {
    const st = (status[r.row] && status[r.row].status) || 'unverified';
    if (fStat && st !== fStat) continue;
    if (fVen && r.E !== fVen) continue;
    if (fFlag && (!r.auto_flags || !r.auto_flags.length)) continue;
    if (fPdf && !r.pdf_rel) continue;

    if (tokens.length) {
      const blob = ROW_BLOBS.get(r.row) || '';
      if (sectionPrefix) {
        // Section query: row's section must start with the prefix
        const sec = (r.section || '').toString();
        if (sec !== sectionPrefix && !sec.startsWith(sectionPrefix + '.')) continue;
      } else {
        // All tokens must appear (AND)
        let ok = true;
        for (const t of tokens) { if (!blob.includes(t)) { ok = false; break; } }
        if (!ok) continue;
      }
    }
    VISIBLE_ROWS.add(r.row);
  }

  // Auto-expand path to all matched rows when a search is active
  if (tokens.length) {
    for (const rn of VISIBLE_ROWS) expandPathToRow(rn);
  }
  renderTree();
}

// ── Tree rendering ─────────────────────────────────────────────
function visibleRowsInNode(node) {
  // count rows in this node + descendants that pass filter
  let count = 0;
  for (const r of node.rows) if (VISIBLE_ROWS.has(r)) count++;
  for (const c of node.children) count += visibleRowsInNode(c);
  return count;
}
function nodeHasMatch(node) { return visibleRowsInNode(node) > 0; }

function renderTree() {
  const out = [];
  for (const c of TREE_ROOT.children) renderNode(c, 0, out);
  document.getElementById('tree').innerHTML = out.join('');
  document.getElementById('tree-count').textContent = `(${VISIBLE_ROWS.size})`;
}

function renderNode(node, depth, out) {
  if (!nodeHasMatch(node)) return;
  const expanded = EXPANDED.has(node.key);
  const hasKids = node.children.length > 0 || node.rows.length > 0;

  // Render the node row(s):
  // Section/item/sub nodes show their representative row (first one) + children
  const repRow = node.rows[0];
  const r = repRow ? ROWS_BY_NUM[repRow] : null;
  const status = (DATA.status && DATA.status[repRow]) ? DATA.status[repRow].status : 'unverified';
  const stIcon = {pass:'✓', fail:'✗', need_fix:'⚠', skip:'⏭', unverified:''}[status] || '';
  const stCol  = {pass:'#10b981', fail:'#ef4444', need_fix:'#f59e0b', skip:'#6b7280', unverified:''}[status] || '';

  const flagCount = r && r.auto_flags ? r.auto_flags.length : 0;
  const flagHtml = flagCount ? `<span class="tree-flags" title="${escapeHtml(r.auto_flags.map(f=>f.msg).join(' / '))}" style="color:#d97706">⚠${flagCount}</span>` : '';

  const icon = ({section:'§', item:'›', sub:'·', root:''})[node.type] || '';
  const iconCls = node.type;
  const indent = depth * 12;

  let label = node.label;
  // Append context for items/subs
  if (r) {
    const preview = (r.B || r.C || '').toString().trim().replace(/\s+/g,' ').slice(0, 60);
    if (node.type === 'section' && r.D && /ยี่ห้อ/.test(r.D)) {
      const bm = r.D.match(/ยี่ห้อ\s+([^รุ่น]+?)\s+รุ่น\s+(.+)/);
      if (bm) label = `${node.label} — ${bm[1].trim()} ${bm[2].trim()}`.slice(0, 60);
    } else if (preview && (node.type === 'item' || node.type === 'sub')) {
      label = `${node.label} ${preview}`.slice(0, 70);
    }
  }
  const rowAttr = repRow ? `data-row="${repRow}"` : '';
  const sel = (SELECTED_ROW === repRow) ? ' selected' : '';

  out.push(`<div class="tree-node${expanded?' expanded':''}" data-key="${escapeHtml(node.key)}" style="padding-left:${indent}px">`);
  out.push(`<div class="tree-row${sel}" ${rowAttr} onclick="onTreeRowClick(event, '${escapeHtml(node.key)}', ${repRow||'null'})">`);
  if (hasKids && node.children.length) {
    out.push(`<span class="tree-chev" onclick="event.stopPropagation();toggleNode('${escapeHtml(node.key)}')">${expanded?'▼':'▶'}</span>`);
  } else {
    out.push(`<span class="tree-chev empty">·</span>`);
  }
  out.push(`<span class="tree-icon ${iconCls}">${icon}</span>`);
  out.push(`<span class="tree-label">${escapeHtml(label)}</span>`);
  if (stIcon) out.push(`<span class="tree-status" style="color:${stCol}">${stIcon}</span>`);
  // Confidence dot — at-a-glance for which rows still need review
  out.push(confDotHtml(r));
  out.push(flagHtml);
  out.push('</div>');

  // Additional rows in this node beyond the first → render as siblings
  for (let i = 1; i < node.rows.length; i++) {
    const rn = node.rows[i];
    if (!VISIBLE_ROWS.has(rn)) continue;
    const rr = ROWS_BY_NUM[rn];
    if (!rr) continue;
    const sel2 = (SELECTED_ROW === rn) ? ' selected' : '';
    const st2 = (DATA.status && DATA.status[rn]) ? DATA.status[rn].status : 'unverified';
    const ic2 = {pass:'✓', fail:'✗', need_fix:'⚠', skip:'⏭', unverified:''}[st2] || '';
    const sc2 = {pass:'#10b981', fail:'#ef4444', need_fix:'#f59e0b', skip:'#6b7280', unverified:''}[st2] || '';
    const fc2 = rr.auto_flags ? rr.auto_flags.length : 0;
    const fh2 = fc2 ? `<span class="tree-flags" title="${escapeHtml(rr.auto_flags.map(f=>f.msg).join(' / '))}" style="color:#d97706">⚠${fc2}</span>` : '';
    const lab2 = (rr.B || rr.C || '').toString().trim().replace(/\s+/g,' ').slice(0, 70);
    out.push(`<div class="tree-row${sel2}" data-row="${rn}" onclick="onTreeRowClick(event, null, ${rn})" style="padding-left:${(depth+1)*12+18}px">`);
    out.push(`<span class="tree-chev empty">·</span><span class="tree-icon">·</span>`);
    out.push(`<span class="tree-label">R${rn} ${escapeHtml(lab2)}</span>`);
    if (ic2) out.push(`<span class="tree-status" style="color:${sc2}">${ic2}</span>`);
    out.push(fh2);
    out.push('</div>');
  }

  out.push('<div class="tree-children">');
  for (const c of node.children) renderNode(c, depth + 1, out);
  out.push('</div></div>');
}

function toggleNode(key) {
  if (EXPANDED.has(key)) EXPANDED.delete(key);
  else EXPANDED.add(key);
  renderTree();
}
function expandAll() {
  function walk(n){ EXPANDED.add(n.key); for (const c of n.children) walk(c); }
  for (const c of TREE_ROOT.children) walk(c);
  renderTree();
}
function collapseAll() {
  EXPANDED.clear();
  for (const c of TREE_ROOT.children) EXPANDED.add(c.key);
  renderTree();
}

function onTreeRowClick(e, nodeKey, rowNum) {
  if (rowNum != null) selectRow(rowNum);
  // also expand the clicked node so its children are visible
  if (nodeKey) {
    EXPANDED.add(nodeKey);
    renderTree();
  }
}

// ── Selection: drives all 4 panes ──────────────────────────────
async function selectRow(rowNum, scroll = true) {
  SELECTED_ROW = rowNum;
  localStorage.setItem('lastRow', rowNum);
  const r = ROWS_BY_NUM[rowNum];
  if (!r) return;

  // 1) action bar
  renderActionBar(r);

  // 2) TOR
  loadTOR(rowNum);

  // 3) xlsx preview
  loadXlsx(rowNum);

  // 4) catalog PDF
  if (r.pdf_rel) {
    loadPdf(r.pdf_rel, r.parsed.page || 1, computeHighlight(r));
  } else {
    CURRENT_PDF = null;
    document.getElementById('pdf-canvas').innerHTML =
      `<div class="empty-canvas">Row นี้ไม่มี catalog อ้างอิง<br>(${r.parsed.type || 'empty'})</div>`;
    document.getElementById('pdf-filename').textContent = '(ไม่มี)';
    document.getElementById('pdf-page-info').textContent = '— / —';
    document.getElementById('pdf-annots').innerHTML = '';
  }

  // 5) update tree highlight + auto-expand path + scroll into view
  expandPathToRow(rowNum);
  renderTree();
  if (scroll) {
    requestAnimationFrame(() => {
      const el = document.querySelector(`.tree-row[data-row="${rowNum}"]`);
      if (el) el.scrollIntoView({ block: 'center', behavior: 'smooth' });
    });
  }
}

/* Walk the tree, find any node containing the row, expand all ancestors. */
function expandPathToRow(rowNum) {
  const path = [];
  function walk(node, ancestors) {
    if (node.rows && node.rows.indexOf(rowNum) >= 0) {
      for (const a of ancestors) path.push(a.key);
      return true;
    }
    for (const c of (node.children||[])) {
      if (walk(c, ancestors.concat([node]))) return true;
    }
    return false;
  }
  walk(TREE_ROOT, []);
  for (const k of path) if (k && k !== 'ROOT') EXPANDED.add(k);
}

// ── Action bar ─────────────────────────────────────────────────
function renderActionBar(r) {
  document.getElementById('action-bar').style.display = '';
  const cur = (DATA.status && DATA.status[r.row]) || {};
  const st = cur.status || 'unverified';

  let secLabel = r.section || '';
  if (r.parsed.brand || r.parsed.model) {
    secLabel += ` · ${r.parsed.brand||''} ${r.parsed.model||''}`.trim();
  }
  document.getElementById('ab-row-info').innerHTML =
    `<span class="row-num">R${r.row}</span> <span class="section">${escapeHtml(secLabel)}</span> ${r.E ? `<span class="vendor-tag ${r.E}">${escapeHtml(r.E)}</span>` : ''}`;

  const flags = r.auto_flags || [];
  document.getElementById('ab-flags').innerHTML = flags.length
    ? flags.map(f => `<span title="${escapeHtml(f.msg)}">⚠ ${escapeHtml(f.msg.slice(0,60))}</span>`).join(' · ')
    : '';

  for (const cls of ['pass','fail','fix','skip']) {
    const btn = document.querySelector('.ab-btn.' + cls);
    if (btn) btn.classList.toggle('active', st === (cls==='fix'?'need_fix':cls));
  }
  document.getElementById('ab-notes').value = cur.notes || '';

  // Pulse the 📍 Mark button when Col D is "ยินดีปฏิบัติ" (commitment) —
  // the most common case where the user might want to manually locate
  // the spec the AI failed to find.
  const markBtn = document.getElementById('ab-mark-btn');
  if (markBtn) {
    const dCommit = (r.D || '').trim().startsWith('ยินดีปฏิบัติ');
    markBtn.classList.toggle('commitment', dCommit && !!r.pdf_rel);
    markBtn.disabled = !r.pdf_rel;
    markBtn.title = !r.pdf_rel
      ? 'ไม่มี catalog PDF ให้ annotate'
      : (dCommit ? '⚠ Col D เป็น commitment — คลิกเพื่อหาเนื้อหาใน catalog เอง'
                 : 'ลาก rect ใน catalog เพื่อ override Col D');
  }
}

let notesTimer = null;
function saveNotesDebounced() { clearTimeout(notesTimer); notesTimer = setTimeout(saveNotes, 400); }
async function saveNotes() {
  if (!SELECTED_ROW) return;
  const notes = document.getElementById('ab-notes').value;
  await fetch('/api/status', {method: 'POST', headers: {'Content-Type':'application/json'},
    body: JSON.stringify({row: SELECTED_ROW, notes})});
  if (!DATA.status) DATA.status = {};
  DATA.status[SELECTED_ROW] = DATA.status[SELECTED_ROW] || {};
  DATA.status[SELECTED_ROW].notes = notes;
}
async function setStatus(status) {
  if (!SELECTED_ROW) return;
  const notes = document.getElementById('ab-notes').value || '';
  const r = await fetch('/api/status', {method: 'POST', headers: {'Content-Type':'application/json'},
    body: JSON.stringify({row: SELECTED_ROW, status, notes})});
  const j = await r.json();
  if (!DATA.status) DATA.status = {};
  DATA.status[SELECTED_ROW] = j.entry;
  renderStats();
  renderTree();
  const row = ROWS_BY_NUM[SELECTED_ROW];
  if (row) renderActionBar(row);
  // Verdict counts as a learning signal — accumulate; auto-retrain at threshold
  if (status && status !== 'unverified') {
    if (typeof tickRetrain === 'function') tickRetrain('status');
  }
}

// ── TOR pane ───────────────────────────────────────────────────
async function loadTOR(rowNum) {
  const url = `/api/tor_page?row=${rowNum}&dpi=${TOR_DPI}`;
  const c = document.getElementById('tor-canvas');
  c.innerHTML = '<div class="empty-canvas">กำลังโหลด TOR…</div>';
  try {
    const r = await fetch(url);
    if (!r.ok) {
      c.innerHTML = '<div class="empty-canvas">โหลด TOR ไม่ได้</div>';
      return;
    }
    TOR_TARGET_PAGE = parseInt(r.headers.get('X-TOR-Page') || '1');
    TOR_HITS = parseInt(r.headers.get('X-TOR-Hits') || '0');
    const hY0 = parseFloat(r.headers.get('X-Highlight-Y0'));
    const hY1 = parseFloat(r.headers.get('X-Highlight-Y1'));
    const torSection = r.headers.get('X-Section') || '';
    const torRange = r.headers.get('X-Section-Range') || '';
    TOR_PAGE = TOR_TARGET_PAGE;
    const blob = await r.blob();
    const imgUrl = URL.createObjectURL(blob);
    showImageWithScroll(c, imgUrl, hY0, hY1, `TOR p${TOR_PAGE}`);
    if (!TOR_PAGES) {
      const m = await fetch('/api/tor_meta').then(r => r.json());
      TOR_PAGES = m.pages || 0;
      document.getElementById('tor-info').textContent = m.filename || '';
    }
    document.getElementById('tor-page-info').textContent = `${TOR_PAGE} / ${TOR_PAGES}`;
    let status = '';
    if (TOR_HITS > 0) {
      status = `✓ ${TOR_HITS} hit${TOR_HITS>1?'s':''}`;
    } else if (torRange) {
      status = `ไม่เจอข้อความ — แสดงต้น section ${torSection}`;
    } else {
      status = 'ไม่เจอข้อความ';
    }
    if (torSection && torRange) status += `  ·  section ${torSection} P${torRange}`;
    document.getElementById('tor-status').textContent = status;
  } catch (e) {
    c.innerHTML = '<div class="empty-canvas">โหลด TOR error</div>';
  }
}
async function torReload() {
  if (!SELECTED_ROW) return;
  const c = document.getElementById('tor-canvas');
  const url = `/api/tor_page?row=${SELECTED_ROW}&page=${TOR_PAGE}&dpi=${TOR_DPI}`;
  const r = await fetch(url);
  const hY0 = parseFloat(r.headers.get('X-Highlight-Y0'));
  const hY1 = parseFloat(r.headers.get('X-Highlight-Y1'));
  const blob = await r.blob();
  showImageWithScroll(c, URL.createObjectURL(blob), hY0, hY1, `TOR p${TOR_PAGE}`);
  document.getElementById('tor-page-info').textContent = `${TOR_PAGE} / ${TOR_PAGES}`;
}

/* Render an image into a scrollable container and auto-scroll so that the
   highlight Y range (in source-pixel coords) sits ~25% from the top of the
   visible viewport. Falls back to top if no Y given. */
function showImageWithScroll(container, src, y0, y1, alt) {
  container.innerHTML = `<img src="${src}" alt="${alt||''}">`;
  const img = container.querySelector('img');
  img.onload = () => {
    if (!isFinite(y0)) { container.scrollTop = 0; return; }
    // Source-pixel → rendered-pixel scale (image may be max-width: 100%)
    const scale = img.clientWidth / img.naturalWidth;
    const targetMid = ((y0 + (isFinite(y1) ? y1 : y0)) / 2) * scale;
    // place highlight ~25% from top of visible area
    const offset = container.clientHeight * 0.25;
    container.scrollTop = Math.max(0, targetMid - offset);
  };
}
function torPrev() { if (TOR_PAGE > 1) { TOR_PAGE--; torReload(); } }
function torNext() { if (TOR_PAGE < TOR_PAGES) { TOR_PAGE++; torReload(); } }
function torZoom(d) { TOR_DPI = Math.max(70, Math.min(200, TOR_DPI + d*20)); torReload(); }
function torJumpToMatch() { TOR_PAGE = TOR_TARGET_PAGE; torReload(); }

// ── XLSX preview ───────────────────────────────────────────────
async function loadXlsx(rowNum) {
  const w = document.getElementById('xlsx-wrap');
  const r = await fetch(`/api/row_context?row=${rowNum}&radius=${CTX_RADIUS}`);
  const j = await r.json();
  document.getElementById('ctx-radius').textContent = CTX_RADIUS;
  let html = `<table class="xlsx-table">
    <thead><tr>
      <th class="row-num">#</th>
      <th class="col-A">A</th>
      <th class="col-B">B คุณลักษณะที่ต้องการ (TOR)</th>
      <th class="col-C">C คุณลักษณะที่เสนอ</th>
      <th class="col-D">D เอกสารอ้างอิง</th>
      <th class="col-E">E</th>
      <th class="col-F">F สถานะ TOR</th>
    </tr></thead><tbody>`;
  for (const row of j.rows) {
    const cls = row.row === j.target ? 'target' : '';
    html += `<tr class="${cls}" data-row="${row.row}" onclick="selectRow(${row.row})">`;
    html += `<td class="row-num">${row.row}</td>`;
    html += `<td class="col-A">${escapeHtml(row.A||'')}</td>`;
    html += `<td class="col-B">${escapeHtml(row.B||'')}</td>`;
    html += `<td class="col-C">${escapeHtml(row.C||'')}</td>`;
    // Col D — single-click for dropdown (annotate / revert / edit / auto),
    // double-click for inline editing
    const dVal = row.D || '';
    const dCommit = dVal.trim().startsWith('ยินดีปฏิบัติ');
    const dCls = 'col-D editable' + (dCommit ? ' commitment' : ' has-ref');
    html += `<td class="${dCls}" onclick="onColDClick(event, ${row.row})" ondblclick="editColD(event, ${row.row})" title="คลิกเพื่อเปิดเมนู / ดับเบิลคลิกเพื่อแก้">`;
    html += `<span class="d-text">${escapeHtml(dVal)}</span>`;
    html += `<span class="d-caret" aria-hidden="true">▾</span>`;
    html += `</td>`;
    html += `<td class="col-E">${row.E ? `<span class="vendor-tag ${row.E}">${escapeHtml(row.E)}</span>` : ''}</td>`;
    html += `<td class="col-F">${escapeHtml(row.F||'')}</td>`;
    html += `</tr>`;
  }
  html += '</tbody></table>';
  w.innerHTML = html;
  // scroll target into view
  setTimeout(() => {
    const t = w.querySelector('tr.target');
    if (t) t.scrollIntoView({block: 'center'});
  }, 0);
  document.getElementById('xlsx-info').textContent = `R${j.target}`;
}
function ctxRadius(d) {
  CTX_RADIUS = Math.max(2, Math.min(30, CTX_RADIUS + d));
  if (SELECTED_ROW) loadXlsx(SELECTED_ROW);
}

// Inline Col D editing — double-click in xlsx preview → editable textarea →
// save → records HITL feedback
function editColD(e, rowNum) {
  e.stopPropagation();
  const td = e.currentTarget;
  const original = td.textContent;
  td.contentEditable = 'plaintext-only';
  td.classList.add('editing');
  td.focus();
  const sel = window.getSelection();
  sel.selectAllChildren(td);

  const finish = async (commit) => {
    td.contentEditable = 'false';
    td.classList.remove('editing');
    const newVal = td.textContent.trim();
    if (!commit || newVal === original.trim()) {
      td.textContent = original;
      return;
    }
    // Send to server
    try {
      const r = await fetch('/api/row/col_d', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({row: rowNum, col_d: newVal, original}),
      });
      const j = await r.json();
      if (j.ok) {
        toast(`✓ R${rowNum} Col D updated`,
              `${original.slice(0, 40)} → ${newVal.slice(0, 40)}`,
              'learn', 4000);
        // Trigger auto-retrain since this is a strong correction signal
        if (typeof tickRetrain === 'function') tickRetrain('inline');
        // Reload xlsx + tree
        if (SELECTED_ROW) loadXlsx(SELECTED_ROW);
        // Refresh in-memory rows snapshot
        const idx = await fetch('/api/index').then(r => r.json());
        DATA = idx;
        ROWS_BY_NUM = Object.fromEntries(idx.rows.map(r => [r.row, r]));
        renderTree();
      } else {
        toast('Save failed', j.error || 'unknown', 'error', 5000);
        td.textContent = original;
      }
    } catch (err) {
      toast('Save failed', err.message, 'error', 5000);
      td.textContent = original;
    }
  };

  td.addEventListener('blur', () => finish(true), {once: true});
  td.addEventListener('keydown', (ev) => {
    if (ev.key === 'Enter' && !ev.shiftKey) { ev.preventDefault(); td.blur(); }
    if (ev.key === 'Escape') { ev.preventDefault(); finish(false); td.blur(); }
  });
}

// ── Col D context dropdown menu ───────────────────────────────
// Single-click on a Col D cell shows a dropdown of actions tailored to
// the row's current state:
//   • commitment ("ยินดีปฏิบัติ") → Mark / Auto / Edit
//   • has reference  → Revert / Re-mark / Re-auto / Edit
//
// The double-click handler (editColD) still runs for power users.

let _COL_D_MENU = null;

function closeColDMenu() {
  if (_COL_D_MENU) { _COL_D_MENU.remove(); _COL_D_MENU = null; }
  document.removeEventListener('click', _onDocClickForColDMenu, true);
  document.removeEventListener('keydown', _onEscForColDMenu, true);
}
function _onDocClickForColDMenu(e) {
  if (!_COL_D_MENU) return;
  if (!_COL_D_MENU.contains(e.target)) closeColDMenu();
}
function _onEscForColDMenu(e) {
  if (e.key === 'Escape') closeColDMenu();
}

function onColDClick(e, rowNum) {
  e.stopPropagation();   // don't double-trigger row select
  if (e.target.closest('td').classList.contains('editing')) return;  // inline-edit mode, skip
  selectRow(rowNum);     // make sure the row context is loaded
  showColDMenu(e, rowNum);
}

function showColDMenu(e, rowNum) {
  closeColDMenu();
  const row = ROWS_BY_NUM[rowNum];
  if (!row) return;
  const dVal = (row.D || '').trim();
  const isCommit = dVal.startsWith('ยินดีปฏิบัติ');
  const hasPdf = !!row.pdf_rel;
  const sec = row.section || '?';

  const menu = document.createElement('div');
  menu.className = 'col-d-menu';
  let html = `<div class="menu-header">R${rowNum} · ${escapeHtml(sec)} · ${isCommit ? 'commitment' : 'has reference'}</div>`;

  if (isCommit) {
    // → Switch FROM commitment TO real annotation
    html += `<button class="primary" data-act="mark" ${!hasPdf?'disabled':''}>
      <span class="icon">📍</span>
      <span class="label">Mark in catalog → annotate</span>
      <span class="hint">${hasPdf?'รับ rect + label':'no PDF'}</span>
    </button>`;
    html += `<button data-act="auto" ${!hasPdf?'disabled':''}>
      <span class="icon">✨</span>
      <span class="label">Auto-annotate (AI tries again)</span>
      <span class="hint">preview ก่อน apply</span>
    </button>`;
    html += `<div class="sep"></div>`;
    html += `<button data-act="edit">
      <span class="icon">✏</span>
      <span class="label">Edit Col D manually</span>
      <span class="hint">double-click ก็ได้</span>
    </button>`;
  } else {
    // → Switch FROM annotation TO commitment, or modify
    html += `<button data-act="auto">
      <span class="icon">✨</span>
      <span class="label">Re-run auto-annotate</span>
      <span class="hint">เริ่มใหม่</span>
    </button>`;
    html += `<button data-act="mark" ${!hasPdf?'disabled':''}>
      <span class="icon">📍</span>
      <span class="label">Re-mark in catalog</span>
      <span class="hint">วาด rect ใหม่</span>
    </button>`;
    html += `<button data-act="edit">
      <span class="icon">✏</span>
      <span class="label">Edit Col D manually</span>
    </button>`;
    html += `<div class="sep"></div>`;
    html += `<button class="danger" data-act="revert">
      <span class="icon">↩</span>
      <span class="label">Revert to "ยินดีปฏิบัติตามข้อกำหนด"</span>
      <span class="hint">บันทึกประวัติก่อน</span>
    </button>`;
  }

  menu.innerHTML = html;
  document.body.appendChild(menu);
  _COL_D_MENU = menu;

  // Position near the click, but keep on-screen
  const W = menu.offsetWidth, H = menu.offsetHeight;
  let x = e.clientX + 4, y = e.clientY + 4;
  if (x + W > window.innerWidth - 8)  x = window.innerWidth - W - 8;
  if (y + H > window.innerHeight - 8) y = e.clientY - H - 4;
  menu.style.left = `${Math.max(8, x)}px`;
  menu.style.top  = `${Math.max(8, y)}px`;

  // Wire button actions
  menu.querySelectorAll('button[data-act]').forEach(btn => {
    btn.addEventListener('click', (ev) => {
      ev.stopPropagation();
      const act = btn.dataset.act;
      closeColDMenu();
      handleColDAction(act, rowNum);
    });
  });

  setTimeout(() => {
    document.addEventListener('click', _onDocClickForColDMenu, true);
    document.addEventListener('keydown', _onEscForColDMenu, true);
  }, 0);
}

async function handleColDAction(action, rowNum) {
  const row = ROWS_BY_NUM[rowNum];
  if (!row) return;
  switch (action) {
    case 'mark':
      startManualAnnotate();
      break;
    case 'auto':
      showAutoAnnotate();
      break;
    case 'edit':
      // Programmatically focus the Col D cell and put it in edit mode
      const td = document.querySelector(`tr[data-row="${rowNum}"] td.col-D`);
      if (td) {
        const fakeEvent = {currentTarget: td, stopPropagation: () => {}};
        editColD(fakeEvent, rowNum);
      }
      break;
    case 'revert':
      await revertColDToCommitment(rowNum);
      break;
  }
}

async function revertColDToCommitment(rowNum) {
  const row = ROWS_BY_NUM[rowNum];
  if (!row) return;
  const old = row.D || '';
  if (old.trim().startsWith('ยินดีปฏิบัติ')) {
    toast('Already commitment', 'Col D เป็น "ยินดีปฏิบัติ" อยู่แล้ว', 'info', 2500);
    return;
  }
  if (!confirm(`Revert Col D ของ R${rowNum} เป็น "ยินดีปฏิบัติตามข้อกำหนด"?\n\n` +
               `เดิม: ${old.slice(0, 100)}\n\n` +
               `(ระบบ snapshot ก่อน — annotations ใน PDF ไม่ลบ)`)) return;
  try {
    const r = await fetch('/api/row/col_d', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({
        row: rowNum,
        col_d: 'ยินดีปฏิบัติตามข้อกำหนด',
        original: old,
      }),
    });
    const j = await r.json();
    if (j.ok) {
      toast(`↩ R${rowNum} reverted`, 'Col D = ยินดีปฏิบัติตามข้อกำหนด', 'learn', 4000);
      if (typeof tickRetrain === 'function') tickRetrain('revert');
      // Reload xlsx + tree
      if (SELECTED_ROW) loadXlsx(SELECTED_ROW);
      const idx = await fetch('/api/index').then(r => r.json());
      DATA = idx;
      ROWS_BY_NUM = Object.fromEntries(idx.rows.map(r => [r.row, r]));
      renderTree();
    } else {
      toast('Revert failed', j.error || 'unknown', 'error', 5000);
    }
  } catch (err) {
    toast('Revert failed', err.message, 'error', 5000);
  }
}

// ── Catalog PDF ────────────────────────────────────────────────
function computeHighlight(r) {
  const p = r.parsed;
  if (p.type === 'brand_model') return 'ยี่ห้อ|รุ่น';
  // Trailing period is what real annotation labels use ("ข้อย่อย 1.").
  // It's redundant for the structured matcher but safer if older annotations
  // ever fall back to substring matching.
  if (p.item != null && p.subitem != null) return `ข้อ ${p.item}) ข้อย่อย ${p.subitem}.`;
  if (p.subitem != null) return `ข้อย่อย ${p.subitem}.`;
  if (p.item != null) return `ข้อ ${p.item})`;
  return null;
}
async function loadPdf(rel, page, highlight) {
  CURRENT_PDF = {rel, meta: null};
  CURRENT_HIGHLIGHT = highlight;
  PDF_PAGE = page;
  document.getElementById('pdf-filename').textContent = rel;
  const r = await fetch('/api/pdf_meta?rel=' + encodeURIComponent(rel));
  if (!r.ok) {
    document.getElementById('pdf-canvas').innerHTML = '<div class="empty-canvas">โหลด PDF ไม่ได้</div>';
    return;
  }
  CURRENT_PDF.meta = await r.json();
  if (PDF_PAGE > CURRENT_PDF.meta.pages) PDF_PAGE = CURRENT_PDF.meta.pages;
  if (PDF_PAGE < 1) PDF_PAGE = 1;
  // If edit mode is on, re-initialize the working set with the new PDF's annots
  if (EDIT_MODE) {
    initEditAnnots();
  } else {
    EDIT_ANNOTS = [];
    ORIGINAL_ANNOTS_BY_ID = {};
    UNDO_STACK = [];
    REDO_STACK = [];
    SELECTED_ANN_ID = null;
  }
  renderPdf();
}
async function renderPdf() {
  if (!CURRENT_PDF) return;
  const useHl = !EDIT_MODE && document.getElementById('hl-toggle').checked ? CURRENT_HIGHLIGHT : null;
  let url = '/api/pdf_page?rel=' + encodeURIComponent(CURRENT_PDF.rel)
          + '&page=' + PDF_PAGE + '&dpi=' + PDF_DPI;
  if (EDIT_MODE) url += '&edit=1&_t=' + Date.now();
  if (useHl) url += '&highlight=' + encodeURIComponent(useHl);
  const c = document.getElementById('pdf-canvas');
  try {
    const r = await fetch(url);
    if (!r.ok) {
      c.innerHTML = '<div class="empty-canvas">โหลด PDF ไม่ได้</div>';
      return;
    }
    const hY0 = parseFloat(r.headers.get('X-Highlight-Y0'));
    const hY1 = parseFloat(r.headers.get('X-Highlight-Y1'));
    const blob = await r.blob();
    const imgUrl = URL.createObjectURL(blob);
    if (EDIT_MODE) {
      showPdfPageWithOverlay(c, imgUrl, hY0, hY1);
    } else {
      showImageWithScroll(c, imgUrl, hY0, hY1, `page ${PDF_PAGE}`);
    }
  } catch (e) {
    c.innerHTML = '<div class="empty-canvas">PDF render error</div>';
    return;
  }
  document.getElementById('pdf-page-info').textContent = `${PDF_PAGE} / ${CURRENT_PDF.meta.pages}`;
  renderAnnots();
}
// Structured matcher mirroring the Python _match_annot_label (handles
// "ข้อย่อย 1" ≠ "ข้อย่อย 10" and other digit-boundary cases).
const _LBL_ITEM_RE    = /ข้อ\s*(\d+)\s*\)/;
const _LBL_SUBITEM_RE = /ข้อย่อย\s*(\d+)/;
const _LBL_LITERALS   = new Set(['ยี่ห้อ', 'รุ่น']);
function _parseLabel(s) {
  if (!s) return {};
  s = s.trim();
  if (_LBL_LITERALS.has(s)) return {kind: 'literal', literal: s};
  const out = {};
  const i = s.match(_LBL_ITEM_RE);
  const sub = s.match(_LBL_SUBITEM_RE);
  if (i) out.item = parseInt(i[1]);
  if (sub) out.subitem = parseInt(sub[1]);
  if (Object.keys(out).length) out.kind = 'section_label';
  return out;
}
function matchAnnotLabel(query, content) {
  if (!query || !content) return false;
  const qf = _parseLabel(query);
  const cf = _parseLabel(content);
  if (qf.kind === 'literal') return cf.kind === 'literal' && cf.literal === qf.literal;
  if (qf.kind !== 'section_label') {
    if (query === content) return true;
    // word-boundary fallback: query not followed by another digit
    const re = new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(?!\\d)');
    return re.test(content);
  }
  if (cf.kind !== 'section_label') return false;
  if (qf.subitem != null) {
    if (cf.subitem == null) {
      return qf.item != null && cf.item === qf.item;
    }
    if (cf.subitem !== qf.subitem) return false;
    if (qf.item != null && cf.item != null && cf.item !== qf.item) return false;
    return true;
  }
  if (qf.item != null) {
    return cf.item === qf.item && cf.subitem == null;
  }
  return false;
}

function renderAnnots() {
  const a = document.getElementById('pdf-annots');
  if (!CURRENT_PDF || !CURRENT_PDF.meta) { a.innerHTML = ''; return; }
  const annots = EDIT_MODE ? EDIT_ANNOTS.filter(x => !x._deleted) : CURRENT_PDF.meta.annots;
  const cur = annots.filter(x => x.page === PDF_PAGE);
  if (!cur.length) { a.innerHTML = '<div style="color:#9ca3af">(ไม่มี annotation บนหน้านี้)</div>'; return; }
  const queries = (CURRENT_HIGHLIGHT || '').split('|').map(s=>s.trim()).filter(Boolean);
  a.innerHTML = cur.map(x => {
    const m = queries.some(q => matchAnnotLabel(q, x.contents || ''));
    const sel = (x._id != null && x._id === SELECTED_ANN_ID) ? ' selected' : '';
    return `<div class="ann-row ${m?'matched':''}${sel}"><b>${x.type}</b> ${x.contents ? escapeHtml(x.contents) : '<i style="color:#9ca3af">(empty)</i>'}</div>`;
  }).join('');
}
function pdfPrev() { if (CURRENT_PDF && PDF_PAGE > 1) { PDF_PAGE--; renderPdf(); } }
function pdfNext() { if (CURRENT_PDF && PDF_PAGE < CURRENT_PDF.meta.pages) { PDF_PAGE++; renderPdf(); } }
function pdfZoom(d) { PDF_DPI = Math.max(70, Math.min(260, PDF_DPI + d*25)); renderPdf(); }
function openInBrowser() {
  if (!CURRENT_PDF) return;
  window.open('/api/raw_pdf?rel=' + encodeURIComponent(CURRENT_PDF.rel) + '#page=' + PDF_PAGE, '_blank');
}

// ── Edit-mode state + SVG overlay ──────────────────────────────
let EDIT_MODE = false;
let EDIT_TOOL = 'select';
let EDIT_ANNOTS = [];                  // working copy with _id, _deleted, _isNew flags
let ORIGINAL_ANNOTS_BY_ID = {};        // _id → original snapshot (for diff at save)
let SELECTED_ANN_ID = null;
let UNDO_STACK = [];
let REDO_STACK = [];
let DIRTY = false;
let NEW_ID_COUNTER = 0;
let DRAG_STATE = null;                 // {mode, ann, start, origRect}

function newClientId() { NEW_ID_COUNTER++; return 'new-' + NEW_ID_COUNTER; }

function snapshotAnnots() {
  return EDIT_ANNOTS.map(a => ({...a}));
}
function commitChange() {
  UNDO_STACK.push(snapshotAnnots());
  if (UNDO_STACK.length > 200) UNDO_STACK.shift();
  REDO_STACK = [];
  setDirty(true);
  refreshOverlay();
  refreshUndoRedoButtons();
}
function undo() {
  if (UNDO_STACK.length === 0) return;
  REDO_STACK.push(snapshotAnnots());
  EDIT_ANNOTS = UNDO_STACK.pop();
  if (!EDIT_ANNOTS.find(a => a._id === SELECTED_ANN_ID)) SELECTED_ANN_ID = null;
  setDirty(hasUnsavedChanges());
  refreshOverlay();
  refreshUndoRedoButtons();
}
function redo() {
  if (REDO_STACK.length === 0) return;
  UNDO_STACK.push(snapshotAnnots());
  EDIT_ANNOTS = REDO_STACK.pop();
  setDirty(hasUnsavedChanges());
  refreshOverlay();
  refreshUndoRedoButtons();
}
function hasUnsavedChanges() {
  for (const a of EDIT_ANNOTS) {
    if (a._isNew && !a._deleted) return true;
    if (a._deleted && !a._isNew) return true;
    if (!a._isNew) {
      const o = ORIGINAL_ANNOTS_BY_ID[a._id];
      if (!o) continue;
      if (a.contents !== o.contents) return true;
      const r1 = a.rect, r2 = o.rect;
      if (Math.abs(r1[0]-r2[0])>0.5 || Math.abs(r1[1]-r2[1])>0.5 ||
          Math.abs(r1[2]-r2[2])>0.5 || Math.abs(r1[3]-r2[3])>0.5) return true;
    }
  }
  return false;
}
function setDirty(v) {
  DIRTY = v;
  document.getElementById('save-btn').disabled = !v;
  document.getElementById('dirty-ind').textContent = v ? '● ยังไม่บันทึก' : '';
}
function refreshUndoRedoButtons() {
  document.getElementById('undo-btn').disabled = UNDO_STACK.length === 0;
  document.getElementById('redo-btn').disabled = REDO_STACK.length === 0;
}

function toggleEditMode() {
  if (!CURRENT_PDF) return;
  if (EDIT_MODE && DIRTY) {
    if (!confirm('มีการแก้ไขที่ยังไม่บันทึก — ออกจาก edit mode จะทิ้งงาน?')) return;
  }
  EDIT_MODE = !EDIT_MODE;
  document.getElementById('edit-toggle-btn').classList.toggle('active', EDIT_MODE);
  document.getElementById('edit-toggle-btn').textContent = EDIT_MODE ? '👁 View' : '✏ Edit';
  document.getElementById('edit-toolbar').style.display = EDIT_MODE ? '' : 'none';
  const c = document.getElementById('pdf-canvas');
  c.classList.toggle('edit-mode', EDIT_MODE);
  if (EDIT_MODE) {
    initEditAnnots();
    setTool('select');
  } else {
    closeTextEditor();
    SELECTED_ANN_ID = null;
  }
  renderPdf();
}

function initEditAnnots() {
  EDIT_ANNOTS = [];
  ORIGINAL_ANNOTS_BY_ID = {};
  for (const a of (CURRENT_PDF.meta.annots || [])) {
    const id = 'x' + a.xref;
    const copy = {
      _id: id,
      xref: a.xref,
      page: a.page,
      type: a.type,
      rect: a.rect.slice(),
      contents: a.contents,
    };
    EDIT_ANNOTS.push(copy);
    ORIGINAL_ANNOTS_BY_ID[id] = {
      contents: a.contents,
      rect: a.rect.slice(),
    };
  }
  UNDO_STACK = [];
  REDO_STACK = [];
  SELECTED_ANN_ID = null;
  setDirty(false);
  refreshUndoRedoButtons();
}

function setTool(name) {
  EDIT_TOOL = name;
  for (const b of document.querySelectorAll('.edit-toolbar button.tool')) {
    b.classList.toggle('active', b.dataset.tool === name);
  }
  const c = document.getElementById('pdf-canvas');
  c.classList.remove('tool-drawRect', 'tool-addText');
  if (name === 'drawRect') c.classList.add('tool-drawRect');
  else if (name === 'addText') c.classList.add('tool-addText');
}

// SVG overlay rendering
let OVERLAY_HOST = null;
let OVERLAY_SVG = null;

function showPdfPageWithOverlay(container, src, y0, y1) {
  container.innerHTML = '';
  const host = document.createElement('div');
  host.className = 'pdf-page-host';
  const img = document.createElement('img');
  img.className = 'pdf-page-img';
  img.src = src;
  host.appendChild(img);
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.classList.add('pdf-overlay', 'editable');
  host.appendChild(svg);
  container.appendChild(host);
  OVERLAY_HOST = host;
  OVERLAY_SVG = svg;
  img.onload = () => {
    const pageSize = CURRENT_PDF.meta.page_sizes[PDF_PAGE - 1] || [img.naturalWidth, img.naturalHeight];
    svg.setAttribute('viewBox', `0 0 ${pageSize[0]} ${pageSize[1]}`);
    svg.style.width = '100%';
    svg.style.height = '100%';
    refreshOverlay();
    // auto-scroll
    if (isFinite(y0)) {
      const scale = img.clientWidth / img.naturalWidth;
      const targetMid = ((y0 + (isFinite(y1) ? y1 : y0)) / 2) * scale;
      container.scrollTop = Math.max(0, targetMid - container.clientHeight * 0.25);
    }
  };
  // Drag-to-create / click handlers on the SVG
  svg.addEventListener('pointerdown', onOverlayPointerDown);
}

function refreshOverlay() {
  if (!OVERLAY_SVG) return;
  while (OVERLAY_SVG.firstChild) OVERLAY_SVG.removeChild(OVERLAY_SVG.firstChild);
  const onPage = EDIT_ANNOTS.filter(a => a.page === PDF_PAGE && !a._deleted);
  for (const a of onPage) {
    OVERLAY_SVG.appendChild(buildAnnotNode(a));
  }
  if (SELECTED_ANN_ID != null) {
    const sel = onPage.find(a => a._id === SELECTED_ANN_ID);
    if (sel) OVERLAY_SVG.appendChild(buildHandles(sel));
  }
  renderAnnots();
}

function buildAnnotNode(a) {
  const NS = 'http://www.w3.org/2000/svg';
  const g = document.createElementNS(NS, 'g');
  g.classList.add('annot');
  g.dataset.id = a._id;
  if (a._id === SELECTED_ANN_ID) g.classList.add('selected');
  const [x0, y0, x1, y1] = a.rect;
  // Visible outline rect
  const rect = document.createElementNS(NS, 'rect');
  rect.classList.add('ann-rect');
  if (a.type === 'FreeText') rect.classList.add('freetext');
  rect.setAttribute('x', x0); rect.setAttribute('y', y0);
  rect.setAttribute('width', Math.max(0.5, x1 - x0));
  rect.setAttribute('height', Math.max(0.5, y1 - y0));
  g.appendChild(rect);
  // Hit area (slightly larger, transparent fill — enables click + hover)
  const hit = document.createElementNS(NS, 'rect');
  hit.classList.add('ann-hit');
  hit.setAttribute('x', x0); hit.setAttribute('y', y0);
  hit.setAttribute('width', Math.max(0.5, x1 - x0));
  hit.setAttribute('height', Math.max(0.5, y1 - y0));
  g.appendChild(hit);
  // FreeText label
  if (a.type === 'FreeText' && a.contents) {
    const fontSize = Math.min(12, Math.max(6, (y1 - y0) * 0.7));
    // Multi-line: split by \n
    const lines = a.contents.split('\n');
    for (let i = 0; i < lines.length; i++) {
      const t = document.createElementNS(NS, 'text');
      t.classList.add('ann-text');
      t.setAttribute('x', x0 + 2);
      t.setAttribute('y', y0 + fontSize * (i + 0.95));
      t.setAttribute('font-size', fontSize);
      t.textContent = lines[i];
      g.appendChild(t);
    }
  }
  return g;
}

function buildHandles(a) {
  const NS = 'http://www.w3.org/2000/svg';
  const g = document.createElementNS(NS, 'g');
  g.classList.add('handles');
  const [x0, y0, x1, y1] = a.rect;
  const cx = (x0 + x1) / 2, cy = (y0 + y1) / 2;
  const handles = [
    ['nw', x0, y0], ['n', cx, y0], ['ne', x1, y0],
    ['w', x0, cy], ['e', x1, cy],
    ['sw', x0, y1], ['s', cx, y1], ['se', x1, y1],
  ];
  const HSIZE = 4;  // PDF-pt; will be smaller on screen due to viewBox scaling
  for (const [name, hx, hy] of handles) {
    const h = document.createElementNS(NS, 'rect');
    h.classList.add('handle', 'h-' + name);
    h.setAttribute('x', hx - HSIZE/2);
    h.setAttribute('y', hy - HSIZE/2);
    h.setAttribute('width', HSIZE);
    h.setAttribute('height', HSIZE);
    h.dataset.handle = name;
    g.appendChild(h);
  }
  return g;
}

// Pointer event coords: screen → PDF-pt
function eventToPdfPt(e) {
  const svg = OVERLAY_SVG;
  const pt = svg.createSVGPoint();
  pt.x = e.clientX; pt.y = e.clientY;
  const ctm = svg.getScreenCTM();
  if (!ctm) return [0,0];
  const inv = ctm.inverse();
  const p = pt.matrixTransform(inv);
  return [p.x, p.y];
}

function onOverlayPointerDown(e) {
  if (!EDIT_MODE) return;
  closeTextEditor();
  // Resize handle?
  if (e.target.classList.contains('handle')) {
    const sel = EDIT_ANNOTS.find(a => a._id === SELECTED_ANN_ID);
    if (!sel) return;
    // Snapshot BEFORE the drag starts so undo can restore it
    UNDO_STACK.push(snapshotAnnots());
    if (UNDO_STACK.length > 200) UNDO_STACK.shift();
    REDO_STACK = [];
    DRAG_STATE = {
      mode: 'resize',
      handle: e.target.dataset.handle,
      ann: sel,
      origRect: sel.rect.slice(),
      start: eventToPdfPt(e),
    };
    e.target.setPointerCapture(e.pointerId);
    e.preventDefault(); e.stopPropagation();
    return;
  }
  // Click on annotation?
  const annNode = e.target.closest('g.annot');
  if (annNode) {
    const id = annNode.dataset.id;
    SELECTED_ANN_ID = id;
    const sel = EDIT_ANNOTS.find(a => a._id === id);
    if (!sel) return;
    if (EDIT_TOOL === 'select') {
      // Snapshot before move starts
      UNDO_STACK.push(snapshotAnnots());
      if (UNDO_STACK.length > 200) UNDO_STACK.shift();
      REDO_STACK = [];
      DRAG_STATE = {
        mode: 'move',
        ann: sel,
        origRect: sel.rect.slice(),
        start: eventToPdfPt(e),
      };
      OVERLAY_SVG.setPointerCapture(e.pointerId);
    }
    refreshOverlay();
    e.preventDefault(); e.stopPropagation();
    return;
  }
  // Click on empty area
  if (EDIT_TOOL === 'drawRect') {
    const [x, y] = eventToPdfPt(e);
    DRAG_STATE = {
      mode: 'drawRect',
      start: [x, y],
      cur: [x, y],
    };
    OVERLAY_SVG.setPointerCapture(e.pointerId);
    drawPreviewRect();
  } else if (EDIT_TOOL === 'addText') {
    const [x, y] = eventToPdfPt(e);
    addTextAt(x, y);
  } else {
    SELECTED_ANN_ID = null;
    refreshOverlay();
  }
  e.preventDefault();
}

function drawPreviewRect() {
  if (!DRAG_STATE || DRAG_STATE.mode !== 'drawRect') return;
  // Remove previous preview
  const prev = OVERLAY_SVG.querySelector('rect.draw-preview');
  if (prev) prev.remove();
  const [x0, y0] = DRAG_STATE.start, [x1, y1] = DRAG_STATE.cur;
  const r = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
  r.classList.add('draw-preview');
  r.setAttribute('x', Math.min(x0, x1));
  r.setAttribute('y', Math.min(y0, y1));
  r.setAttribute('width', Math.abs(x1 - x0));
  r.setAttribute('height', Math.abs(y1 - y0));
  OVERLAY_SVG.appendChild(r);
}

document.addEventListener('pointermove', e => {
  if (!DRAG_STATE) return;
  const [px, py] = eventToPdfPt(e);
  if (DRAG_STATE.mode === 'move') {
    const dx = px - DRAG_STATE.start[0], dy = py - DRAG_STATE.start[1];
    const o = DRAG_STATE.origRect;
    DRAG_STATE.ann.rect = [o[0]+dx, o[1]+dy, o[2]+dx, o[3]+dy];
    refreshOverlay();
  } else if (DRAG_STATE.mode === 'resize') {
    const dx = px - DRAG_STATE.start[0], dy = py - DRAG_STATE.start[1];
    const o = DRAG_STATE.origRect;
    let r = o.slice();
    const h = DRAG_STATE.handle;
    if (h.includes('n')) r[1] = o[1] + dy;
    if (h.includes('s')) r[3] = o[3] + dy;
    if (h.includes('w')) r[0] = o[0] + dx;
    if (h.includes('e')) r[2] = o[2] + dx;
    // ensure x0 < x1, y0 < y1
    if (r[0] > r[2]) [r[0], r[2]] = [r[2], r[0]];
    if (r[1] > r[3]) [r[1], r[3]] = [r[3], r[1]];
    DRAG_STATE.ann.rect = r;
    refreshOverlay();
  } else if (DRAG_STATE.mode === 'drawRect') {
    DRAG_STATE.cur = [px, py];
    drawPreviewRect();
  }
});
document.addEventListener('pointerup', e => {
  if (!DRAG_STATE) return;
  if (DRAG_STATE.mode === 'move' || DRAG_STATE.mode === 'resize') {
    const ann = DRAG_STATE.ann;
    const orig = DRAG_STATE.origRect;
    const moved = ann.rect.some((v,i) => Math.abs(v - orig[i]) > 0.1);
    if (moved) {
      setDirty(true);
      refreshUndoRedoButtons();
    } else {
      // No real change — discard the speculative undo we pushed at pointerdown
      UNDO_STACK.pop();
      refreshUndoRedoButtons();
    }
  } else if (DRAG_STATE.mode === 'drawRect') {
    const [x0, y0] = DRAG_STATE.start, [x1, y1] = DRAG_STATE.cur;
    const w = Math.abs(x1 - x0), h = Math.abs(y1 - y0);
    if (w >= 4 && h >= 4) {
      addRectAt(Math.min(x0,x1), Math.min(y0,y1), Math.max(x0,x1), Math.max(y0,y1));
    }
    const prev = OVERLAY_SVG.querySelector('rect.draw-preview');
    if (prev) prev.remove();
  }
  DRAG_STATE = null;
});

// Helper: push current state to undo stack (BEFORE applying a change).
function _commitBeforeChange() {
  UNDO_STACK.push(snapshotAnnots());
  if (UNDO_STACK.length > 200) UNDO_STACK.shift();
  REDO_STACK = [];
}

// ── Add rect / text / delete ───────────────────────────────────
function addRectAt(x0, y0, x1, y1) {
  _commitBeforeChange();
  const id = newClientId();
  const ann = {
    _id: id, _isNew: true,
    xref: null,
    page: PDF_PAGE,
    type: 'Square',
    rect: [x0, y0, x1, y1],
    contents: '',
  };
  EDIT_ANNOTS.push(ann);
  SELECTED_ANN_ID = id;
  setDirty(true);
  refreshOverlay();
  refreshUndoRedoButtons();
  setTool('select');
}
function addTextAt(x, y, defaultText) {
  _commitBeforeChange();
  const id = newClientId();
  const w = 130, h = 14;
  const ann = {
    _id: id, _isNew: true,
    xref: null,
    page: PDF_PAGE,
    type: 'FreeText',
    rect: [x, y, x + w, y + h],
    contents: defaultText || 'text',
  };
  EDIT_ANNOTS.push(ann);
  SELECTED_ANN_ID = id;
  setDirty(true);
  refreshOverlay();
  refreshUndoRedoButtons();
  // Open inline editor immediately
  setTool('select');
  setTimeout(() => openTextEditor(ann), 50);
}
function deleteSelected() {
  if (!SELECTED_ANN_ID) return;
  const idx = EDIT_ANNOTS.findIndex(a => a._id === SELECTED_ANN_ID);
  if (idx < 0) return;
  _commitBeforeChange();
  const a = EDIT_ANNOTS[idx];
  if (a._isNew) {
    EDIT_ANNOTS.splice(idx, 1);
  } else {
    a._deleted = true;
  }
  SELECTED_ANN_ID = null;
  setDirty(true);
  refreshOverlay();
  refreshUndoRedoButtons();
}

// ── Inline text editor ─────────────────────────────────────────
let TEXT_EDITOR = null;
function openTextEditor(a) {
  closeTextEditor();
  if (!OVERLAY_HOST || a.type !== 'FreeText') return;
  const img = OVERLAY_HOST.querySelector('img.pdf-page-img');
  if (!img) return;
  const pageSize = CURRENT_PDF.meta.page_sizes[PDF_PAGE - 1];
  const scaleX = img.clientWidth / pageSize[0];
  const scaleY = img.clientHeight / pageSize[1];
  const ed = document.createElement('textarea');
  ed.className = 'text-editor';
  ed.value = a.contents || '';
  ed.style.left = (a.rect[0] * scaleX) + 'px';
  ed.style.top = (a.rect[1] * scaleY) + 'px';
  ed.style.width = ((a.rect[2] - a.rect[0]) * scaleX) + 'px';
  ed.style.height = Math.max(20, (a.rect[3] - a.rect[1]) * scaleY) + 'px';
  ed.style.fontSize = Math.max(8, (a.rect[3] - a.rect[1]) * scaleY * 0.7) + 'px';
  OVERLAY_HOST.appendChild(ed);
  ed.focus();
  ed.select();
  TEXT_EDITOR = {ed, ann: a, originalContents: a.contents};
  ed.addEventListener('blur', commitTextEditor);
  ed.addEventListener('keydown', e => {
    if (e.key === 'Escape') { e.preventDefault(); cancelTextEditor(); }
    else if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) { e.preventDefault(); commitTextEditor(); }
  });
}
function commitTextEditor() {
  if (!TEXT_EDITOR) return;
  const newVal = TEXT_EDITOR.ed.value;
  const a = TEXT_EDITOR.ann;
  if (newVal !== TEXT_EDITOR.originalContents) {
    _commitBeforeChange();
    a.contents = newVal;
    setDirty(true);
    refreshUndoRedoButtons();
  }
  TEXT_EDITOR.ed.remove();
  TEXT_EDITOR = null;
  refreshOverlay();
}
function cancelTextEditor() {
  if (!TEXT_EDITOR) return;
  TEXT_EDITOR.ed.remove();
  TEXT_EDITOR = null;
}
function closeTextEditor() { commitTextEditor(); }

// Double-click annotation → open text editor (FreeText only)
document.addEventListener('dblclick', e => {
  if (!EDIT_MODE) return;
  const annNode = e.target.closest('g.annot');
  if (!annNode) return;
  const id = annNode.dataset.id;
  const a = EDIT_ANNOTS.find(x => x._id === id);
  if (a && a.type === 'FreeText') {
    SELECTED_ANN_ID = id;
    openTextEditor(a);
  }
});

// ── Save ───────────────────────────────────────────────────────
function computeEditDiff() {
  const edits = [];
  for (const a of EDIT_ANNOTS) {
    if (a._isNew && !a._deleted) {
      edits.push({action:'create', client_id: a._id, page: a.page,
                  type: a.type, rect: a.rect, contents: a.contents});
    } else if (!a._isNew && a._deleted) {
      edits.push({action:'delete', xref: a.xref});
    } else if (!a._isNew && !a._deleted) {
      const o = ORIGINAL_ANNOTS_BY_ID[a._id];
      if (!o) continue;
      const u = {action: 'update', xref: a.xref};
      let dirty = false;
      if (a.contents !== o.contents) { u.contents = a.contents; dirty = true; }
      const r1 = a.rect, r2 = o.rect;
      if (Math.abs(r1[0]-r2[0])>0.5 || Math.abs(r1[1]-r2[1])>0.5 ||
          Math.abs(r1[2]-r2[2])>0.5 || Math.abs(r1[3]-r2[3])>0.5) {
        u.rect = a.rect; dirty = true;
      }
      if (dirty) edits.push(u);
    }
  }
  return edits;
}
async function saveEdits() {
  if (!CURRENT_PDF || !DIRTY) return;
  closeTextEditor();
  const edits = computeEditDiff();
  if (!edits.length) { setDirty(false); return; }
  const r = await fetch('/api/pdf_save', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({rel: CURRENT_PDF.rel, edits}),
  });
  const j = await r.json();
  if (!j.ok) {
    alert('Save failed: ' + (j.error || JSON.stringify(j)));
    return;
  }
  // Reload PDF + meta to get fresh xrefs
  const meta = await fetch('/api/pdf_meta?rel=' + encodeURIComponent(CURRENT_PDF.rel)).then(r => r.json());
  CURRENT_PDF.meta = meta;
  initEditAnnots();
  setDirty(false);
  renderPdf();
  console.log(`✓ saved: ${j.applied} applied, ${j.errors} errors, snapshot: ${j.snapshot}`);
}

// ── History panel ──────────────────────────────────────────────
async function showHistory() {
  if (!CURRENT_PDF) return;
  const r = await fetch('/api/pdf_history?rel=' + encodeURIComponent(CURRENT_PDF.rel));
  const j = await r.json();
  const list = document.getElementById('history-list');
  list.innerHTML = '';
  document.getElementById('history-info').textContent =
    `${j.snapshots.length} snapshots — ${CURRENT_PDF.rel}`;
  for (const s of j.snapshots) {
    const ts = new Date(s.mtime * 1000).toLocaleString();
    const li = document.createElement('li');
    li.className = 'history-item';
    li.innerHTML = `
      <span class="ts">${escapeHtml(s.id)}</span>
      <span class="size">${(s.size/1024).toFixed(1)} KB · ${escapeHtml(ts)}</span>
      <button onclick="restoreSnapshot('${escapeHtml(s.id)}')">↺ Restore</button>`;
    list.appendChild(li);
  }
  if (!j.snapshots.length) {
    list.innerHTML = '<div style="padding:10px;color:#9ca3af;">ยังไม่มี snapshot</div>';
  }
  document.getElementById('history-modal').classList.add('show');
}
function closeHistory() {
  document.getElementById('history-modal').classList.remove('show');
}
async function restoreSnapshot(snapshotId) {
  if (!confirm(`Restore PDF จาก snapshot "${snapshotId}"?\nสถานะปัจจุบันจะถูก snapshot ก่อน restore`)) return;
  const r = await fetch('/api/pdf_restore', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({rel: CURRENT_PDF.rel, snapshot: snapshotId}),
  });
  const j = await r.json();
  if (!j.ok) { alert('Restore failed: ' + (j.error || '?')); return; }
  closeHistory();
  // reload
  const meta = await fetch('/api/pdf_meta?rel=' + encodeURIComponent(CURRENT_PDF.rel)).then(r => r.json());
  CURRENT_PDF.meta = meta;
  if (EDIT_MODE) initEditAnnots();
  renderPdf();
}
async function manualSnapshot() {
  if (!CURRENT_PDF) return;
  await fetch('/api/pdf_snapshot', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({rel: CURRENT_PDF.rel, tag: 'manual'}),
  });
  showHistory();  // refresh list
}

// ── Project versions modal (wraps scripts/version.py) ─────────
let DIFF_SELECTION = [];   // up to 2 ids selected for diff

function showVersions() {
  document.getElementById('versions-modal').classList.add('show');
  loadVersions();
}
function closeVersions() {
  document.getElementById('versions-modal').classList.remove('show');
  hideVersionsBusy();
  document.getElementById('versions-output').style.display = 'none';
  DIFF_SELECTION = [];
}
function showVersionsBusy(msg) {
  document.getElementById('versions-busy-msg').textContent = msg || 'working…';
  document.getElementById('versions-busy').style.display = '';
}
function hideVersionsBusy() {
  document.getElementById('versions-busy').style.display = 'none';
}
function showVersionsOutput(text) {
  const pre = document.getElementById('versions-output');
  pre.textContent = text || '(no output)';
  pre.style.display = '';
}

async function loadVersions() {
  const list = document.getElementById('versions-list');
  list.innerHTML = '<div style="color:#9ca3af;padding:20px;text-align:center;">Loading…</div>';
  try {
    const [rv, rs] = await Promise.all([
      fetch('/api/versions'),
      fetch('/api/versions/sync'),
    ]);
    const j = await rv.json();
    const sync = await rs.json();
    document.getElementById('versions-info').innerHTML =
      j.available
        ? `📁 ${escapeHtml(j.root)} · <strong>${j.snapshots.length}</strong> snapshot(s)`
        : '<span style="color:#b91c1c">scripts/version.py ไม่พบ — ระบบ versioning ไม่พร้อม</span>';
    renderSyncBadge(sync);
    refreshSyncIndicators(sync);
    if (!j.snapshots.length) {
      list.innerHTML = '<div style="color:#9ca3af;padding:24px;text-align:center;">ยังไม่มี snapshot — กด Quick snap เพื่อเริ่ม</div>';
      return;
    }
    // Mark the current latest visually so users see at a glance which is "live"
    const latestId = sync && sync.latest && sync.latest.id;
    list.innerHTML = j.snapshots.map(s => renderVersionItem(s, s.id === latestId)).join('');
    updateDiffHelper();
  } catch (e) {
    list.innerHTML = `<div style="color:#b91c1c;padding:20px;">โหลดไม่ได้: ${escapeHtml(e.message)}</div>`;
  }
}

const _SYNC_LABELS = {
  in_sync:           {label: '✓ ตรงกับ snapshot ล่าสุด',         cls: 'in-sync'},
  working_ahead:     {label: '⬆ working dir ใหม่กว่า snapshot',  cls: 'working-ahead'},
  working_behind:    {label: '⬇ working dir เก่ากว่า snapshot',  cls: 'working-behind'},
  divergent:         {label: '⚠ ต่างจาก snapshot ล่าสุด',         cls: 'divergent'},
  incomplete_local:  {label: '⚠ ไฟล์หายจาก working dir',          cls: 'incomplete-local'},
  no_snapshots:      {label: 'ยังไม่มี snapshot',                  cls: 'no-snapshots'},
};

function renderSyncBadge(sync) {
  const host = document.getElementById('versions-sync-badge');
  if (!sync || !host) { if (host) host.innerHTML = ''; return; }
  const meta = _SYNC_LABELS[sync.state] || _SYNC_LABELS.no_snapshots;
  const latest = sync.latest;
  let detail = '';
  if (latest) {
    detail = `<span class="latest-id">↳ ${escapeHtml(latest.id)}` +
             `${latest.tag ? ' · ' + escapeHtml(latest.tag) : ''}</span>`;
  }
  host.innerHTML = `<div class="sync-badge ${meta.cls}">${escapeHtml(meta.label)} ${detail}</div>`;
}

function refreshSyncIndicators(sync) {
  // 1. Top-level Versions button — change colour for warn / danger states
  const btn = document.querySelector('.versions-btn');
  if (btn) {
    btn.classList.remove('warn', 'danger');
    if (sync && (sync.state === 'working_behind' || sync.state === 'divergent' ||
                 sync.state === 'incomplete_local')) {
      btn.classList.add('danger');
    } else if (sync && sync.state === 'working_ahead') {
      btn.classList.add('warn');
    }
  }
  // 2. Top banner — only for behind / divergent / incomplete (these need user
  //    decision; "ahead" is auto-handled at boot so no banner needed)
  const banner = document.getElementById('sync-banner');
  const app = document.getElementById('app');
  if (!sync || !banner) return;
  if (window._SYNC_BANNER_DISMISSED) {
    banner.classList.remove('show');
    if (app) app.classList.remove('has-sync-banner');
    return;
  }
  let bannerCls = '', msg = '', actions = '';
  if (sync.state === 'working_behind') {
    bannerCls = 'danger';
    msg = `⚠ Working dir เก่ากว่า snapshot ล่าสุด (${escapeHtml(sync.latest.id)}). โหลดเวอร์ชันล่าสุด?`;
    actions = `<button onclick="quickRestoreLatest(false)">↺ Restore xlsx</button>`
            + `<button onclick="quickRestoreLatest(true)">↻ Restore full</button>`;
  } else if (sync.state === 'divergent') {
    bannerCls = 'danger';
    msg = `⚠ ต่างจาก snapshot ล่าสุด (${escapeHtml(sync.latest.id)})`;
    actions = `<button onclick="showVersions()">📚 ตรวจสอบ</button>`;
  } else if (sync.state === 'incomplete_local') {
    bannerCls = 'danger';
    msg = `⚠ ไฟล์บางส่วนหาย — เทียบกับ snapshot ${escapeHtml(sync.latest.id)}`;
    actions = `<button onclick="showVersions()">📚 ตรวจสอบ</button>`;
  }
  if (msg) {
    banner.className = 'sync-banner show ' + bannerCls;
    document.getElementById('sync-banner-msg').textContent = '';
    document.getElementById('sync-banner-msg').textContent = msg;
    document.getElementById('sync-banner-actions').innerHTML = actions;
    if (app) app.classList.add('has-sync-banner');
  } else {
    banner.classList.remove('show');
    if (app) app.classList.remove('has-sync-banner');
  }
}

function dismissSyncBanner() {
  window._SYNC_BANNER_DISMISSED = true;
  document.getElementById('sync-banner').classList.remove('show');
  const app = document.getElementById('app');
  if (app) app.classList.remove('has-sync-banner');
}

async function quickRestoreLatest(full) {
  const r = await fetch('/api/versions/sync');
  const sync = await r.json();
  if (!sync.latest) { alert('ไม่มี snapshot ให้ restore'); return; }
  await restoreVersion(sync.latest.id, full);
}

function renderVersionItem(s, isLatest) {
  const sizeMB = (s.size / (1024 * 1024)).toFixed(1);
  const ts = s.timestamp ? new Date(s.timestamp).toLocaleString() : '';
  const diffSel = DIFF_SELECTION.includes(s.id) ? ' diff-selected' : '';
  const latestBadge = isLatest ? ' <span style="background:#10b981;color:white;padding:1px 6px;border-radius:3px;font-size:9px;font-weight:600;">LATEST</span>' : '';
  return `
    <div class="version-item${diffSel}" style="${isLatest ? 'border-left: 3px solid #10b981;' : ''}">
      <input type="checkbox" class="v-check" ${DIFF_SELECTION.includes(s.id) ? 'checked' : ''}
             onchange="toggleDiffSelect('${escapeHtml(s.id)}', this.checked)"
             title="เลือกเพื่อ diff">
      <div>
        <div class="v-tag">
          <span class="v-kind ${s.kind}">${s.kind}</span>
          ${s.tag ? escapeHtml(s.tag) : '<span class="untagged">(no tag)</span>'}
          ${latestBadge}
        </div>
        <div class="v-meta">${escapeHtml(s.id)} · ${sizeMB} MB · ${s.n_output} files · ${escapeHtml(ts)}</div>
      </div>
      <div class="v-actions">
        <button onclick="diffOne('${escapeHtml(s.id)}')" title="Diff กับ snapshot ล่าสุด">📋 diff</button>
        <button class="restore" onclick="restoreVersion('${escapeHtml(s.id)}', false)">↺ restore</button>
        ${s.has_tarball ? `<button class="restore-full" onclick="restoreVersion('${escapeHtml(s.id)}', true)" title="restore output/ ทั้งหมด">↻ full</button>` : ''}
      </div>
    </div>`;
}

function toggleDiffSelect(id, checked) {
  if (checked) {
    if (!DIFF_SELECTION.includes(id)) DIFF_SELECTION.push(id);
    if (DIFF_SELECTION.length > 2) DIFF_SELECTION.shift();
  } else {
    DIFF_SELECTION = DIFF_SELECTION.filter(x => x !== id);
  }
  loadVersions();
  if (DIFF_SELECTION.length === 2) {
    runDiff(DIFF_SELECTION[0], DIFF_SELECTION[1]);
  }
}

function updateDiffHelper() {
  const helper = document.getElementById('diff-helper');
  if (DIFF_SELECTION.length === 0) helper.textContent = 'เลือก 2 snapshots เพื่อ diff';
  else if (DIFF_SELECTION.length === 1) helper.textContent = 'เลือกอีก 1 snapshot';
  else helper.textContent = `diff: ${DIFF_SELECTION[0]} ↔ ${DIFF_SELECTION[1]}`;
}

async function runDiff(id1, id2) {
  showVersionsBusy('comparing snapshots…');
  try {
    const r = await fetch(`/api/versions/diff?id1=${encodeURIComponent(id1)}&id2=${encodeURIComponent(id2)}`);
    const j = await r.json();
    showVersionsOutput(j.stdout || j.error || '(empty)');
  } finally {
    hideVersionsBusy();
  }
}
async function diffOne(id) {
  // diff between this snapshot and the latest one
  const r = await fetch('/api/versions');
  const j = await r.json();
  const latest = j.snapshots[0];
  if (!latest || latest.id === id) {
    showVersionsOutput('(this is already the latest snapshot)');
    return;
  }
  await runDiff(id, latest.id);
}

async function takeProjectSnap(full) {
  const tag = document.getElementById('versions-tag').value.trim();
  if (full) {
    if (!confirm('Full snapshot จะ tar.gz ทั้ง output/ (~200 MB) ใช้เวลาประมาณ 30-60 วินาที — ดำเนินการต่อ?')) return;
  }
  showVersionsBusy(full ? 'compressing output/ (อาจใช้เวลา 30-60s)…' : 'snapshotting…');
  try {
    const r = await fetch('/api/versions/snap', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({tag, full}),
    });
    const j = await r.json();
    showVersionsOutput(j.stdout || j.stderr || j.error || '(no output)');
    if (j.ok) document.getElementById('versions-tag').value = '';
    await loadVersions();
  } finally {
    hideVersionsBusy();
  }
}

async function autoSnapNow() {
  showVersionsBusy('checking xlsx for changes…');
  try {
    const r = await fetch('/api/versions/auto-snap', {method: 'POST'});
    const j = await r.json();
    showVersionsOutput(j.stdout || j.error || '(no output)');
    await loadVersions();
  } finally {
    hideVersionsBusy();
  }
}

async function pruneVersions() {
  const keep = parseInt(prompt('เก็บ N snapshots ล่าสุด (ลบที่เหลือ):', '10'));
  if (!keep || keep < 1) return;
  showVersionsBusy(`pruning to ${keep}…`);
  try {
    const r = await fetch('/api/versions/prune', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({keep}),
    });
    const j = await r.json();
    showVersionsOutput(j.stdout || '(no output)');
    await loadVersions();
  } finally {
    hideVersionsBusy();
  }
}

async function restoreVersion(id, full) {
  const msg = full
    ? `⚠️ FULL RESTORE จะแทนที่ทั้ง output/ ของคุณด้วย snapshot ${id}\n` +
      `(ระบบจะ snapshot สถานะปัจจุบันก่อนเสมอ)\n\nดำเนินการต่อ?`
    : `Restore xlsx + SKILL.md จาก ${id}?\n` +
      `(ระบบจะ snapshot สถานะปัจจุบันก่อน)`;
  if (!confirm(msg)) return;
  showVersionsBusy(full ? 'restoring full output/ (อาจใช้เวลา)…' : 'restoring…');
  try {
    const r = await fetch('/api/versions/restore', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({id, full}),
    });
    const j = await r.json();
    showVersionsOutput(j.stdout || j.stderr || j.error || '(no output)');
    if (j.ok) {
      await loadVersions();
      alert('✓ Restore เสร็จแล้ว — กำลัง reload หน้าเพื่อ reset state');
      setTimeout(() => location.reload(), 500);
    }
  } finally {
    hideVersionsBusy();
  }
}

// ── Auto-annotate modal ───────────────────────────────────────
let _AUTO_PLAN = null;
let _AUTO_BATCH_PLANS = [];

function showAutoAnnotate() {
  if (!SELECTED_ROW) { alert('เลือก row ก่อน'); return; }
  document.getElementById('auto-modal').classList.add('show');
  document.getElementById('auto-single').style.display = '';
  document.getElementById('auto-batch').style.display = 'none';
  document.getElementById('auto-output').style.display = 'none';
  loadAutoPreview(SELECTED_ROW);
}
function closeAuto() {
  document.getElementById('auto-modal').classList.remove('show');
}

async function loadAutoPreview(rowNum) {
  const r = await fetch(`/api/auto_annotate/preview?row=${rowNum}`);
  const plan = await r.json();
  _AUTO_PLAN = plan;
  const role = (plan.role && plan.role.role) || 'unknown';
  document.getElementById('auto-title').textContent =
    `R${rowNum} · section ${plan.section || '?'} · role: ${role}`;
  // Confidence badge — high/med/low based on threshold
  const cf = plan.confidence ?? 0;
  const cfCls = cf >= 0.85 ? 'high' : cf >= 0.65 ? 'med' : 'low';
  const genTag = plan.generator
    ? `<span class="src">${escapeHtml(plan.generator)}</span>` : '';
  const cfBadge = `<span class="conf-badge ${cfCls}">${(cf*100).toFixed(0)}% ${genTag}</span>`;
  let provLine = '';
  if (plan.provenance && Object.keys(plan.provenance).length) {
    provLine = `<br>🧠 <em>learned</em>: ` +
      Object.entries(plan.provenance).map(([k, v]) =>
        `${k} ← <code>${escapeHtml(v.trigger || '')}</code> (${((v.confidence||0)*100).toFixed(0)}%, ${v.samples} samples)`
      ).join(' · ');
  }
  document.getElementById('auto-meta').innerHTML =
    `<strong>Catalog:</strong> ${escapeHtml((plan.pdf_rel || '').split('/').pop() || '(none)')}` +
    (plan.match ? ` · matched <strong>${plan.match.score}</strong> tokens on page <strong>${plan.match.page}</strong>` : '') +
    ` ${cfBadge}${provLine}`;

  const warn = document.getElementById('auto-warn');
  if (plan.warnings && plan.warnings.length) {
    warn.style.display = '';
    warn.textContent = plan.warnings.map(w => '⚠ ' + w).join('\n');
  } else {
    warn.style.display = 'none';
  }

  const setPre = (id, text) => {
    const el = document.getElementById(id);
    if (text) { el.classList.remove('empty'); el.textContent = text; }
    else { el.classList.add('empty'); el.textContent = '(empty)'; }
  };
  setPre('auto-c', plan.proposed_c || '');
  setPre('auto-d', plan.proposed_d || '');
  const ann = plan.annotations || [];
  document.getElementById('auto-ann-count').textContent = `(${ann.length})`;
  setPre('auto-ann', ann.length
    ? ann.map(a => `[${a.type}] page ${a.page} rect=${JSON.stringify(a.rect)} contents=${JSON.stringify(a.contents)}`).join('\n')
    : '(no annotations)');
}

async function applyAutoSingle() {
  if (!_AUTO_PLAN) return;
  if (!_AUTO_PLAN.ok) {
    if (!confirm('Plan has warnings — apply anyway?')) return;
  }
  const writeXlsx = document.getElementById('auto-write-xlsx').checked;
  const writePdf = document.getElementById('auto-write-pdf').checked;
  const r = await fetch('/api/auto_annotate/apply', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({row: _AUTO_PLAN.row, write_xlsx: writeXlsx, write_pdf: writePdf}),
  });
  const j = await r.json();
  const out = document.getElementById('auto-output');
  out.style.display = '';
  out.textContent = JSON.stringify(j, null, 2);
  if (j.ok) {
    alert('✓ Applied');
    closeAuto();
    location.reload();
  }
}

async function showBatchAuto() {
  document.getElementById('auto-single').style.display = 'none';
  document.getElementById('auto-batch').style.display = '';
  const list = document.getElementById('auto-batch-list');
  list.innerHTML = '<div style="padding:20px;color:#9ca3af;text-align:center;">Computing plans for all rows…</div>';
  const r = await fetch('/api/auto_annotate/batch_preview');
  const j = await r.json();
  _AUTO_BATCH_PLANS = j.plans || [];
  document.getElementById('auto-batch-info').innerHTML =
    `<strong>${j.count}</strong> rows have plans · ` +
    Object.entries(j.by_role).filter(([k,v]) => v).map(([k,v]) => `${k}: ${v}`).join(' · ');
  if (!_AUTO_BATCH_PLANS.length) {
    list.innerHTML = '<div style="padding:20px;color:#9ca3af;text-align:center;">ไม่มี row ที่ต้อง auto-annotate</div>';
    return;
  }
  list.innerHTML = _AUTO_BATCH_PLANS.map(p => renderBatchRow(p)).join('');
}
function showSingleAuto() {
  document.getElementById('auto-single').style.display = '';
  document.getElementById('auto-batch').style.display = 'none';
}

function renderBatchRow(plan) {
  const role = (plan.role && plan.role.role) || 'unknown';
  const cls = (plan.warnings && plan.warnings.length) ? 'warn' : '';
  const okBox = plan.ok && !(plan.warnings && plan.warnings.length);
  return `
    <div class="auto-batch-row ${cls}">
      <input type="checkbox" class="b-check" data-row="${plan.row}" ${okBox ? 'checked' : ''}>
      <span class="b-row">R${plan.row}</span>
      <span class="b-role ${role}">${role}</span>
      <span class="b-d" title="${escapeHtml(plan.proposed_d || '')}">${escapeHtml(plan.proposed_d || '(empty)').slice(0, 80)}</span>
      <button onclick="loadAutoPreviewFromBatch(${plan.row})" style="font-size:10px;padding:2px 6px;">👁 view</button>
    </div>`;
}
function loadAutoPreviewFromBatch(rowNum) {
  showSingleAuto();
  loadAutoPreview(rowNum);
}

async function applyAutoBatch() {
  const checked = [...document.querySelectorAll('.b-check:checked')]
    .map(el => parseInt(el.dataset.row));
  if (!checked.length) { alert('เลือก row อย่างน้อย 1 อัน'); return; }
  if (!confirm(`Apply auto-annotate ${checked.length} rows?\n(ระบบจะ snap ก่อน)`)) return;
  const list = document.getElementById('auto-batch-list');
  list.innerHTML = `<div style="padding:20px;color:#4b5563;text-align:center;"><div class="spinner"></div>กำลัง apply ${checked.length} rows…</div>`;
  const r = await fetch('/api/auto_annotate/batch_apply', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({rows: checked}),
  });
  const j = await r.json();
  const out = document.getElementById('auto-output');
  out.style.display = '';
  out.textContent = `Applied ${j.applied_ok}/${j.count} rows\n\n` +
    JSON.stringify(j.results.slice(0, 20), null, 2) +
    (j.results.length > 20 ? `\n... and ${j.results.length - 20} more` : '');
  alert(`✓ ${j.applied_ok}/${j.count} applied`);
  closeAuto();
  location.reload();
}

// ── Manual annotate workflow (when AI fell back to "ยินดีปฏิบัติ") ──
// State machine:
//   1. user clicks 📍 Mark on a row that has a catalog
//   2. we load context (suggested label) + force edit mode + drawRect tool
//   3. user click-drags to create a Square; we auto-add a paired FreeText
//      label nearby with the SKILL.md content
//   4. user can drag/resize either rect
//   5. user clicks ✓ Save → POST /api/manual_annotate/save which writes
//      both annotations into the PDF AND auto-fills Col D in xlsx based on
//      the page where the rect sits

let MANUAL_MODE = false;
let MANUAL_TARGET_ROW = null;
let MANUAL_SUGGESTED_LABEL = '';
let MANUAL_SQUARE_ID = null;
let MANUAL_LABEL_ID = null;

async function startManualAnnotate() {
  if (!SELECTED_ROW) { alert('เลือก row ก่อน'); return; }

  const cr = await fetch(`/api/manual_annotate/context?row=${SELECTED_ROW}`);
  const cx = await cr.json();
  if (!cx.ok) { alert('โหลด context ไม่ได้: ' + cx.error); return; }

  // If row didn't have a pdf_rel but we got candidates, let user pick
  let chosen_pdf_rel = cx.pdf_rel;
  if (cx.candidates && cx.candidates.length > 1 && (!ROWS_BY_NUM[SELECTED_ROW] || !ROWS_BY_NUM[SELECTED_ROW].pdf_rel)) {
    const lines = cx.candidates.map((c, i) => `${i+1}. [${c.folder}] ${c.name}`).join('\n');
    const ans = prompt(
      `Row นี้เป็น "ยินดีปฏิบัติ" และไม่มี PDF อ้างอิง — เลือก catalog ที่จะ annotate:\n\n${lines}\n\nพิมพ์เลข 1-${cx.candidates.length}:`,
      '1');
    if (!ans) return;
    const idx = parseInt(ans) - 1;
    if (idx < 0 || idx >= cx.candidates.length) { alert('เลขไม่ถูกต้อง'); return; }
    chosen_pdf_rel = cx.candidates[idx].rel;
  }
  if (!chosen_pdf_rel) {
    alert('ไม่มี catalog PDF สำหรับ row นี้');
    return;
  }

  MANUAL_MODE = true;
  MANUAL_TARGET_ROW = SELECTED_ROW;
  MANUAL_SUGGESTED_LABEL = cx.suggested_label || '';
  MANUAL_SQUARE_ID = null;
  MANUAL_LABEL_ID = null;

  // If chosen PDF differs from current, load it
  if (!CURRENT_PDF || CURRENT_PDF.rel !== chosen_pdf_rel) {
    await loadPdf(chosen_pdf_rel, 1, null);
  }

  const banner = document.getElementById('manual-banner');
  document.getElementById('manual-target-info').innerHTML =
    `🎯 กำลัง mark <strong>R${cx.row}</strong> · section <code>${escapeHtml(cx.section || '')}</code>` +
    ` · label: <code>${escapeHtml(cx.suggested_label)}</code>` +
    `<br><span style="font-size:11px">👉 ลากกรอบรอบเนื้อหาใน catalog → ระบบจะแปะ label และอัพเดต Col D อัตโนมัติ ` +
    `(B: <em>${escapeHtml((cx.col_b||'').slice(0,80).trim())}</em>)</span>`;
  banner.classList.add('show');
  document.getElementById('manual-save-btn').disabled = true;

  if (!EDIT_MODE) toggleEditMode();
  setTool('drawRect');
}

function cancelManualAnnotate() {
  if (MANUAL_MODE && DIRTY) {
    if (!confirm('ยกเลิกการ annotate? การเปลี่ยนแปลงที่ยังไม่ save จะหาย')) return;
  }
  MANUAL_MODE = false;
  MANUAL_TARGET_ROW = null;
  MANUAL_SUGGESTED_LABEL = '';
  document.getElementById('manual-banner').classList.remove('show');
  // Drop unsaved changes — easiest way is to exit edit mode (which prompts)
  if (EDIT_MODE) {
    if (DIRTY) UNDO_STACK = [];  // clear dirty pending changes
    setDirty(false);
    EDIT_ANNOTS = [];
    SELECTED_ANN_ID = null;
    if (EDIT_MODE) toggleEditMode();
  }
}

// Hook into addRectAt: when in MANUAL_MODE and the user creates the FIRST
// Square, auto-add a paired FreeText label with the suggested content.
const _origAddRectAt = addRectAt;
addRectAt = function(x0, y0, x1, y1) {
  _origAddRectAt(x0, y0, x1, y1);
  if (!MANUAL_MODE) return;
  const sq = EDIT_ANNOTS[EDIT_ANNOTS.length - 1];
  if (!sq) return;
  MANUAL_SQUARE_ID = sq._id;
  // Auto-place a label to the right of the square (fallback below)
  const pageSize = CURRENT_PDF.meta.page_sizes[PDF_PAGE - 1] || [595, 842];
  const labelW = Math.min(180, Math.max(80, MANUAL_SUGGESTED_LABEL.length * 6));
  const labelH = 14;
  let lx0 = x1 + 5, ly0 = (y0 + y1) / 2 - labelH / 2;
  if (lx0 + labelW > pageSize[0] - 5) {
    // place below
    lx0 = x0; ly0 = y1 + 4;
  }
  const id = newClientId();
  const label = {
    _id: id, _isNew: true, xref: null,
    page: PDF_PAGE, type: 'FreeText',
    rect: [lx0, ly0, lx0 + labelW, ly0 + labelH],
    contents: MANUAL_SUGGESTED_LABEL,
  };
  EDIT_ANNOTS.push(label);
  MANUAL_LABEL_ID = id;
  SELECTED_ANN_ID = id;  // select the label so user can drag it
  setDirty(true);
  refreshOverlay();
  // Enable Save & Update Col D
  document.getElementById('manual-save-btn').disabled = false;
  // Switch to Select tool so user can drag/edit the label
  setTool('select');
};

async function saveManualAnnotate() {
  if (!MANUAL_MODE || !MANUAL_TARGET_ROW) return;
  const sq = EDIT_ANNOTS.find(a => a._id === MANUAL_SQUARE_ID);
  const lb = EDIT_ANNOTS.find(a => a._id === MANUAL_LABEL_ID);
  if (!sq || !lb) {
    alert('ต้องลาก rect ก่อน save');
    return;
  }
  const btn = document.getElementById('manual-save-btn');
  btn.disabled = true;
  btn.textContent = 'Saving…';
  try {
    const r = await fetch('/api/manual_annotate/save', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({
        row: MANUAL_TARGET_ROW,
        page: sq.page,
        content_rect: sq.rect,
        label_rect: lb.rect,
        label_text: lb.contents,
        pdf_rel: CURRENT_PDF.rel,
      }),
    });
    const j = await r.json();
    if (!j.ok) {
      alert('Save failed: ' + (j.error || JSON.stringify(j)));
      btn.disabled = false; btn.textContent = '✓ Save & update Col D';
      return;
    }
    alert(`✓ Saved\n\nOld Col D: ${j.old_d || '(empty)'}\n\nNew Col D: ${j.new_d}`);
    MANUAL_MODE = false;
    document.getElementById('manual-banner').classList.remove('show');
    location.reload();
  } catch (e) {
    alert('Save error: ' + e.message);
    btn.disabled = false; btn.textContent = '✓ Save & update Col D';
  }
}

// ── Toast notifications ───────────────────────────────────────
function toast(title, body, kind = 'info', timeout = 5000) {
  const wrap = document.getElementById('toasts');
  if (!wrap) return;
  const el = document.createElement('div');
  el.className = 'toast ' + kind;
  el.setAttribute('role', kind === 'error' ? 'alert' : 'status');
  el.setAttribute('aria-live', kind === 'error' ? 'assertive' : 'polite');
  el.innerHTML = `<span class="close" role="button" aria-label="ปิด" tabindex="0" onclick="this.parentElement.remove()">✕</span>` +
                 `<div class="title">${escapeHtml(title)}</div>` +
                 (body ? `<div class="body">${escapeHtml(body)}</div>` : '');
  // Keep stack reasonable
  while (wrap.children.length >= 6) wrap.firstChild.remove();
  wrap.appendChild(el);
  if (timeout) {
    setTimeout(() => {
      el.classList.add('fading');
      setTimeout(() => el.remove(), 400);
    }, timeout);
  }
}

// ── Theme toggle (light/dark) ─────────────────────────────────
function toggleTheme() {
  const root = document.documentElement;
  const cur = root.getAttribute('data-theme');
  let next;
  if (cur === 'dark') next = 'light';
  else if (cur === 'light') next = 'dark';
  else {
    // No explicit preference yet → flip from system
    const isDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
    next = isDark ? 'light' : 'dark';
  }
  root.setAttribute('data-theme', next);
  try { localStorage.setItem('comply-theme', next); } catch (e) {}
  const btn = document.getElementById('theme-toggle');
  if (btn) btn.textContent = next === 'dark' ? '☀' : '🌓';
}
(function initTheme() {
  try {
    const saved = localStorage.getItem('comply-theme');
    if (saved === 'dark' || saved === 'light') {
      document.documentElement.setAttribute('data-theme', saved);
      const btn = document.getElementById('theme-toggle');
      if (btn) btn.textContent = saved === 'dark' ? '☀' : '🌓';
    }
  } catch (e) {}
})();

// ── Modal infrastructure: ESC to close + focus trap ────────────
const MODAL_CLOSERS = {
  'history-modal':  () => closeHistory(),
  'versions-modal': () => closeVersions(),
  'learn-modal':    () => closeLearning(),
  'audit-modal':    () => closeAudit(),
  'auto-modal':     () => closeAuto(),
};
function _topMostOpenModal() {
  const ids = Object.keys(MODAL_CLOSERS);
  for (const id of ids.reverse()) {
    const el = document.getElementById(id);
    if (el && (el.classList.contains('show') || el.style.display === 'flex')) return el;
  }
  return null;
}
document.addEventListener('keydown', (e) => {
  if (e.key !== 'Escape') return;
  const m = _topMostOpenModal();
  if (!m) return;
  // Don't intercept if user is in the middle of editing a Col D cell or text editor
  if (typeof TEXT_EDITOR !== 'undefined' && TEXT_EDITOR) return;
  const closer = MODAL_CLOSERS[m.id];
  if (closer) { e.preventDefault(); closer(); }
}, true);
// Simple focus trap: when a modal opens, focus its first focusable element
const _modalObserver = new MutationObserver((muts) => {
  for (const m of muts) {
    if (m.type !== 'attributes') continue;
    const el = m.target;
    if (!el.classList || !el.classList.contains('modal-bg')) continue;
    const open = el.classList.contains('show') || el.style.display === 'flex';
    if (open) {
      const first = el.querySelector('input, select, textarea, button, [tabindex="0"]');
      if (first) setTimeout(() => first.focus({preventScroll: true}), 50);
    }
  }
});
document.querySelectorAll('.modal-bg').forEach(m => {
  _modalObserver.observe(m, {attributes: true, attributeFilter: ['class', 'style']});
});

// ── Auto-retrain (HITL "system gets smarter") ─────────────────
// Track feedback events client-side; whenever we cross a threshold,
// trigger a background retrain and show a toast with what was learned.
let _PENDING_RETRAIN_COUNT = 0;
const RETRAIN_THRESHOLD = 5;
async function tickRetrain(reason = 'feedback') {
  _PENDING_RETRAIN_COUNT++;
  if (_PENDING_RETRAIN_COUNT < RETRAIN_THRESHOLD) return;
  _PENDING_RETRAIN_COUNT = 0;
  try {
    const r = await fetch('/api/learn/retrain', {method: 'POST'});
    const j = await r.json();
    const total = (j.promoted || 0) + (j.updated || 0);
    if (total > 0) {
      toast(`🧠 Learned from your edits`,
            `${j.promoted} new pattern${j.promoted!==1?'s':''}, ${j.updated} updated. Future suggestions will use these.`,
            'learn', 7000);
    }
  } catch (e) {}
}

// ── Smart navigation: jump to next "uncertain" row ────────────
// "Uncertain" = unverified + has flags + has PDF.  Navigates within the
// currently-visible (filtered) tree only.
function nextUncertainRow() {
  const status = DATA.status || {};
  const visible = [...document.querySelectorAll('.tree-row[data-row]')]
    .map(el => parseInt(el.dataset.row));
  const start = visible.indexOf(SELECTED_ROW);
  // Score each candidate: lower = more interesting to inspect
  function score(rn) {
    const row = ROWS_BY_NUM[rn];
    if (!row) return 1e9;
    const st = (status[rn] && status[rn].status) || 'unverified';
    if (st !== 'unverified') return 1e9;        // skip already-verified
    if (!row.pdf_rel) return 1e6;                // no PDF → low priority
    let s = 0;
    if (row.auto_flags && row.auto_flags.length) s -= 100;  // flags = high priority
    if (row.parsed && row.parsed.type === 'commitment') s += 50;
    if (row.needs_col_d) s -= 50;
    return s;
  }
  // Look forward from the selected row first
  const ahead = visible.slice(start + 1).map(rn => [score(rn), rn]);
  const before = visible.slice(0, start + 1).map(rn => [score(rn), rn]);
  ahead.sort((a, b) => a[0] - b[0]);
  before.sort((a, b) => a[0] - b[0]);
  const target = (ahead[0] && ahead[0][0] < 1e6 ? ahead[0][1]
                 : before[0] && before[0][0] < 1e6 ? before[0][1] : null);
  if (target) {
    selectRow(target);
    toast('▸ Next uncertain', `R${target}`, 'info', 1500);
  } else {
    toast('No uncertain rows left', 'ตรวจครบทุก row ที่มี catalog แล้ว 🎉', 'info', 3000);
  }
}

// ── Mobile tab navigation ─────────────────────────────────────
function setMobileTab(tab) {
  for (const b of document.querySelectorAll('.mobile-tabs button')) {
    const active = b.dataset.tab === tab;
    b.classList.toggle('active', active);
    b.setAttribute('aria-selected', active ? 'true' : 'false');
  }
  const map = {tree: '.tree-pane', center: '.center-pane', pdf: '.pdf-pane'};
  for (const [k, sel] of Object.entries(map)) {
    const el = document.querySelector(sel);
    if (el) el.classList.toggle('mobile-active', k === tab);
  }
}

// Auto-show "center" tab on row select for mobile
const _origSelectRow = selectRow;
selectRow = async function(rowNum, scroll) {
  await _origSelectRow(rowNum, scroll);
  if (window.innerWidth <= 700) {
    // After picking a row on mobile, switch to the work area
    setMobileTab('center');
  }
};

// ── HITL Learning modal ───────────────────────────────────────
function showLearning() {
  document.getElementById('learn-modal').classList.add('show');
  loadLearnStats();
  loadLearnPatterns();
}
function closeLearning() {
  document.getElementById('learn-modal').classList.remove('show');
  document.getElementById('learn-output').style.display = 'none';
}

async function loadLearnStats() {
  const r = await fetch('/api/learn/stats');
  const j = await r.json();
  const acc = (j.accuracy * 100).toFixed(0);
  const cards = [
    ['total_feedbacks', `feedbacks (${j.window_days}d)`],
    ['accepted', 'accepted'],
    ['edited', 'edited'],
    ['rejected', 'rejected'],
    ['accuracy_pct', `accuracy ${acc}%`],
    ['patterns_total', 'patterns'],
    ['patterns_enabled', 'enabled'],
  ];
  const data = {...j, accuracy_pct: `${acc}%`};
  document.getElementById('learn-stats').innerHTML = cards.map(([k, l]) =>
    `<div class="stat-card"><div class="v">${data[k] ?? 0}</div><div class="l">${l}</div></div>`
  ).join('');
  document.getElementById('learn-llm-name').textContent =
    j.llm.available ? `${j.llm.name} ✓` : 'off (rule-only)';
}

async function loadLearnPatterns() {
  const t = document.getElementById('learn-pattern-filter').value;
  const url = '/api/learn/patterns' + (t ? `?type=${encodeURIComponent(t)}` : '');
  const r = await fetch(url);
  const j = await r.json();
  const list = document.getElementById('learn-pattern-list');
  if (!j.patterns.length) {
    list.innerHTML = '<div style="padding:20px;color:#9ca3af;text-align:center;">ยังไม่มี learned patterns — แก้ไข Col D หลายๆ row แล้วกด Retrain</div>';
    return;
  }
  list.innerHTML = j.patterns.map(p => {
    const cls = p.confidence >= 0.8 ? 'high' : p.confidence < 0.5 ? 'low' : '';
    const dis = p.enabled ? '' : ' disabled';
    return `<div class="learn-pattern-row${dis}" data-id="${p.id}">
      <span class="pt ${p.type}">${p.type}</span>
      <span class="trigger">${escapeHtml(p.trigger_key)}</span>
      <span><span class="arrow">→</span> <span class="output">${escapeHtml(p.output)}</span></span>
      <span class="conf"><span class="${cls}">${(p.confidence*100).toFixed(0)}%</span> · ${p.samples_total}</span>
      <label style="font-size:10px;display:flex;gap:3px;align-items:center;">
        <input type="checkbox" ${p.enabled?'checked':''} onchange="togglePattern(${p.id}, this.checked)"> on
      </label>
    </div>`;
  }).join('');
}

async function togglePattern(id, enabled) {
  await fetch(`/api/learn/pattern/${id}`, {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({enabled}),
  });
  loadLearnPatterns();
}

async function runRetrain() {
  const out = document.getElementById('learn-output');
  out.style.display = '';
  out.textContent = 'Retraining…';
  const r = await fetch('/api/learn/retrain', {method: 'POST'});
  const j = await r.json();
  out.textContent = `✓ Retrain complete\n  promoted: ${j.promoted}\n  updated:  ${j.updated}\n  patterns_examined: ${j.patterns_examined}`;
  loadLearnStats();
  loadLearnPatterns();
}

// ── Audit + DB modal ──────────────────────────────────────────
function showAudit() {
  document.getElementById('audit-modal').classList.add('show');
  loadAuditStats();
  loadAudit();
  document.getElementById('audit-search').value = '';
  document.getElementById('audit-search-results').style.display = 'none';
  document.getElementById('audit-list-wrap').style.display = '';
}
function closeAudit() {
  document.getElementById('audit-modal').classList.remove('show');
}

async function loadAuditStats() {
  const r = await fetch('/api/db/stats');
  const j = await r.json();
  const cards = [
    ['rows_total', 'rows'],
    ['rows_with_pdf', 'with PDF'],
    ['pdfs', 'catalog PDFs'],
    ['annotations', 'annotations'],
    ['snapshots', 'snapshots'],
    ['audit_entries', 'audit entries'],
    ['auto_annotates_applied', 'auto-annotates'],
    ['rows_needs_col_d', 'need Col D'],
  ];
  document.getElementById('audit-stats').innerHTML = cards.map(([k, l]) =>
    `<div class="stat-card"><div class="v">${j[k] ?? 0}</div><div class="l">${l}</div></div>`
  ).join('');
}

async function loadAudit() {
  const action = document.getElementById('audit-action-filter').value;
  let url = '/api/db/audit?limit=100';
  if (action) url += '&action=' + encodeURIComponent(action);
  const r = await fetch(url);
  const j = await r.json();
  const list = document.getElementById('audit-list');
  if (!j.entries.length) {
    list.innerHTML = '<div style="padding:20px;color:#9ca3af;text-align:center;">ยังไม่มี audit log</div>';
    return;
  }
  list.innerHTML = j.entries.map(e => {
    const tsHuman = e.ts ? e.ts.replace('T', ' ').slice(5, 19) : '';
    let body = '';
    if (e.before != null || e.after != null) {
      body = `<span class="ba"><span class="before">${escapeHtml(e.before || '∅')}</span><span class="arrow">→</span><span class="after">${escapeHtml(e.after || '∅')}</span></span>`;
    } else if (e.details) {
      body = `<span class="ba">${escapeHtml(JSON.stringify(e.details).slice(0, 80))}</span>`;
    }
    const targetClick = (e.target_type === 'row')
      ? `onclick="closeAudit();selectRow(${parseInt(e.target_id)||0})"` : '';
    return `<div class="audit-row">
      <span class="ts">${escapeHtml(tsHuman)}</span>
      <span class="ac ${e.action}">${escapeHtml(e.action)}</span>
      <span class="tg" ${targetClick} title="${escapeHtml(e.target_type||'')}">${escapeHtml((e.target_type||'')+(e.target_id ? ' '+e.target_id : ''))}</span>
      ${body}
    </div>`;
  }).join('');
}

let _AUDIT_SEARCH_TIMER = null;
function onAuditSearch() {
  clearTimeout(_AUDIT_SEARCH_TIMER);
  _AUDIT_SEARCH_TIMER = setTimeout(runAuditSearch, 250);
}
async function runAuditSearch() {
  const q = document.getElementById('audit-search').value.trim();
  const wrap = document.getElementById('audit-search-results');
  if (!q) {
    wrap.style.display = 'none';
    document.getElementById('audit-list-wrap').style.display = '';
    return;
  }
  document.getElementById('audit-list-wrap').style.display = 'none';
  wrap.style.display = '';
  wrap.innerHTML = '<div style="padding:14px;color:#9ca3af;">กำลังค้น…</div>';
  try {
    const r = await fetch('/api/db/search?q=' + encodeURIComponent(q));
    const j = await r.json();
    if (!j.results.length) {
      wrap.innerHTML = '<div style="padding:14px;color:#9ca3af;">ไม่พบ</div>';
      return;
    }
    wrap.innerHTML = `<h4 style="margin:8px 0 4px;font-size:12px;color:#4b5563;">FTS Results (${j.results.length})</h4>` +
      j.results.map(h =>
        `<div class="ar-row" onclick="closeAudit();selectRow(${h.row})">
          <strong>R${h.row}</strong> <span style="color:#6b7280">${escapeHtml(h.section || '')}</span>
          <div style="margin-top:2px">${h.snippet_b || ''}</div>
          ${h.snippet_d ? `<div style="margin-top:2px;color:#6b7280">D: ${h.snippet_d}</div>` : ''}
        </div>`
      ).join('');
  } catch (e) {
    wrap.innerHTML = `<div style="padding:14px;color:#b91c1c;">error: ${escapeHtml(e.message)}</div>`;
  }
}

// ── Stats ──────────────────────────────────────────────────────
function renderStats() {
  const st = DATA.stats, status = DATA.status || {};
  const c = {pass:0, fail:0, need_fix:0, skip:0};
  for (const k in status) if (c[status[k].status] !== undefined) c[status[k].status]++;
  const done = c.pass + c.fail + c.need_fix + c.skip;
  const remain = st.total - done;
  const pct = st.total ? Math.round((done / st.total) * 100) : 0;
  document.getElementById('stats').innerHTML = `
    <span><strong>${st.total}</strong> rows</span>
    <span><strong style="color:var(--c-success-text)">${c.pass}</strong> ✓</span>
    <span><strong style="color:var(--c-danger-text)">${c.fail}</strong> ✗</span>
    <span><strong style="color:var(--c-warn-text)">${c.need_fix}</strong> ⚠</span>
    <span><strong style="color:var(--c-text-muted)">${c.skip}</strong> ⏭</span>
    <span><strong>${remain}</strong> เหลือ</span>`;
  // Topbar stats-pill — compact summary
  const pill = document.getElementById('stats-pill-progress');
  if (pill) {
    pill.innerHTML =
      `<strong>${pct}%</strong> <span class="sep">·</span> ` +
      `<span style="color:var(--c-success-text)">${c.pass} ✓</span> <span class="sep">·</span> ` +
      `<span style="color:var(--c-danger-text)">${c.fail} ✗</span> <span class="sep">·</span> ` +
      `<span>${remain} เหลือ</span>`;
  }
}

// ── helpers ────────────────────────────────────────────────────
function escapeHtml(s) {
  return (s == null ? '' : String(s)).replace(/[&<>"']/g, c => ({
    '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'
  }[c]));
}

/** Confidence-dot for a row: green = looks OK, amber = commitment but
 *  has a catalog (might need manual mark), red = no PDF / parse error,
 *  gray = empty. */
function confDotHtml(r) {
  if (!r) return '';
  let cls = 'none', tip = 'no data';
  if (r.is_commitment === undefined) {
    r.is_commitment = (r.D || '').trim().startsWith('ยินดีปฏิบัติ');
  }
  if (r.parsed && r.parsed.type === 'commitment' && r.pdf_rel) {
    cls = 'med'; tip = 'commitment but catalog exists — consider Mark';
  } else if (r.pdf_rel && r.D) {
    cls = 'high'; tip = 'has catalog + Col D';
  } else if (!r.pdf_rel && (r.D || '').trim()) {
    cls = 'low'; tip = 'Col D set but no PDF resolved';
  } else if (r.D) {
    cls = 'med'; tip = 'has Col D, no catalog';
  }
  return `<span class="conf-dot ${cls}" title="${escapeHtml(tip)}"></span>`;
}

// ── filter wiring ──────────────────────────────────────────────
['tree-search','filter-status','filter-vendor','filter-flags','filter-haspdf'].forEach(id => {
  document.getElementById(id).addEventListener('input', applyFilters);
  document.getElementById(id).addEventListener('change', applyFilters);
});

// ── Keyboard ───────────────────────────────────────────────────
document.addEventListener('keydown', e => {
  const tag = e.target.tagName;
  const inText = (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT');
  // Cmd/Ctrl-S — save (works even from input)
  if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 's') {
    e.preventDefault();
    if (EDIT_MODE) saveEdits();
    return;
  }
  // Cmd/Ctrl-Z / Shift+Cmd+Z — undo/redo (only when in edit mode and not editing text)
  if (EDIT_MODE && !inText && (e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'z') {
    e.preventDefault();
    if (e.shiftKey) redo(); else undo();
    return;
  }
  if (EDIT_MODE && !inText && (e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'y') {
    e.preventDefault(); redo(); return;
  }
  if (inText) return;
  // Edit-mode tool shortcuts
  if (EDIT_MODE) {
    if (e.key === 'v' || e.key === 'V') { setTool('select'); e.preventDefault(); return; }
    if (e.key === 'r' || e.key === 'R') { setTool('drawRect'); e.preventDefault(); return; }
    if (e.key === 't' || e.key === 'T') { setTool('addText'); e.preventDefault(); return; }
    if ((e.key === 'Delete' || e.key === 'Backspace') && SELECTED_ANN_ID) {
      e.preventDefault(); deleteSelected(); return;
    }
    if (e.key === 'Escape') {
      if (TEXT_EDITOR) cancelTextEditor();
      else if (DIRTY) toggleEditMode();
      else { EDIT_MODE = false; toggleEditMode(); toggleEditMode(); /* clean exit */ }
      e.preventDefault(); return;
    }
    // Arrow keys nudge selected annotation (when not editing text)
    if (SELECTED_ANN_ID && (e.key === 'ArrowUp' || e.key === 'ArrowDown' ||
                            e.key === 'ArrowLeft' || e.key === 'ArrowRight')) {
      const sel = EDIT_ANNOTS.find(a => a._id === SELECTED_ANN_ID);
      if (sel) {
        const step = e.shiftKey ? 5 : 1;
        const dx = e.key === 'ArrowLeft' ? -step : e.key === 'ArrowRight' ? step : 0;
        const dy = e.key === 'ArrowUp' ? -step : e.key === 'ArrowDown' ? step : 0;
        _commitBeforeChange();
        sel.rect = [sel.rect[0]+dx, sel.rect[1]+dy, sel.rect[2]+dx, sel.rect[3]+dy];
        setDirty(true);
        refreshOverlay();
        refreshUndoRedoButtons();
        e.preventDefault();
        return;
      }
    }
  }
  // Global navigation
  if (e.key === 'j' || e.key === 'ArrowDown')  { moveRow(1);  e.preventDefault(); }
  else if (e.key === 'k' || e.key === 'ArrowUp') { moveRow(-1); e.preventDefault(); }
  else if (e.key === 'n' || e.key === 'N') { nextUncertainRow(); e.preventDefault(); }
  else if (e.key === '1') setStatus('pass');
  else if (e.key === '2') setStatus('fail');
  else if (e.key === '3') setStatus('need_fix');
  else if (e.key === '4') setStatus('skip');
  else if (e.key === '[') pdfPrev();
  else if (e.key === ']') pdfNext();
  else if (e.key === ',') torPrev();
  else if (e.key === '.') torNext();
  else if (e.key === '+' || e.key === '=') pdfZoom(1);
  else if (e.key === '-') pdfZoom(-1);
});
function moveRow(delta) {
  // walk through visible rows in tree order
  const visible = [...document.querySelectorAll('.tree-row[data-row]')]
    .map(el => parseInt(el.dataset.row));
  if (!visible.length) return;
  let idx = visible.indexOf(SELECTED_ROW);
  if (idx < 0) idx = 0;
  else idx = Math.max(0, Math.min(visible.length - 1, idx + delta));
  selectRow(visible[idx]);
}

// Auto-advance after verdict (toggleable)
let AUTO_ADVANCE = (localStorage.getItem('autoAdvance') !== '0');
function toggleAutoAdvance() {
  AUTO_ADVANCE = !AUTO_ADVANCE;
  localStorage.setItem('autoAdvance', AUTO_ADVANCE ? '1' : '0');
  const cb = document.getElementById('auto-advance');
  if (cb) cb.checked = AUTO_ADVANCE;
}
// patch setStatus to optionally advance
const _origSetStatus = setStatus;
setStatus = async function(status) {
  await _origSetStatus(status);
  if (AUTO_ADVANCE && status !== 'unverified') moveRow(1);
};

init();
</script>
</body>
</html>
"""



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
    print(f"[boot] indexing TOR sections + page text…")
    build_tor_section_index()
    index_tor_text()
    ch5 = sum(1 for s in TOR_SECTION_INDEX if s.startswith("5"))
    print(f"[boot] {len(TOR_SECTION_INDEX)} TOR sections ({ch5} in chapter 5), "
          f"{len(TOR_PAGE_TEXTS)} pages indexed for text-search")

    # Mirror everything into the DB so queries can hit it instead of the
    # in-memory caches.
    sync_db_from_memory()

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
            print(f"[boot] ⚠ working dir differs from latest snapshot — review in 📚 Versions")


def main() -> None:
    boot()
    host = "127.0.0.1"
    port = 5173
    url = f"http://{host}:{port}"
    print(f"\n  ✦ Comply Verify GUI ✦")
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
