"""
export.py — print-ready PDF package builder.

Assembles a single PDF deliverable that includes:

    [Cover page]                  project / company / date / version
    [Table of Contents]           section tree with page numbers
    [Comply Spec Sheet]           verbatim from output/Comply spec*.pdf
    [Catalog section]             one divider page per section, then the
                                   catalog PDFs in order (annotations
                                   baked-in, same as on disk)
    [Audit appendix (optional)]   recent audit_log entries

Each page gets a footer with page number + project name. PDF outline
(bookmarks) is set so PDF readers show a nav tree.

Public API
----------
build_package(out_path, *, include_catalogs=True,
              section_filter=None, include_audit=False,
              max_audit_entries=200,
              progress_cb=None) -> dict

The result dict carries page_count, byte_size, generated_at, and a
list of section bookmarks (for the UI to show before download).

Notes
-----
* xlsx → PDF conversion is NOT done here. We use the existing
  ``output/Comply spec *.pdf`` which the user maintains via Excel/
  LibreOffice. If it's missing we still produce a package but warn.
* Catalog annotations are picked up *because* PyMuPDF's ``insert_pdf``
  copies the source PDF's appearance streams. Phase 17's WYSIWYG bake
  means what's on disk is what the user sees — and what the export
  produces.
"""

from __future__ import annotations

import io
from collections.abc import Callable
from datetime import datetime
from pathlib import Path
from typing import Any

import fitz  # PyMuPDF

# ---------------------------------------------------------------------------
# Layout constants
# ---------------------------------------------------------------------------

PAGE_W = 595.0   # A4 portrait, points (1pt = 1/72 inch)
PAGE_H = 842.0
MARGIN = 50.0
FOOTER_H = 30.0
FONT_REG = "helv"          # PyMuPDF built-in (Latin); fine for headings
FONT_BOLD = "hebo"         # Helvetica Bold


# ---------------------------------------------------------------------------
# Page primitives
# ---------------------------------------------------------------------------

def _new_page(doc: fitz.Document) -> fitz.Page:
    return doc.new_page(width=PAGE_W, height=PAGE_H)


def _draw_text(page: fitz.Page, x: float, y: float, text: str, *,
               size: float = 11, bold: bool = False,
               color: tuple[float, float, float] = (0, 0, 0),
               max_width: float | None = None) -> float:
    """Draw text at (x, y) and return the y of the next line."""
    font = FONT_BOLD if bold else FONT_REG
    if max_width is None:
        page.insert_text((x, y), text, fontname=font, fontsize=size, color=color)
        return y + size * 1.4
    # Wrap at max_width using textbox
    rect = fitz.Rect(x, y - size, x + max_width, y + size * 30)
    rc = page.insert_textbox(rect, text, fontname=font, fontsize=size,
                              color=color, align=0)
    # rc returns y where text overflowed (positive = remaining), or
    # the height of inserted text (negative). Just compute from line count.
    lines = max(1, len(text) // max(1, int(max_width / (size * 0.55))) + 1)
    return y + size * 1.4 * lines


def _draw_hr(page: fitz.Page, y: float, *, color=(0.6, 0.6, 0.6),
              width: float = 0.5) -> None:
    page.draw_line((MARGIN, y), (PAGE_W - MARGIN, y),
                   color=color, width=width)


def _add_footer(page: fitz.Page, page_num: int, total: int,
                  project_name: str) -> None:
    """Bottom strip with project name (left) and page number (right)."""
    y = PAGE_H - FOOTER_H + 10
    _draw_hr(page, y - 8, color=(0.85, 0.85, 0.85), width=0.3)
    page.insert_text((MARGIN, y), project_name,
                      fontname=FONT_REG, fontsize=8, color=(0.4, 0.4, 0.4))
    txt = f"Page {page_num} of {total}"
    # Right-align
    text_w = fitz.get_text_length(txt, fontname=FONT_REG, fontsize=8)
    page.insert_text((PAGE_W - MARGIN - text_w, y), txt,
                      fontname=FONT_REG, fontsize=8, color=(0.4, 0.4, 0.4))


# ---------------------------------------------------------------------------
# Cover page
# ---------------------------------------------------------------------------

def _add_cover_page(doc: fitz.Document, *, company_name: str,
                     project_name: str, project_code: str | None,
                     version: str, generated_at: str) -> None:
    page = _new_page(doc)

    # Decorative top band
    page.draw_rect(fitz.Rect(0, 0, PAGE_W, 80),
                   color=None, fill=(0.31, 0.27, 0.90))  # primary indigo

    # Company name (white on band)
    page.insert_text((MARGIN, 50), company_name,
                      fontname=FONT_BOLD, fontsize=20, color=(1, 1, 1))

    # Big title
    y = 200
    page.insert_text((MARGIN, y), "COMPLIANCE PACKAGE",
                      fontname=FONT_REG, fontsize=14,
                      color=(0.45, 0.45, 0.45))
    y += 50
    title_max_w = PAGE_W - 2 * MARGIN
    # Wrap project name at the page width if it's long
    rect = fitz.Rect(MARGIN, y - 30, MARGIN + title_max_w, y + 60)
    page.insert_textbox(rect, project_name, fontname=FONT_BOLD,
                        fontsize=28, color=(0.13, 0.13, 0.13), align=0)
    y += 80

    if project_code:
        page.insert_text((MARGIN, y), f"Code: {project_code}",
                          fontname=FONT_REG, fontsize=12,
                          color=(0.45, 0.45, 0.45))
        y += 30

    # Metadata block at bottom
    y_meta = PAGE_H - 200
    _draw_hr(page, y_meta - 16)
    rows = [
        ("Generated",    generated_at),
        ("Version",      version or "—"),
        ("Format",       "Print-ready PDF (A4 portrait)"),
    ]
    for label, value in rows:
        page.insert_text((MARGIN, y_meta), label,
                          fontname=FONT_REG, fontsize=10,
                          color=(0.45, 0.45, 0.45))
        page.insert_text((MARGIN + 120, y_meta), value,
                          fontname=FONT_BOLD, fontsize=10,
                          color=(0.13, 0.13, 0.13))
        y_meta += 22


# ---------------------------------------------------------------------------
# Table of Contents
# ---------------------------------------------------------------------------

def _add_toc(doc: fitz.Document, entries: list[tuple[int, str, int]]) -> int:
    """Render TOC over as many pages as needed.

    ``entries`` is a list of ``(level, title, page1)`` where level is
    1 (top) / 2 (sub). page1 is 1-based.

    Returns the number of TOC pages added.
    """
    page = _new_page(doc)
    page.insert_text((MARGIN, MARGIN + 20), "Table of Contents",
                      fontname=FONT_BOLD, fontsize=20,
                      color=(0.13, 0.13, 0.13))
    _draw_hr(page, MARGIN + 36, color=(0.31, 0.27, 0.90), width=1.5)

    y = MARGIN + 60
    pages_used = 1
    line_h = 18

    for level, title, p1 in entries:
        if y > PAGE_H - FOOTER_H - 40:
            page = _new_page(doc)
            pages_used += 1
            y = MARGIN + 20
        indent = MARGIN + (level - 1) * 18
        size = 12 if level == 1 else 10
        bold = (level == 1)

        # Title (truncate if would collide with right-side page #)
        text_max_x = PAGE_W - MARGIN - 50  # reserve space for page #
        page.insert_text((indent, y), _truncate(title, 70 if level == 1 else 80),
                          fontname=FONT_BOLD if bold else FONT_REG,
                          fontsize=size,
                          color=(0.13, 0.13, 0.13) if bold else (0.3, 0.3, 0.3))

        # Dotted leader (cheap: just draw a thin line)
        leader_y = y + 1
        # Calc title width to know where leader starts
        title_w = fitz.get_text_length(_truncate(title, 70), fontname=FONT_BOLD if bold else FONT_REG,
                                         fontsize=size)
        leader_x0 = indent + title_w + 4
        if leader_x0 < text_max_x - 4:
            page.draw_line((leader_x0, leader_y),
                           (text_max_x - 4, leader_y),
                           color=(0.7, 0.7, 0.7), width=0.3,
                           dashes="[1 2] 0")

        # Page number (right-aligned)
        page_str = str(p1)
        page_w = fitz.get_text_length(page_str, fontname=FONT_REG, fontsize=size)
        page.insert_text((PAGE_W - MARGIN - page_w, y), page_str,
                          fontname=FONT_REG, fontsize=size,
                          color=(0.3, 0.3, 0.3))
        y += line_h

    return pages_used


def _truncate(s: str, maxlen: int) -> str:
    return s if len(s) <= maxlen else s[:maxlen - 1] + "…"


# ---------------------------------------------------------------------------
# Section divider
# ---------------------------------------------------------------------------

def _add_section_divider(doc: fitz.Document, *, section: str,
                          subtitle: str = "") -> None:
    page = _new_page(doc)
    # Big colored band
    page.draw_rect(fitz.Rect(0, PAGE_H * 0.32, PAGE_W, PAGE_H * 0.5),
                   color=None, fill=(0.95, 0.96, 1.0))
    page.insert_text((MARGIN, PAGE_H * 0.4),
                      f"Section {section}",
                      fontname=FONT_BOLD, fontsize=36,
                      color=(0.31, 0.27, 0.90))
    if subtitle:
        page.insert_text((MARGIN, PAGE_H * 0.45),
                          subtitle,
                          fontname=FONT_REG, fontsize=14,
                          color=(0.4, 0.4, 0.4))


# ---------------------------------------------------------------------------
# Audit appendix
# ---------------------------------------------------------------------------

def _add_audit_pages(doc: fitz.Document, entries: list[dict]) -> None:
    page = _new_page(doc)
    page.insert_text((MARGIN, MARGIN + 20), "Audit Log",
                      fontname=FONT_BOLD, fontsize=20,
                      color=(0.13, 0.13, 0.13))
    _draw_hr(page, MARGIN + 36, color=(0.31, 0.27, 0.90), width=1.5)
    y = MARGIN + 60

    for e in entries:
        if y > PAGE_H - FOOTER_H - 40:
            page = _new_page(doc)
            y = MARGIN + 20
        ts = (e.get("ts") or "")[:19]
        action = (e.get("action") or "")[:30]
        target = f"{e.get('target_type','?')} {e.get('target_id','')}".strip()
        line = f"{ts}  {action:<30}  {target}"
        page.insert_text((MARGIN, y), line,
                          fontname=FONT_REG, fontsize=8,
                          color=(0.2, 0.2, 0.2))
        y += 12


# ---------------------------------------------------------------------------
# Main builder
# ---------------------------------------------------------------------------

def build_package(
    *,
    out_path: Path | str,
    company_name: str,
    project_name: str,
    project_code: str | None = None,
    version: str = "",
    comply_pdf_path: Path | None = None,
    catalogs: list[dict] | None = None,
    output_root: Path | None = None,
    include_audit: bool = False,
    audit_entries: list[dict] | None = None,
    progress_cb: Callable[[str, int, int], None] | None = None,
) -> dict:
    """Build the print-ready package.

    Args:
        out_path: where to write the final PDF.
        company_name / project_name / project_code / version: cover info.
        comply_pdf_path: existing rendered comply spec PDF (xlsx export).
            Pass None to skip that section.
        catalogs: list of dicts ``{section_hint, brand, model, pdf_rel}``,
            *already filtered + sorted* by caller.
        output_root: base path where catalog ``pdf_rel`` resolves.
        include_audit: if True, append audit_entries as a final section.
        audit_entries: list of audit_log dicts.
        progress_cb: callable(stage, done, total) for UI progress.

    Returns dict with: page_count, byte_size, sections (TOC), out_path.
    """
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    catalogs = catalogs or []
    audit_entries = audit_entries or []

    def _tick(stage: str, done: int, total: int) -> None:
        if progress_cb:
            try:
                progress_cb(stage, done, total)
            except Exception:
                pass

    doc = fitz.open()  # blank
    bookmark_entries: list[tuple[int, str, int]] = []  # (level, title, page1)

    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M")

    # === 1. Cover (page 1) ============================================
    _tick("cover", 0, 1)
    _add_cover_page(doc, company_name=company_name, project_name=project_name,
                     project_code=project_code, version=version,
                     generated_at=generated_at)
    cover_page1 = 1
    bookmark_entries.append((1, "Cover", cover_page1))

    # === 2. TOC placeholder — render at the end with real page numbers
    # We can't know the real TOC entries until catalogs are inserted, so
    # we'll insert TOC pages AFTER assembling everything else, then move
    # them to position 2 with PDF page reordering.
    # For simplicity: pre-allocate a generous TOC page count, fill after.
    # Estimate: 1 page per ~30 entries.
    est_toc_entries = 2 + len(catalogs) + (1 if include_audit else 0)
    est_toc_pages = max(1, (est_toc_entries + 29) // 30)
    toc_start_page1 = cover_page1 + 1
    for _ in range(est_toc_pages):
        _new_page(doc)  # placeholder
    bookmark_entries.append((1, "Table of Contents", toc_start_page1))

    # === 3. Comply spec sheet =========================================
    if comply_pdf_path and Path(comply_pdf_path).exists():
        _tick("comply_sheet", 0, 1)
        comply_start = doc.page_count + 1  # 1-based
        try:
            with fitz.open(str(comply_pdf_path)) as src:
                doc.insert_pdf(src)
            bookmark_entries.append((1, "Comply Spec Sheet", comply_start))
        except Exception as e:
            # Insert a fallback page noting the failure
            page = _new_page(doc)
            page.insert_text((MARGIN, MARGIN + 40),
                              f"⚠ Could not embed Comply Spec PDF:\n{e}",
                              fontname=FONT_REG, fontsize=10,
                              color=(0.7, 0, 0))
            bookmark_entries.append(
                (1, "Comply Spec Sheet (failed)", doc.page_count))

    # === 4. Catalogs grouped by section ===============================
    if catalogs:
        catalogs_section_top_page = doc.page_count + 1
        bookmark_entries.append((1, "Catalogs", catalogs_section_top_page))

        # Group by section_hint
        from collections import OrderedDict
        groups: OrderedDict[str, list[dict]] = OrderedDict()
        for c in catalogs:
            key = c.get("section_hint") or "Other"
            groups.setdefault(key, []).append(c)

        total_cats = len(catalogs)
        done_cats = 0
        for section, items in groups.items():
            sec_start = doc.page_count + 1
            _add_section_divider(doc, section=str(section),
                                  subtitle=f"{len(items)} catalog"
                                            + ("s" if len(items) != 1 else ""))
            bookmark_entries.append((2, f"§ {section}", sec_start))

            for cat in items:
                done_cats += 1
                _tick("catalogs", done_cats, total_cats)
                pdf_rel = cat.get("pdf_rel")
                if not pdf_rel or not output_root:
                    continue
                src_path = (Path(output_root) / pdf_rel).resolve()
                if not src_path.exists():
                    continue
                cat_start = doc.page_count + 1
                title = " · ".join(filter(None, [
                    cat.get("brand"), cat.get("model")
                ])) or Path(pdf_rel).stem
                try:
                    with fitz.open(str(src_path)) as src:
                        doc.insert_pdf(src)
                    bookmark_entries.append(
                        (3, _truncate(title, 60), cat_start))
                except Exception:
                    # Skip unreadable catalogs but keep building
                    page = _new_page(doc)
                    page.insert_text((MARGIN, MARGIN + 40),
                                      f"⚠ Could not embed: {pdf_rel}",
                                      fontname=FONT_REG, fontsize=9,
                                      color=(0.7, 0, 0))

    # === 5. Audit appendix (optional) =================================
    if include_audit and audit_entries:
        _tick("audit", 0, 1)
        audit_start = doc.page_count + 1
        _add_audit_pages(doc, audit_entries)
        bookmark_entries.append((1, "Audit Log", audit_start))

    # === 6. Now fill the TOC placeholder pages ========================
    _tick("toc", 0, 1)
    # Take pages [toc_start_page1-1 .. toc_start_page1-1+est_toc_pages) and
    # rewrite them. We can't easily replace pages in PyMuPDF without
    # losing other state, so we write into them fresh (the placeholders
    # are blank).
    # First, render TOC into a temporary doc, then use insert_pdf to
    # substitute.
    tmp = fitz.open()
    actual_toc_pages = _add_toc(tmp, [
        (lvl, title, p1) for (lvl, title, p1) in bookmark_entries
        # Skip Cover / Table of Contents from listing themselves
        if title not in ("Cover", "Table of Contents")
    ])
    # Delete the est_toc_pages placeholders, replace with actual TOC.
    # PyMuPDF: delete_page uses 0-based.
    placeholder_start = toc_start_page1 - 1   # 0-based
    placeholder_end = placeholder_start + est_toc_pages - 1
    doc.delete_pages(from_page=placeholder_start, to_page=placeholder_end)
    # Now insert the TOC pages at that position
    doc.insert_pdf(tmp, from_page=0, to_page=actual_toc_pages - 1,
                    start_at=placeholder_start)
    tmp.close()

    # If actual_toc_pages != est_toc_pages, all bookmark entries with
    # page1 > toc_start_page1 + est_toc_pages need to shift. We compute
    # the delta and re-emit. We KEPT the bookmark entries in original
    # form; now adjust them.
    delta = actual_toc_pages - est_toc_pages
    if delta != 0:
        adjusted = []
        for lvl, title, p1 in bookmark_entries:
            if title == "Cover" or title == "Table of Contents":
                adjusted.append((lvl, title, p1))
            elif p1 > toc_start_page1:
                adjusted.append((lvl, title, p1 + delta))
            else:
                adjusted.append((lvl, title, p1))
        bookmark_entries = adjusted

    # === 7. Footers (page numbers) on every page =====================
    _tick("footers", 0, doc.page_count)
    total = doc.page_count
    for i, page in enumerate(doc):
        _add_footer(page, page_num=i + 1, total=total,
                     project_name=project_name)

    # === 8. PDF outline (bookmarks tree) =============================
    doc.set_toc([[lvl, title, p1] for (lvl, title, p1) in bookmark_entries])

    # === 9. Save ======================================================
    _tick("save", 0, 1)
    doc.save(str(out_path), garbage=4, deflate=True, clean=True)
    page_count = doc.page_count
    doc.close()

    byte_size = out_path.stat().st_size
    return {
        "ok": True,
        "out_path": str(out_path),
        "page_count": page_count,
        "byte_size": byte_size,
        "generated_at": generated_at,
        "sections": [{"level": lvl, "title": title, "page": p1}
                     for (lvl, title, p1) in bookmark_entries],
    }
