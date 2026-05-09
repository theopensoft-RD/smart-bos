"""
comply_core.py — Object-oriented domain model for the Comply Verify tool.

The codebase grew organically as a collection of module-level dicts and
functions (ROWS, PDF_INDEX, etc.).  This module wraps those concepts in
classes so:

  • State has a single owner (Project) instead of leaking through globals.
  • Each entity (Row, CatalogPDF, TORDocument) carries its own behaviour.
  • The Flask layer in ``comply_verify_gui.py`` becomes thin glue that
    delegates to these classes.

Migration is intentionally gradual — the existing module-level functions
in ``comply_verify_gui.py`` still work and, behind the scenes, can use
these classes.  New code should prefer the classes.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterator


# ---------------------------------------------------------------------------
# Row — one line in the Comply spec spreadsheet
# ---------------------------------------------------------------------------

@dataclass
class Row:
    """One row of comply.xlsx, plus everything we infer about it.

    This is a lightweight dataclass — it holds state but doesn't own
    persistence.  The Project orchestrator hydrates Row instances from
    openpyxl and maintains the canonical list.
    """
    row: int
    A: str | None = None
    B: str | None = None
    C: str | None = None
    D: str | None = None
    E: str | None = None
    F: str | None = None

    # Derived (filled by Project on load)
    section: str | None = None
    section_inherited: bool = False
    pdf_rel: str | None = None
    pdf_inherited: bool = False
    needs_col_d: bool = False
    parsed: dict = field(default_factory=dict)

    # ------------------------------------------------------------------
    @property
    def col_d_kind(self) -> str:
        """Coarse classification of Col D ('brand_model', 'equivalent',
        'commitment', 'empty', etc.)."""
        return self.parsed.get("type") or "empty"

    @property
    def is_commitment(self) -> bool:
        return (self.D or "").strip().startswith("ยินดีปฏิบัติ")

    @property
    def is_section_header(self) -> bool:
        b = (self.B or "").strip()
        return bool(re.match(r"^\d+(?:\.\d+){2,3}\.\s+", b))

    @property
    def is_item(self) -> bool:
        b = (self.B or "").strip()
        return bool(re.match(r"^\d+\s*\)", b)) and not self.is_section_header

    @property
    def is_sub_item(self) -> bool:
        b = (self.B or "").strip()
        return bool(re.match(r"^\d+\s*\.", b)) and not self.is_section_header

    @property
    def role(self) -> str:
        if self.is_section_header: return "section_header"
        if self.is_item: return "item"
        if self.is_sub_item: return "sub_item"
        return "unknown"

    def b_preview(self, n: int = 80) -> str:
        return (self.B or "").strip()[:n]

    def to_dict(self) -> dict:
        """JSON-serialisable view (matches the legacy rows_payload shape)."""
        return {
            "row": self.row,
            "A": self.A, "B": self.B, "C": self.C, "D": self.D,
            "E": self.E, "F": self.F,
            "section": self.section,
            "section_inherited": self.section_inherited,
            "pdf_rel": self.pdf_rel,
            "pdf_inherited": self.pdf_inherited,
            "needs_col_d": self.needs_col_d,
            "parsed": self.parsed,
            "col_d_kind": self.col_d_kind,
            "role": self.role,
            "is_commitment": self.is_commitment,
        }


# ---------------------------------------------------------------------------
# CatalogPDF — one catalog file with its annotations
# ---------------------------------------------------------------------------

@dataclass
class CatalogPDF:
    rel_path: str                 # relative to OUTPUT
    folder_key: str | None = None       # "5.1.2-2" if folder follows that pattern
    section_prefix: str | None = None    # set when folder doesn't have -N (sensors / เสา)
    full_path: Path | None = None        # absolute path (set by Project)
    size: int = 0
    mtime: float = 0
    num_pages: int | None = None
    detected_brand: str | None = None
    detected_model: str | None = None

    @property
    def stem(self) -> str:
        return self.rel_path.rsplit("/", 1)[-1].rsplit(".", 1)[0]

    @property
    def folder_name(self) -> str:
        return self.rel_path.rsplit("/", 2)[-2] if "/" in self.rel_path else ""

    def to_dict(self) -> dict:
        return {
            "rel_path": self.rel_path,
            "folder_key": self.folder_key,
            "section_prefix": self.section_prefix,
            "size": self.size,
            "mtime": self.mtime,
            "num_pages": self.num_pages,
            "brand": self.detected_brand,
            "model": self.detected_model,
        }


# ---------------------------------------------------------------------------
# Project — orchestrator that wires Rows, CatalogPDFs, TOR and persistence
# ---------------------------------------------------------------------------

class Project:
    """Top-level container for one Comply project.

    Responsibilities:
      • own the canonical list of Rows
      • own the catalog PDF index
      • coordinate boot sync (xlsx + filesystem → memory)
      • expose query helpers (find rows by section, find PDF for ref, etc.)

    This class is intentionally thin — it delegates actual file I/O to the
    existing module-level functions for now.  Over time, those should move
    in here and the global mutable state should disappear.
    """

    def __init__(self, root: Path):
        self.root = Path(root)
        self.output = self.root / "output"
        self.tor_dir = self.root / "TOR"
        self.boq_dir = self.root / "BOQ"
        self.scripts_dir = self.root / "scripts"
        self.versions_dir = self.root / "_versions"
        self.snaps_dir = self.versions_dir / "snapshots"
        self.db_path = self.root / "_db" / "comply.db"
        self.xlsx_path = self.output / "Comply spec Smart Plant 1.xlsx"
        self.status_path = self.output / "verification_status.json"

        # In-memory caches (kept in sync with filesystem on boot/refresh)
        self._rows: list[Row] = []
        self._catalogs: list[CatalogPDF] = []

    # ------- collections ----------------------------------------------
    @property
    def rows(self) -> list[Row]:
        return self._rows

    @property
    def catalogs(self) -> list[CatalogPDF]:
        return self._catalogs

    def set_rows(self, rows: list[Row]) -> None:
        self._rows = rows

    def set_catalogs(self, pdfs: list[CatalogPDF]) -> None:
        self._catalogs = pdfs

    def row(self, row_num: int) -> Row | None:
        return next((r for r in self._rows if r.row == row_num), None)

    def rows_in_section(self, section: str, include_descendants: bool = True
                        ) -> Iterator[Row]:
        prefix = section + "."
        for r in self._rows:
            if not r.section: continue
            if r.section == section: yield r
            elif include_descendants and r.section.startswith(prefix): yield r

    def stats(self) -> dict:
        n = len(self._rows)
        with_pdf = sum(1 for r in self._rows if r.pdf_rel)
        commitment = sum(1 for r in self._rows if r.is_commitment)
        return {
            "total": n,
            "with_pdf": with_pdf,
            "commitment": commitment,
            "by_role": {
                "section_header": sum(1 for r in self._rows if r.is_section_header),
                "item":           sum(1 for r in self._rows if r.is_item),
                "sub_item":       sum(1 for r in self._rows if r.is_sub_item),
            },
        }


# ---------------------------------------------------------------------------
# Convenience: convert legacy row dicts ↔ Row objects
# ---------------------------------------------------------------------------

def from_legacy_dict(d: dict) -> Row:
    """Build a Row from the dict shape used by ROWS in comply_verify_gui."""
    return Row(
        row=d.get("row", 0),
        A=d.get("A"), B=d.get("B"), C=d.get("C"), D=d.get("D"),
        E=d.get("E"), F=d.get("F"),
        section=d.get("section_inferred"),
        section_inherited=bool(d.get("section_inherited")),
        pdf_rel=d.get("pdf_rel"),
        pdf_inherited=bool(d.get("pdf_inherited")),
        needs_col_d=bool(d.get("needs_col_d")),
        parsed=d.get("parsed") or {},
    )
