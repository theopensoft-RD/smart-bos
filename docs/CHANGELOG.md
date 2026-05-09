# Changelog & Lessons Learned

A chronological-ish log of what was built, what broke, and why we decided
the way we did. Not exhaustive — just the things future-me (or another
agent) needs to avoid re-discovering.

---

## Phase 1 — Foundation (initial GUI)

**Built**:
- 3-column Flask app (Tree / Center / Catalog) with PDF rendering via PyMuPDF
- Row-list with section/vendor/status filters
- TOR PDF preview with text-search highlight
- xlsx context preview (±N rows around selected)
- Catalog PDF viewer with annotation list + zoom

**Lessons**:
- PyMuPDF crashes on PDFs with broken appearance streams (`m_internal NoneType`).
  → Always wrap iteration in `safe_iter_annots()` with try/except per annot.
- Thai SARA AM (ำ) vs decomposed NIKHAHIT+SARA AA (ํา) breaks `search_for`.
  → Need NFKD + manual replace, not just NFC/NFD.

---

## Phase 2 — TOR section-aware search

**Problem**: User opened a row in chapter 5 of comply, but TOR jumped to
chapter 1-3 because the same generic phrases appeared earlier.

**Built**:
- TOR section index (`5.1.2 → P12`, etc.) via line-start regex on each page
- `find_in_tor(text, section)` searches within the section's page range first,
  expands forward only (content overflows down, not up), falls back to
  chapter-bounded global search — but **never** to full-document fallback.
- "Last resort = section start with no highlight" beats "wrong-chapter match".

**Lessons**:
- Section inheritance: must propagate from preceding section header to all
  sub-rows. R15 had D ref `5.1.1-2` but its section is `5.1.1.2` — depth-3
  ref vs depth-4 actual section. Always pick the **most-specific** of
  `(last_section, d_ref_derived)`.
- Forward widening only: TOR section content spills onto next pages, never
  backwards. Symmetric widening pulled R159 onto P13 (wrong section).

---

## Phase 3 — R71 highlight bug + structured matcher

**Problem**: R71 (ข้อย่อย 1.) highlighted both rows 1 and 10 in the catalog
because "ข้อย่อย 1" is a substring of "ข้อย่อย 10".

**Fix**: Replaced substring matcher with `_match_annot_label()` that parses
both query and content into `{item, subitem}` and matches digit-for-digit.

**Now handles**:
- ✓ `ข้อย่อย 1` ≠ `ข้อย่อย 10`
- ✓ Sub-item query lights up its parent rect (context cue)
- ✓ Parent-only query doesn't light up sub-items
- ✓ `ยี่ห้อ` / `รุ่น` literals match exactly
- ✓ 14/14 unit tests in the smoke suite

---

## Phase 4 — เสา catalog inline annotations

**Problem**: 5.1.3.2 / 5.1.4.3 / 5.1.6.3 (เสา = pole) PDFs had annotations
that didn't show up. PyMuPDF returned 0 annots for them.

**Root cause**: เสา PDFs embed annotations as **literal dicts** in the page's
`/Annots` array (not indirect references). PyMuPDF skips them.

**Fix**:
1. Added `parse_inline_annots()` that reads the raw `/Annots` value via
   `doc.xref_get_key(page_xref, "Annots")` and parses each `<<...>>` block.
2. Apply `page.transformation_matrix` to the rect (PDF user-space is
   bottom-up Y; PyMuPDF's `Annot.rect` is top-down — inline parsing must
   transform manually).
3. Replaced y-only label↔square pairing with `_rect_edge_distance()` so
   labels-below-square (เสา convention) pair correctly.

---

## Phase 5 — R21 / R24 wrong highlight

**R21**: "Power Supply Redundant Hot Swap" appears twice on TOR P10 (Server
item 7 + NGFW item 11). Original matcher highlighted both.

**Fix**: `_section_y_bounds_on_page()` finds the next-sibling section
header on the page; rects beyond that y are filtered out.

**R24**: Token-fallback (when full phrase didn't match in PyMuPDF) gathered
"Firewall" rects from 3 different lines — scattered highlights.

**Fix**: `_anchor_cluster_by_rarest()` picks the rarest-frequency token as
anchor (e.g. "Throughput" appears once), keeps only rects within ±14pt of
the anchor's y. One-line cluster every time.

---

## Phase 6 — comply-module migration + version system

**What changed**: User created `comply-module/` as the project root and
moved everything in. Added `scripts/version.py` (snapshot tool).

**Built**:
- Path auto-detection (`PROJECT = ROOT if (ROOT / "TOR").exists() else ROOT.parent`)
  works in both legacy and new layout.
- `/api/versions/*` endpoints wrap version.py via subprocess.
- "Always-load-latest" invariant: on boot, if working files differ from
  the latest snapshot, banner + auto-snap (working_ahead) or alert
  (working_behind/divergent).
- Sync status badge in 📚 Versions modal.

**Lessons**:
- Don't use `detect_types=PARSE_DECLTYPES` with sqlite3 — its bundled
  timestamp converter chokes on ISO strings. Store timestamps as plain
  TEXT and parse manually.
- snap/restore is the user's safety net. Every destructive operation
  (xlsx write, PDF edit, restore) must pre-snap.

---

## Phase 7 — Catalog resolution rabbit hole

The biggest source of confusion. Multiple bugs over multiple iterations.

**Bug 1**: Col D was empty for 50+ rows after user added new catalogs
(Lenovo Server, FortiGate, Ruijie, NAS, iPad, UPS, etc.) but didn't yet
update Col D.
- **Fix**: Pass 4 in `load_rows()` — for empty-Col-D rows with a section,
  resolve a candidate PDF via folder convention.

**Bug 2**: `model_only` resolver picked WRONG PDF when MODEL_INDEX had
multiple matches (e.g. R349 matched UFC9312A instead of UF-2010A because
"UF" appears in both).
- **Fix**: `_pick_best_in_folder()` scores each candidate by token
  overlap with Col D + bonuses for section-prefix match.

**Bug 3**: User's xlsx mixed dot-form (`5.1.1.2`) and dash-form (`5.1.1-2`)
Col D refs. Resolver only knew dash-form.
- **Fix**: `_ref_to_folder_keys()` translates between both forms
  bidirectionally. Try all candidate keys.

**Bug 4 (R19)**: Commitment rows (`ยินดีปฏิบัติ`) had no pdf_rel — user
couldn't use 📍 Mark to fix them.
- **Fix**: Pass 4 now also runs for commitment rows. Col D itself is
  unchanged (still says "ยินดีปฏิบัติ"); only `pdf_rel` is surfaced.

**Bug 5 (R8)**: Section header for 5.1.1 (depth-3, `len(parts)=3`) wildcard-
matched all `5.1.1-N` keys, but Python dict iteration order made `-4`
(L2 Switch) come first.
- **Fix**: Sort wildcard matches by trailing N ascending → rack parent
  (`-1`) wins.

**The big lesson**: Catalog resolution must be **deterministic and
sorted**. Never trust dict iteration order. Always tie-break by
section number ascending so the parent (-1) wins.

---

## Phase 8 — DB layer (SQLite)

**Built `app/database.py`** with 12 tables:
- Mirrors of xlsx state (rows + FTS5)
- Catalog PDFs + annotations
- TOR sections + per-page text
- Verification status (replaces `verification_status.json`)
- Snapshots + per-PDF history (mirrors of filesystem)
- Audit log (append-only)
- Auto-annotate plan history
- Learning feedback + learned patterns

**Endpoints**: `/api/db/{stats,audit,search,section_progress}`

**UI**: 📊 Audit modal with 8-card stats, audit timeline, FTS search box

**Lessons**:
- DB is **derived state**. Always rebuildable from xlsx + filesystem.
- Write-through: writes go to xlsx → then mirror to DB. Never DB-only.
- FTS5 + `unicode61` tokenizer handles Thai well enough for Col B/D search.

---

## Phase 9 — HITL Learning loop

**Built `app/learning.py`**:
- `record_feedback()` — logs every (suggestion → user action) triple
- `retrain_patterns()` — distils repeating corrections into rules
- `apply_learned_brand()` / `apply_learned_vendor()` — used by
  auto_annotate_plan to override rules with user-validated patterns
- `set_llm_provider()` — pluggable hook for Anthropic/OpenAI/Ollama

**Pattern types mined**:
- `filename_brand`: filename token → brand string
- `section_vendor`: section root → Col E vendor
- `row_format_d`: (role, section_root) → Col D shape

**UI**: 🧠 Learn modal with stats (accuracy %, feedback counts, pattern
list with on/off toggle, retrain button).

**Frontend integration**:
- `tickRetrain()` — counts feedback events, auto-retrains every 5
- Toast notifications when patterns get promoted
- Confidence badge in auto-annotate preview shows generator + provenance

---

## Phase 10 — Manual-annotate workflow

For commitment rows where AI couldn't find Col B in catalog, user can
manually mark.

**Flow**:
1. 📍 Mark button on action bar (pulses for commitment rows)
2. `/api/manual_annotate/context` returns suggested label + PDF candidates
3. Dialog if multiple candidate PDFs (commitment rows often have many)
4. Force edit mode + drawRect tool + green banner
5. User draws rect → auto-paired FreeText label with SKILL.md format
6. User can drag/resize either annotation
7. ✓ Save → `/api/manual_annotate/save`:
   - pre-snap
   - apply Square + FreeText to PDF
   - compute Col D via `make_col_d_for_row()`
   - write xlsx
   - audit + feedback record

**Lessons**:
- The candidate picker is critical for commitment rows: section
  `5.1.6.1` has 3 catalogs (`-1` rack, `-2` RCBO, `-3` Controller).
  User picks visually, doesn't have to remember which is which.

---

## Phase 11 — UX polish + responsive

**Built**:
- Responsive CSS: 4 breakpoints (3-col / 2-col / stacked / mobile-tabs)
- Toast notification system (top-right, 4 types: info/warn/error/learn)
- Confidence dots in tree (green/amber/red/gray)
- `N` key for "next uncertain row" (smart queue)
- Mobile tab navigation
- Inline Col D editing (double-click → contenteditable → save as feedback)
- Col D dropdown menu (single-click): Mark / Auto / Edit (commit) or
  Auto / Mark / Edit / Revert (has-ref)

**Auto-retrain trigger** moved to client-side `tickRetrain(reason)` —
fires on status verdict, inline edit, manual annotate, revert.

---

## Phase 12 — OOP refactor + file org

**Created `app/` package**:
- `app/core.py` — Row, CatalogPDF, Project (was `comply_core.py`)
- `app/database.py` — SQLite layer (was `comply_db.py`)
- `app/learning.py` — HITL pipeline (was `comply_learning.py`)

**Updated**:
- `comply_verify_gui.py` imports `from app import core, database, learning`
- Cross-import in `learning.py`: `from . import database as db`
- Mirrored `ROWS` into `PROJECT_OBJ.rows` after every `load_rows()` for
  gradual migration to OOP.

**Cleanup**:
- Removed 24 `.DS_Store` files (cosmetic — Finder will regenerate them)

---

## Things we considered and rejected

- **Splitting `comply_verify_gui.py`** (282 KB) into multiple files.
  Rejected for now — it works, splitting risks breakage. Do it when adding
  a feature naturally lives elsewhere.
- **Replacing flat-file xlsx with a real database** (Postgres). Rejected —
  user works on Google Drive, single-file is the constraint.
- **Vector-store / embeddings for FTS**. Rejected — SQLite FTS5 is
  fast enough and has no GPU dependency.
- **Real LLM fine-tuning**. Out of scope. Replaced with pattern-mining
  + LLM provider hook (off by default).
- **PDF.js in the browser** (instead of server-side render to PNG).
  Rejected — PyMuPDF's annotation handling is more reliable than PDF.js
  for these specific catalogs (and Thai font support is consistent
  server-side).

---

## Open questions / future work

1. **`annot_position` learning** is a stub. Could mine: "for catalog
   folder X, label Y position relative to Square is …". Useful for
   the auto-annotate label placement.
2. **Active learning prioritisation**: surface low-confidence rows
   first (already partly via `N` key). Could go further: a queue
   ordered by predicted accuracy gain from user feedback.
3. **Multi-user merge**: if 2 users edit different rows simultaneously,
   the second xlsx write wins. SQLite WAL handles concurrent reads
   but not writes. Either lock at app-level or use a CRDT for Col C/D.
4. **PDF migration tool**: convert inline annotations (เสา catalogs) to
   indirect refs so they become editable.
5. **Section-bounded TOR search edge cases**: ~2% miss rate on commitment
   rows (genuine no-match). Could fall back to a closest-cluster heuristic
   but the false-positive risk is high.
