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

---

## Phase 13 — Production-grade UI polish (2026-05-09 / 2026-05-10)

**Built**:
- **Design system v2**: tokens (color/space/radius/shadow/type), light + dark
  + high-contrast modes via `[data-theme]`; `prefers-reduced-motion` honored;
  custom scrollbars; `:focus-visible` rings.
- **35-icon Lucide-style SVG sprite** + `ico()` JS helper. All chrome
  emoji-free — emojis remain only in toast titles for emotional accent.
- **Verdict segmented control** (Linear-style pill of 4 slices) replaces
  the 4 separate buttons that used to crowd the action bar.
- **Tree status bar**: 3px vertical colored strip (left edge) via
  `[data-status="pass|fail|need_fix|skip"]` replaces emoji-icon column.
  Cleaner visual scan of completion state.
- **Skeleton loaders**: shimmer animation for TOR/PDF/xlsx panes during
  fetch (previously showed plain "loading…" text).
- **iOS-style toggle switch** for "auto-next" replaces tiny checkbox.
- **Modal slot system** (`.modal__header / __body / __footer`) — for
  future modal standardisation; old modals still work.
- **Empty/Error states** with icon-circle + title + body + CTA.
- **Counter pop animation** when stats% changes.

**Lessons**:
- Stacking contexts bite: `.topbar { z-index: 10 }` + dropdown
  `z-index: 250` was clipped behind `.action-bar` (z=60 in document
  context). Solution: use `position: fixed` for menus that need to
  escape parent stacking contexts.
- `vector-effect: non-scaling-stroke` makes SVG strokes inconsistent
  with PyMuPDF baked output (1 screen-px vs 1 viewBox-pt = 1.8px @ 130dpi).
  Removed it for WYSIWYG match.

---

## Phase 14 — Re-annotate Wizard

**Built**:
- Multi-step rect+label drawing wizard for re-annotating brand_model
  rows (2 steps: ยี่ห้อ + รุ่น) and item/sub_item rows (1 step).
- **PDF picker** mid-flow: dropdown lets user switch the catalog while
  in the wizard (e.g. wrong PDF auto-resolved → manually pick correct one).
- Backend `_delete_inline_annot()` rewrites page `/Annots` array via
  `xref_set_key` to support deleting xref=0 inline annots — paired with
  `_split_annots_array_items()` that preserves indirect references in
  their original positions.
- New action `delete_inline` in `apply_pdf_edits` for the wizard's
  delete-existing-then-write flow.
- **Apply-to-siblings** prompt after brand_model save: "Apply same
  pattern to N sibling rows?" → opens bulk modal.
- `/api/reannotate/context` and `/api/reannotate/save` endpoints.

**Lessons**:
- PyMuPDF rejects `border_color` on `add_freetext_annot` unless
  `rich_text=True`. Solution: skip border on FreeText (the saved PDF
  has no border anyway), match by drawing zero-stroke `.ann-rect.freetext`
  in SVG. The paired Square provides the visible frame.
- Inline annotations (xref=0) collide on `_id='x0'` in EDIT_ANNOTS.
  Fix: assign unique IDs `inl-pN-K` from page + inline_index, so
  click+select+delete operate per-annot.
- `find_text_match_in_pdf` failed silently on rack catalog P1 — PyMuPDF's
  `page.annots()` generator throws when it hits a broken Image/Form annot,
  killing the whole iteration. Solution: bypass the generator entirely;
  use `annot_xrefs()` + per-xref `load_annot()` with try/except.

---

## Phase 15 — Anthropic Claude integration (2026-05-10)

**Built**:
- `app/anthropic_provider.py` (~370 lines): `AnthropicProvider` class
  wrapping the SDK with:
  - 3 tool definitions: `propose_col_d`, `propose_brand_model`, `escalate_to_user`
  - 4 cached system blocks (ephemeral, 5-min TTL): role + SKILL.md +
    KB.md + pitfalls.md + top-30 learned_patterns
  - Pricing table for Sonnet 4.5 / Opus 4.5 / Haiku 4.5; per-day USD
    budget enforcement raises `BudgetExceededError`
  - Records every call to `llm_calls` table (tokens, cost, elapsed,
    tool_calls)
  - `get_provider()` singleton + `install_into_learning()` bridge to
    the legacy `app.learning.set_llm_provider` hook
- New SQLite table `llm_calls` with indexes on (ts/row/model)
- `_maybe_refine_with_claude()` in `auto_annotate_plan`: skips when
  rules confidence ≥ 0.85 (no spend); otherwise asks Claude with
  `rule_proposal` + few-shot context. If Claude's confidence beats
  rules → override col_d/col_c/page. The rule output is the floor.
- `POST /api/settings/api_key` — frontend can paste key + model + budget;
  atomic write to `.env` with chmod 600 + hot-reload provider singleton.
  No restart needed.
- `.env` and `.env.local`, `*.key`, `secrets.json` gitignored.

**Lessons**:
- Pasting an API key in chat → permanent leak in transcript. The harness
  blocks subsequent operations using the leaked key (correct safety
  behavior). Always direct the user to put the key in `.env` or via
  the Settings UI — never type it in the conversation.
- `anthropic.RateLimitError` and `APIError` are catchable; `BudgetExceededError`
  is our own and propagates `{ok: False, budget_exceeded: True}`.
- `os.environ.setdefault(k, v)` doesn't override existing env vars
  (used by our `_load_dotenv`). Shell with `export ANTHROPIC_API_KEY=""`
  defeats `setdefault`. Settings POST uses direct `os.environ[k] = v`
  to force-update for hot-reload.

---

## Phase 16 — Acrobat-style layout (Phase A) (2026-05-10)

**Built** (commit `1099daa`):
- **Activity rail** (left, 48px column): VSCode-style icon-only
  navigation. Tree (default), Search → ⌘K, Learn, Versions, Audit,
  AI pane toggle, Theme, Settings, Help.
- **Context ribbon** (between topbar and content): mode tabs
  (Verify / Edit / Re-annotate / Apply Auto) with mode-specific
  sub-toolbars. Edit mode tab syncs with the legacy `EDIT_MODE` flag
  so existing keyboard shortcuts (V/R/T/Del/⌘Z) keep working.
- **AI pane** (right, 340px column, collapsible): persistent Claude
  assistant with sections:
  - Proposal: rendered Col D + confidence bar + rationale +
    inline `[Accept] [Edit] [Reject]`
  - Teach Claude: 6 quick hashtag chips (`#wrong-page`, `#brand-wrong`,
    `#missing-spec`, `#typo`, `#format`, `#commitment`) + free-text
    rationale + Send button → POST `/api/learn/feedback`
  - Recent: feedback count · accuracy · pattern count
- **Status bar** (30px row at bottom): row info + verdict pill +
  progress bar + Claude online status + spend-today + save state.
- **Embedded mode**: auto-detects `window.self !== window.top`,
  applies `--safe-top: 56px` to clear host-app overlay (Claude Preview
  MCP toolbar). User can override via Settings → "Top inset" slider.
- **Floating action buttons** (bottom-LEFT, 3 round): Search · Settings
  · Help — backup access in case the topbar is covered by an overlay.
- **Boot speedup**: `_build_pdf_records_for_db` no longer calls
  `list_pdf_annots()` per PDF (was ~60-90s on Google Drive). The
  `pdf_annotations` table is write-only/count-only — annots are
  fetched live via `/api/pdf_meta` when a row is selected.

**Lessons**:
- Google Drive iCloud-style filesystems make `.git/HEAD` operations
  time out (online-only files). Workflow: `git clone` to `/tmp/` (real
  disk), do work there, `cp` source files back to Drive. Never put
  `.git/` in cloud sync.
- `scrollIntoView({block: 'center'})` cascades to all ancestor
  scrollables — including `<html>` if its `overflow` isn't explicitly
  `hidden`. Pages were scrolling document-level on `J/K` row nav,
  pushing topbar off-screen. Fix: lock `html, body { overflow: hidden;
  overscroll-behavior: none }` and use `block: 'nearest'` on selectRow's
  scrollIntoView.
- Action bar flicker: changing `display: none → ''` re-triggers any
  `animation:` keyframe, plus child content height changes (textarea
  growing as user types) push the bar size around. Solution:
  `height: var(--action-bar-h)` fixed, `contain: layout size`,
  textarea pinned to single-line `28px` with `white-space: nowrap`.
- Topbar buttons get covered by any host-app top overlay (Claude
  Preview MCP, browser extensions). Solution: `--safe-top` CSS variable
  +`data-embedded` body attribute + Settings slider. Plus FAB at
  bottom-LEFT (out of any overlay's reach) for redundant access.

**What's next (Phase A continuation)**:
- A5: Acrobat-style floating annotation toolbar (above selected annot)
- A6: Move verdict 1-2-3-4 buttons from action bar to status bar
- A7: Remove duplicates (FAB + kbd-help redundancy)
- B3: Live AI co-edit Col D autocomplete while typing
- B5: Patterns triggered visualization in AI pane

---

## Phase 17 — WYSIWYG annotation render

**Problem**: Edit mode showed annotations differently from Preview.
Preview rendered yellow header banners, label backgrounds, custom
fonts — Edit mode showed plain red SVG rects + text on a stripped-bare
page. User: "ขึ้นไม่เหมือน".

**Root cause**: `/api/pdf_page?edit=1` was passing `no_annots=True`
to `render_pdf_page_png`, stripping all `/AP` (appearance streams).
The SVG overlay drew generic boxes, missing PyMuPDF's baked custom
appearances.

**Fix** (commit `1099daa` and earlier):
- `/api/pdf_page` always renders WITH annots baked in by default.
  `?bake=0` query param available for active-edit ghost-free rendering.
- SVG `buildAnnotNode` now draws visible rect/text only when the annot
  is `_isNew` (newly drawn, not on disk) or `selected`. Existing annots
  show through the baked image; SVG provides only invisible hit areas
  + handles for selected.
- CSS: removed `vector-effect: non-scaling-stroke` so SVG stroke
  matches PyMuPDF's 1pt baked stroke at any DPI. Removed dotted
  edit-only outline on FreeText (preview has no border, so neither
  should edit).

**Verification**: `edit page bytes === view page bytes` (byte-exact
identical render). Round-trip apply_pdf_edits returns
`applied=2 errors=0` for Square + FreeText pair.

---

## Phase A5 — Floating annotation toolbar (Acrobat-style)

**Goal**: When a user selects an annotation in edit mode, surface
the most useful actions right next to it — the way Acrobat shows
the small properties bubble above a selected shape.

**Built**:
- New element `#float-annot-toolbar` mounted on `<body>` (escapes
  every parent stacking context and the catalog scroll clip).
- Visible only when `EDIT_MODE === true` AND `SELECTED_ANN_ID`
  resolves to a live, non-deleted annot AND that annot's SVG node
  is currently on-screen inside the catalog viewport.
- Contents: type badge (Square / FreeText icon + label) · color
  swatch (red) · width meta (`1pt red` / `red text`) · Duplicate ·
  Delete. An arrow on the bottom edge points back at the annot
  (flips to top edge when toolbar lands below the annot).
- Position: above the annot (centered, viewport-clamped). Falls
  back to below when there isn't ≥ toolbar-height + 8 px of room
  above.
- Reposition triggers (via rAF-coalesced `updateFloatingToolbar`):
  - `refreshOverlay()` (every annot draw / select / drag / resize)
  - `setMode()` (mode tab switch hides immediately if mode != edit)
  - `window scroll` (capture phase, catches the catalog pane scroll)
  - `window resize`
- Keyboard: `D` duplicates the selected annot (offset by 12 pt,
  marked `_isNew` so the SVG overlay paints its red border).
  `Del` / `Backspace` already deletes; the toolbar exposes both.
- New `i-copy` icon symbol added to the SVG sprite.

**Why an out-of-DOM mount**: Earlier topbar-menu issue (Phase 13)
proved that any `position: absolute` inside `.topbar { z-index: 10 }`
gets clipped by the action-bar's stacking context. Same hazard
applies inside `.pdf-canvas`. Mounting on `<body>` with
`position: fixed` guarantees the toolbar floats over everything
the user can see, regardless of catalog scroll or pane z-index.

**Verification**: Python `ast.parse` + `node --check` of extracted
script block both pass. Manual smoke test deferred until next live
session — the user should select a Square in edit mode and confirm
the toolbar appears above with arrow pointing down at the annot,
follows on drag/resize/zoom/scroll, and disappears on Esc / mode
switch / page change.

**Deferred from this phase**: editable color/width are *displayed*
but not *editable* yet. Backend (`apply_pdf_edits`) hard-codes
red + 1pt to keep WYSIWYG with PyMuPDF's set_colors output. Making
them editable requires a coordinated change in both the SVG overlay
CSS and the `set_colors` / `set_border` calls in
`apply_pdf_edits`, plus a spec for which colors are even allowed
(red is currently load-bearing for `_assert_standard_appearance`
on legacy annots).
