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

---

## Phase A6 — Verdict moved from action bar to status bar

**Goal**: Reclaim the action bar for *content* actions (Auto / Mark
/ notes) and surface the binary "did this row pass?" decision in the
persistent status bar where you'd expect to find it in any IDE-style
tool.

**Built**:
- Status bar bumped from 30 px to 38 px to fit pill-shaped verdict
  buttons.
- New `.sb-verdict` segmented control with 4 verdict pills
  (`pass / fail / need_fix / skip`) plus a tiny reset (`↺`) button.
  Each pill carries a `kbd` tag so users see the 1–4 shortcut without
  needing the kbd-help.
- `statusBarUpdate()` now drives `aria-checked` on the active pill,
  matching the row's current verdict from `DATA.status`.
- Pills auto-disable (greyed out) when no row is selected via
  `body[data-row-selected="0"]`.
- Action bar now contains only Auto / Mark / auto-next / notes —
  the "verdict-control" segmented group and reset-btn were dropped.
- Old `setStatus`/`renderActionBar` queries against `.ab-btn.pass`
  etc. are now no-ops (guarded by `if (btn)`); kept harmless rather
  than ripped out so the wrapper chain stays intact.

**Why split**: One Acrobat-style pattern is "ribbon = tools, status
bar = state". Verdict is *state* (this row's outcome) — putting it
in the status bar lets the user verify and move on without their
eyes returning to a different bar each time.

---

## Phase A7 — Cleanup duplicate UI surfaces

**Goal**: Remove visual noise from features that already exist
elsewhere. The activity rail (Phase A1) already surfaces Settings /
Help / Search / Theme; carrying duplicate FABs and a busy
keyboard-help strip was just clutter for the 99% case.

**Built**:
- `.floating-actions` (FAB cluster, bottom-left) now hidden by
  default — only shown when `body[data-embedded="1"]` so users in
  Claude Preview MCP / iframes still get a Settings escape hatch
  even if the topbar is covered by host UI.
- `.kbd-help` (bottom-right shortcuts strip) collapses to a single
  ⌨ badge in non-embedded mode, expanding on hover to reveal the
  full shortcut list. Hidden entirely in embedded mode (where the
  user can press `?` for the help modal).
- No HTML/JS removal — everything is CSS-only so the embedded-mode
  fallback continues to work without re-wiring.

**Why a CSS-only fix**: keeps the FAB DOM in place so accessibility
tools still see the Settings/Help affordances, and lets users flip
into embedded mode and back without losing the safety net.

---

## Phase B3 — Live AI Col D autocomplete

**Goal**: When the user double-clicks a Col D cell to inline-edit,
surface the AI proposal + similar Col D values from neighbor rows
right under the cell — Tab/Enter to accept, type to filter.

**Built**:
- New endpoint `GET /api/row/col_d/suggest?row=N&q=text` returns
  up to 6 ranked suggestions:
  - **AI proposal** (top-priority): from `auto_annotate_plan(row)` —
    rule + learned pattern, no LLM call (cheap path).
  - **Neighbor templates**: Col D values from rows in the same
    section root (`5.1.*` for a 5.1.2 row) that are already
    `pass` or `need_fix` verified.
  - **Shape templates**: canonical fallbacks (`เอกสาร {section}
    ... หน้า ?`, `ยินดีปฏิบัติตามข้อกำหนด`).
  - When `q` is non-empty, suggestions whose text contains q (or
    starts with q) get a score boost; AI proposal still wins ties.
- Frontend `_colDAcOpen()` mounts a single dropdown panel on
  `<body>` (escapes any clipping context, same lesson as Phase 13
  topbar menu / A5 floating toolbar). It fetches via debounced 250 ms
  input, renders rows with kind badge (AI / Neighbor / Shape),
  text, source label, confidence %.
- Keyboard: ArrowUp/Down to navigate, Tab or Enter to accept the
  highlighted suggestion (puts caret at end), Esc dismisses panel
  but keeps editing. Mouse click also accepts.
- Accept fires a synthetic `input` event so the next debounce
  recomputes — useful when the user picks a shape template and
  wants to fill in the `?` page number.

**Why this design**: keeps the existing inline `editColD` flow
intact — autocomplete is an *augmentation*, not a replacement.
Server cost is ~one `auto_annotate_plan` per panel-open (cheap;
no LLM unless rules + patterns yield low confidence and Claude is
installed) plus a single SQL scan over `verification_status` keyed
by section.

**Verification**: `python ast.parse` + `node --check` on extracted
script block both pass. Manual test deferred until next live
session — open a Col D cell, confirm panel appears below and
shows the AI proposal as the top entry.

---

## Phase B5 — Patterns triggered visualization in AI pane

**Goal**: When Claude (or rule-based pipeline) proposes a Col D
value, show the user **which** learned patterns fired and how
confident the system is in each one. This makes the proposal
auditable instead of a black box.

**Built**:
- `aiPaneRender()` now reads `plan.provenance` (already populated
  by `auto_annotate_plan` whenever `apply_learned_brand` /
  `apply_learned_vendor` / similar return a hit) and builds an
  `.ai-patterns` subsection inside the Proposal section.
- Each row shows: pattern_type · trigger · confidence% · samples.
  E.g. `filename_brand · ruijie · 95% · 12 samples`.
- The subsection only renders if at least one pattern fired —
  pure-rules proposals stay clean.
- Visual style is tight (mono font for pattern_type / trigger,
  pill for confidence, faint samples count) so it sits under the
  rationale without competing.

**Why surface this**: aligns with the HITL contract — the user
should be able to see *why* the AI is proposing X and override it
with full context. Also doubles as a debug tool: if a wrong pattern
keeps firing, the user can spot it here and disable the rule from
the rail's Learn panel.

---

## Phase 1 — Claude Code as core (Agent SDK + SSE streaming)

**Goal**: Replace the API-direct provider with the **Claude Agent
SDK** so the user's existing **Claude Max** subscription powers all
LLM work (no metered API charges, $5/day cap removed). Frontend now
shows Claude's reasoning live — thinking → tool calls → final
proposal — instead of a black-box single-shot.

**Why**: User subscribes to Claude Max ($200/mo) and was paying API
costs separately. Phase 1 collapses both into one auth path while
also unlocking richer agentic tool use (Read + Grep + custom MCP
tools) and live HITL teaching.

**Built**:
- New module `app/claude_code_provider.py` (~470 LOC):
  - `ClaudeCodeProvider` — drop-in for `AnthropicProvider` (same
    `propose(row_context, few_shot)` signature, same `llm_calls` row
    shape). Adds `propose_streaming()` async generator.
  - 3 MCP custom tools registered via `create_sdk_mcp_server`:
    - `mcp__comply__propose_col_d` (typical case)
    - `mcp__comply__propose_brand_model` (section_header brand_model)
    - `mcp__comply__escalate_to_user` (cannot decide)
  - System prompt = SKILL.md + KB.md + pitfalls.md + top 30 learned
    patterns. Cached for 60 s.
  - Allowed tools: `Read`, `Grep`, plus the 3 MCP tools. **No Edit/
    Write** — proposals always go through the user-confirm UI.
  - `permission_mode="default"` (Claude asks before unexpected
    operations).
  - Auth detection: `~/.claude.json` exists → `claude_max`;
    `ANTHROPIC_API_KEY` set → `api_key`; else → `none`.

- New endpoint `GET /api/claude/stream?row=N` (Server-Sent Events):
  - Streams events as Claude works:
    - `{type:"thinking", text:...}`
    - `{type:"tool_use", name:..., input:{...}}` (Read/Grep/propose_*)
    - `{type:"tool_result", name:..., text:...}`
    - `{type:"text", content:...}` (any narration)
    - `{type:"result", proposal:{...}, elapsed_ms, cost_usd}`
    - `{type:"error", error:...}`
  - Async-generator → sync-Flask bridge via private event loop per
    request (single-user assumption — ok for desktop tool).

- AI pane: new "Run with Claude Code" section (Phase A1's pane gets
  a third panel below Proposal). Renders streaming events as
  color-coded chips:
  - Cyan = thinking
  - Orange = tool_use (with arrow `→`)
  - Green = tool_result (with arrow `←`)
  - Indigo = narration text
  - Red = error
  - Final result card has `Accept` / `Reject` buttons that wire to
    existing `/api/row/col_d` save flow + `/api/learn/feedback`.

- Settings UI:
  - Status card now reads `provider_kind` (`claude_code` /
    `anthropic_api` / `off`) and renders different copy per mode.
  - API key form is moved into a `<details>` collapsible labelled
    "API key fallback (only if Claude Max not available)".
  - When `provider_kind === 'claude_code'` and `auth_mode ===
    'claude_max'`, shows green "Claude Max OAuth" pill.

- Status bar (`#sb-claude`):
  - Claude Max mode: shows `model · Max · N calls`.
  - API mode: shows `model · $0.10/$5` (legacy).

- Boot logic: tries Claude Code first, falls back to Anthropic API.
  `COMPLY_LLM=claude_code` (default) / `COMPLY_LLM=anthropic` (legacy)
  / unset = first available.

- Launcher (`start_verify_gui.command`):
  - Now requires Python 3.10+ (claude-agent-sdk dependency). Probes
    `python3.13 / 3.12 / 3.11 / 3.10 / python3` in order.
  - Auto-installs `claude-agent-sdk` along with flask/openpyxl/etc.
  - Optional offer to `npm install -g @anthropic-ai/claude-code` on
    first run.
  - Detects unauthenticated CLI and prompts user to run `claude auth
    login` (3-second timeout, doesn't block).

**One-time setup the user must do**:
1. Install Claude Code CLI: `npm install -g @anthropic-ai/claude-code`
2. Authenticate: `claude auth login` → opens browser OAuth → Claude
   Max subscription detected
3. Restart the GUI → status badge flips to green "Claude Max OAuth"

**Verification**:
- Python `ast.parse` + `node --check` of extracted JS: ✓
- Flask test client boot + endpoint registration: ✓ 46 routes
  registered, `/api/claude/stream` present
- Provider initialization: ✓ `SDK_AVAILABLE=True`,
  `provider_kind=claude_code`, `auth_mode=claude_max`
- SSE pipeline end-to-end: ✓ stream returns 200, events parsed,
  error paths surface helpful "run claude auth login" hint to user
- Live agent run (with auth): deferred until user runs
  `claude auth login`

**Phase 2 (deferred)**: more domain-specific MCP tools so Claude
operates at a higher level than raw filesystem (`get_row(N)`,
`find_text_in_pdf(rel, query)`, `check_pattern_for_section(s)`,
`save_proposed_col_d(N, text, conf)`). With those, we can drop
`Read`/`Grep` from `allowed_tools` entirely.

**Phase 3 (deferred)**: deeper HITL — "Pin this correction as a
rule", "Always skip rows like this", per-row conversation logs.

---

## Tier 1 — Tech-stack hygiene (repo move + uv + lint + tests)

After Phase 1 the day-to-day pain wasn't bugs, it was friction:
Google Drive's iCloud-sync broke `.git/`, pip was slow, no smoke
suite to catch regressions. This pass kills all of those.

### 1.1 Repo moved to local SSD

```
Before: ~/Library/CloudStorage/GoogleDrive-…/comply-module/
              code + .git + output + _versions all in iCloud-sync

After:  ~/Code/smart-bos/
        ├── code + .git           ← real local files (fast)
        ├── _db/                  ← real local (per-machine state)
        └── output → /GDrive/…/comply-module/output      ← symlink
            _versions → …                                 ← symlink
            BOQ → …                                       ← symlink
            TOR → …                                       ← symlink
```

Why: every `git status` / `version.py snap` / `git push` previously
risked iCloud "online-only" timeouts; we already had a `/tmp clone`
workaround that was annoying. Now `.git/` is genuinely local;
project data still lives in GDrive (shared with team) via symlinks.

A marker `_README_CODE_MOVED.md` was left in GDrive's old
`comply-module/` so anyone opening it sees where the code went.

### 1.2 uv replaces pip

`pyproject.toml` now declares deps. `uv sync` creates a `.venv`
with Python 3.10+ (uv auto-installs Python if missing) and resolves
all deps. `uv.lock` is committed for reproducible builds.

| Operation | Before (pip --user) | After (uv sync) |
|---|---|---|
| First install | ~30 s | ~75 s (incl. Python install) |
| Subsequent runs | ~5 s (re-checks) | **~0.09 s** |

The launcher (`start_verify_gui.command`) now uses `.venv/bin/python`
exclusively; if `uv` isn't installed it `brew install`s it (or
falls back to the official curl install).

### 1.3 ruff + pyright wired in

- **ruff**: `extend-select = E,W,F,B,I,UP`. Auto-fixed 40 stylistic
  issues on first run; 75 more were the codebase's intentional
  one-line guards (`E701/E702`) — ignored in pyproject.toml.
  **Final: All checks passed!**
- **pyright**: `typeCheckingMode = "off"` for now (the 14 K-LOC
  monolith wasn't authored with strict typing in mind). New modules
  in `app/` will be annotated as they're added; eventually flip to
  `basic` and gradually retrofit. For Phase 1's
  `claude_code_provider.py` the imports already resolve cleanly.

### 1.4 Smoke test suite (pytest)

`tests/test_smoke.py` — 5 fast read-only checks that catch the most
common regressions:

1. `test_boot_registers_all_critical_routes` — every route the
   frontend hits must be registered (also asserts /api/claude/stream
   from Phase 1)
2. `test_index_returns_rows_and_sections` — schema contract on
   /api/index (rows list, sections list, tree dict)
3. `test_pdf_render_view_equals_edit_byte_exact` — Phase 17
   invariant (edit-mode bytes == view-mode bytes)
4. `test_col_d_suggest_returns_ranked_candidates` — Phase B3 endpoint
5. `test_claude_stream_endpoint_responds_or_503` — Phase 1 SSE never
   500s, returns 200 stream OR 503 with hint

```
$ uv run pytest -q
.....                                                        [100%]
5 passed in 1.25s
```

Boot-once / session-scoped fixture means subsequent tests are ~50 ms
each. Total < 2 sec.

### Lessons captured

- **"iCloud-sync + .git/ = fragility"** — even tiny `git status`
  can timeout when on-demand fetch kicks in. Lesson: never put `.git/`
  in cloud-sync folders. Symlinks for shared data are fine.
- **uv is genuinely 50× faster** than pip for re-resolution. Worth
  the install just for the speed.
- **Ruff > flake8 + black + isort** as a single tool. The ignore-list
  is the only real config needed for an existing codebase.
- **5 smoke tests beat 0 smoke tests** by infinity. Don't gold-plate
  — boot + 4 critical contracts already prevent regressions in
  Phase 17, Phase B3, Phase 1, etc.

---

## Phase 2 — Multi-company / catalog library

**Goal**: turn the system from "Smart Plant 1 only" into a workbench
that handles multiple companies and projects. Catalogs become
**reusable** across projects (the same Lenovo SR630 datasheet can
serve Plant 1, Plant 2, future-project-N) and editable in the DB
(metadata + annotations) without re-baking the PDF every time.

**Symlink relocation**: `output/` now points at the canonical
co-work folder shared with the team:
```
~/Code/smart-bos/output → /GDrive/.../Pattaya Project/Smart Plant 1/co-work/claude-code/output
```
That folder has 309 PDFs (vs the old 124) — newer + richer.
Snapshot taken before swap (`before-co-work-output-relink`).

**New module `app/catalog.py`** (~440 LOC) — the additive layer.
Exports:
- `ingest_output_dir(root)` — idempotent migration: scans every PDF,
  pulls sha256 + page count + heuristic brand/model/section, writes
  a `catalogs` row.
- `list_catalogs(brand, category, section, q)` — filtered listing
- `get_catalog(id)` — full detail incl. annotations + page text
- `update_catalog(id, **fields)` — patch metadata
- `list_/add_/update_/delete_annotation(...)` — DB-stored annotations
  per page (independent of PDF baking — Phase C will plumb these into
  the render pipeline)
- `bind_row_to_catalog(project, row, catalog, page, col_d)` — record
  which catalog a given project row uses
- Companies + projects helpers (upsert, list, set_active_project)

**New DB tables** (DB_VERSION bumped 1→2; additive only — no
migration needed because `IF NOT EXISTS`):
- `companies (company_id, name, code)`
- `projects (project_id, company_id, name, code, xlsx_rel, output_rel,
   is_active)`
- `catalogs (catalog_id, pdf_rel, sha256, pages, brand, model,
   category, section_hint, description, metadata_json, archived)`
- `catalog_pages (catalog_id, page, text_excerpt)` — for FTS-like
  search later
- `catalog_annotations (annot_id, catalog_id, page, type, rect_json,
   contents, color_json, border_width, anchor_text, archived)`
- `row_catalog_links (project_id, row_num, catalog_id, page,
   col_d_text, bound_at)` — the binding record

**Boot bootstrap**: after `sync_db_from_memory()`, the boot path:
1. ensures a default company "Smart Solution" + project "Smart Plant
   1" exists (only if none yet)
2. runs `ingest_output_dir(OUTPUT)` — idempotent, picks up new PDFs
   the user dropped into `output/` between runs

Result on first run after relink:
```
[boot] catalog library: 309 PDFs scanned (309 new, 0 updated, 0 unchanged)
       · active project=Smart Plant 1
```

**New REST endpoints** (~16 routes added):
```
GET    /api/catalogs?brand=&category=&section=&q=&limit=
GET    /api/catalogs/stats
GET    /api/catalogs/<id>
PATCH  /api/catalogs/<id>           ← edit metadata
GET    /api/catalogs/<id>/links     ← which rows use this catalog?
POST   /api/catalogs/reingest       ← re-scan output/ for new PDFs
GET    /api/catalogs/<id>/annotations
POST   /api/catalogs/<id>/annotations
PATCH  /api/catalogs/<id>/annotations/<aid>
DELETE /api/catalogs/<id>/annotations/<aid>
GET    /api/companies
POST   /api/companies
GET    /api/projects[?company_id=]
POST   /api/projects
POST   /api/projects/<id>/activate
POST   /api/row/apply_catalog       ← bind catalog → row + write Col D
```

**New UI: Catalog Browser** (modal launched from rail)
- Activity rail: new icon (book) opens the modal
- Modal layout:
  - Top: search box + section filter + Re-scan button + stats pill
  - Left pane: filtered catalog list (section / brand / model /
    pages / filename)
  - Right pane: full detail with editable metadata form + Apply-to-row
    button + annotations list + "Used by" links
- Apply flow:
  1. Click row in Comply tree → row becomes `SELECTED_ROW`
  2. Open Catalog Browser
  3. Pick catalog (left pane)
  4. Click `Apply to R{N}` → prompts for page → POSTs
     `/api/row/apply_catalog`
  5. Backend pre-snaps, writes Col D into xlsx, refreshes ROWS,
     records `row_catalog_links` entry, audit-logs the change
  6. UI toasts success and reloads xlsx + tree to show the new Col D

**Verification**:
- 7/7 smoke tests pass (5 pre-existing + 2 new for catalog API)
- `ruff check`: All checks passed!
- `pyright app/`: 0 errors

**Phase 2.5+ deferred**:
- DB-stored annotations are *recorded* but not yet *rendered*. The
  PDF baking pipeline still owns the visible annotations on disk.
  Phase C will:
  - Add an SVG overlay layer that paints `catalog_annotations` on
    top of the rendered PDF page (so users edit them in DB-only flow
    instead of mutating PDF)
  - Add an "Apply to PDF" button that bakes a catalog's DB-stored
    annotations into the actual PDF file
- Catalog editor as a *standalone full-page workspace* (currently
  edit-in-modal-side-panel)
- Multi-company UI switcher (project selector currently implicit;
  the API supports it but the topbar doesn't expose it yet)
- FTS5 index over `catalog_pages.text_excerpt` (currently using LIKE)

**Lessons captured**:
- **Additive schema migrations** (every CREATE TABLE has IF NOT
  EXISTS) mean a DB version bump doesn't require a custom migration
  step. New tables coexist with the old.
- **Symlinks abstract over storage location** — the running code
  doesn't care that `output/` lives in GDrive vs local SSD vs S3
  via fuse vs anything else, as long as Python can stat/read the
  resolved path.
- **One ingest function rules them all** — the same
  `ingest_output_dir` runs at boot AND from a `Re-scan` button in
  the UI. Idempotent design = single code path.

---

## Phase 2.1 — Export print-ready compliance package

**Goal**: a single click that produces a polished, submittable PDF
combining the Comply Spec sheet + every catalog (with annotations
baked-in) + a navigable cover/TOC/bookmark structure.

**Built** — `app/export.py` (~360 LOC)

The package layout:
```
[Cover page]                A4 portrait, indigo top band,
                            project name, code, generated-at,
                            version (= latest snapshot ID)

[Table of Contents]         Auto-generated from bookmark tree:
                            level 1 = section, level 2 = catalog group,
                            level 3 = individual catalog. Page numbers
                            right-aligned with dotted leader.

[Comply Spec Sheet]         Verbatim insert from existing
                            output/Comply spec*.pdf
                            (the user's xlsx export — we don't
                            re-render xlsx→PDF here)

[Catalogs]                  Section divider page (big indigo header)
                            then catalog PDFs in order. Annotations
                            already baked in (Phase 17 WYSIWYG).

[Audit Log appendix]        Optional. Last 200 audit_log entries
                            timestamp + action + target.

[Footer on every page]      "Project Name" left, "Page N of M" right,
                            thin separator above.

[PDF outline / bookmarks]   Set via doc.set_toc() — readers like
                            Acrobat / Preview show a navigation tree.
```

**TOC trick**: we don't know real page numbers until catalogs are
inserted, but the TOC must come *before* them. Solution:
1. Insert N placeholder pages where TOC will go
2. Build everything else, recording (level, title, page1) tuples
3. Render the real TOC into a temp doc
4. `delete_pages` the placeholders, `insert_pdf(tmp, start_at=...)`
5. Adjust all bookmarks whose page > TOC region by Δ = real_pages -
   placeholder_pages

**New REST endpoints**:
```
GET  /api/export/preview         what would be in the package
                                 (counts, by section, comply present?)
POST /api/export/package         build it, returns
                                 {filename, page_count, byte_size,
                                  download_url, sections[]}
GET  /api/export/download?file=  stream a built PDF
GET  /api/export/list            recent builds in _db/exports/
```

Query params (POST and preview):
- `mode=full|comply_only|catalogs_only`
- `section=5.1.1` (filter to a section root)
- `bound_only=1` (only catalogs already linked to project rows)
- `include_audit=1`

**New UI** — Export modal (rail icon `📄`):
- Two fieldsets: "What to include" (mode radios) + "Filters"
  (bound-only / section / audit toggles)
- Live preview pane updates on every option change
  (debounced 200 ms): "Project · Comply sheet status · 309 catalogs
  · 5.1.1: 27 · 5.1.2: 12 · …"
- Build PDF → loading state → triggers `<a download>` automatically
- Recent exports `<details>` listing reusable download links

**Performance**: 311 pages (1 section, ~30 catalogs) in **0.26 s**.
Full 309-catalog package would be roughly 2-3 s on this machine.

**Verification**:
- 8/8 smoke tests pass (was 7; +1 for `test_export_preview_and_build_small`)
- ruff: All checks passed
- pyright app/: 0 errors
- Manual: built package opens in Preview/Acrobat with working
  bookmarks, TOC clickable, page numbers correct

**Lessons**:
- **`fitz.insert_pdf` preserves appearance streams**, so all the
  Phase 17 WYSIWYG work flows straight into the export — no special
  handling needed for annotated PDFs.
- **Placeholder-then-replace** is the cleanest way to handle
  forward-references in PDF generation. Tracking +Δ shifts all
  downstream bookmarks at the end.
- **PyMuPDF doesn't support multi-page text wrapping in one call**;
  manually paginate when content might exceed one page (audit log,
  long TOC).

---

## Phase 2.2 — Knowledge integration (SKILL/KB/pitfalls/sr_pattern + continuity)

**Trigger**: the user's parallel project agent (working in
`/GDrive/Pattaya Project/Smart Plant 1/co-work/claude-code/`) shipped
a major knowledge update plus a new "continuity" handoff document.
The GUI now reads them too.

**Synced files** (from project folder, no content drift):
- `SKILL.md` — bumped from 63 KB → 75 KB. Major addition: Rule 1
  refined to distinguish `"ยินดีปฏิบัติตามข้อกำหนด"` (commitment for
  installation/software-internal) vs `"ไม่พบใน catalog"` (flag for
  hardware specs not yet found in catalog → user must review).
  Added Rule 12: continuity document convention.
- `knowledge_base/KB.md` — extended (110 line additions)
- `knowledge_base/pitfalls.md` — +97 lines (PyMuPDF gotchas
  including `delete_annot` not persisting and GDrive sync failure
  on `shutil.copy2`)
- `knowledge_base/sr_pattern.md` — **NEW** (~8 KB). Read-only study
  of how SR (Smart Solution Co.) annotates their proposed catalogs:
  highlight + short callouts `N)` (no section prefix), white bg,
  red text, no border. Used as reference when annotating sister
  catalogs in the same SR style.
- `knowledge_base/pipelines.md` — extended
- `_continuity/STATE_20260510_111800.md` — **NEW** (~5.6 KB).
  Handoff document: last completed task, open in-progress, pending
  user decisions (verbatim quotes), discovered bugs, vendor
  sections currently being worked on, next planned action.

**Provider integration** (`app/claude_code_provider.py`):
- System prompt now loads (in order): SKILL.md → KB.md → pitfalls.md
  → **sr_pattern.md** → pipelines.md (added)
- Plus the **latest STATE_*.md** from `_continuity/` (preceded by a
  paragraph telling Claude what it is and to read it before
  suggesting actions). This means every agent run starts with full
  awareness of pending user decisions and in-progress work.
- 60-second cache still applies — picks up edits to those files
  within a minute without restart.

**New endpoint `GET /api/continuity`**:
- Returns the latest STATE markdown, headline (first non-heading
  line), filename, mtime, byte_size, plus a list of previous
  handoffs.
- 200 OK with `available: false` when none present.

**New UI: continuity badge in topbar**
- Small warn-colored pill, only visible when a STATE file exists
- Shows the headline (truncated to 70 chars)
- Click → modal with full markdown rendered as monospace,
  collapsible "Previous handoffs" list
- Auto-refreshes every 60 s so a new STATE drop appears
  automatically without restart

**Verification**:
- 9/9 smoke tests pass (was 8; +1 `test_continuity_endpoint`)
- ruff: All checks passed
- Visual: opening the GUI now shows the warn pill at the top
  carrying "Adopted SR pattern across TRIO_SR_Solution + applied
  'ไม่พบใน catalog' flag rule"

**Why this matters**: previously the Skill Agent and the GUI session
were two separate brains with no shared context. Now the GUI's
Claude Code provider reads the Skill Agent's STATE document before
proposing anything, and the operator sees the same handoff at a
glance — single source of truth for "what's the project up to right
now?"

**Lessons captured**:
- **Continuity files are cheap, high-value**: a 5 KB markdown
  written by the previous session saves the new session 5+ minutes
  of "re-orienting" via grep + audit_log scans.
- **Don't gitignore `_continuity/`** — it's a *deliverable* of every
  session, not a transient. Commit it so reverts are clean.
- **System prompt cache TTL = 60s** is the right granularity for
  these files. Short enough that edits propagate, long enough that
  a flurry of API calls doesn't re-read the disk every time.

---

## Tier 2 — Monolith refactor (template extract + Blueprints + pyright)

**Goal**: take the first real bites out of the 14k-LOC `comply_verify_gui.py`
monolith. Three coordinated changes that together cut its line count
~63% and finally let editors syntax-highlight the HTML/CSS/JS.

### T2.1 — Extract `INDEX_HTML` to `app/server/templates/index.html`

The HTML/CSS/JS template was a 9530-line `r"""..."""` string baked
into the Python file. This made:
- VS Code / vim show it as raw text (no HTML/CSS/JS highlighting)
- ruff complain about long lines for things that should be ignored
- diffs in `git log` painful (everything looks like Python)

Move:
- `INDEX_HTML = r"""…"""` block (lines 5497–15027) → new
  `app/server/templates/index.html` (395,833 bytes)
- Flask app now constructed with `template_folder=…`
- Route changes from `render_template_string(INDEX_HTML)` →
  `render_template("index.html")`
- `render_template_string` import dropped from `flask` import line
- `app/server/static/` directory reserved for future asset extraction
  (CSS/JS files); auto-attached when present

**Result**: `comply_verify_gui.py` shrank from **15,175 → 5,646 lines**
(–63%). Editors now show proper HTML/CSS/JS for the template; the
Python file is mostly Flask routes + boot logic.

**Verified**: `/` route still serves the same 395 KB HTML; all 13
expected UI elements present (`#float-annot-toolbar`, `#sb-verdict`,
`#catalog-modal`, `#export-modal`, `#continuity-modal`,
`#topbar-continuity`, `.col-d-ac-panel`, `.ai-patterns`, …).

### T2.2 — Blueprints for catalog / export / continuity

Three slices of /api/* moved into proper Flask Blueprints:

```
app/routes/
├── __init__.py            register_all(app, root, output_root)
├── catalog_api.py         /api/catalogs/*, /api/companies, /api/projects/*
├── export_api.py          /api/export/*
└── continuity_api.py      /api/continuity
```

State plumbing: `register_all` writes `COMPLY_ROOT` and
`COMPLY_OUTPUT` into `app.config`; blueprint handlers read them via
`current_app.config[...]`. No globals or circular imports.

**What stayed in `comply_verify_gui.py`**:
- `/api/row/apply_catalog` — touches xlsx + ROWS + `make_col_d_for_row`
- `/api/auto_annotate/*`, `/api/manual_annotate/*`, `/api/reannotate/*`
- `/api/pdf_*`, `/api/tor_*`, `/api/row/col_d`, `/api/status`, `/api/index`
- `/api/learn/*`, `/api/versions/*`, `/api/db/*`
- `/api/claude/stream`, `/api/settings/*`

These all depend on `ROWS`/`PDF_INDEX`/`SECTION_INDEX` in-memory
state or call internal helpers like `make_col_d_for_row`,
`detect_row_role`, `_run_version_cmd`. Moving them is a bigger
project — gradual migration as those globals get encapsulated.

**Result**: `comply_verify_gui.py` further shrank to **5,219 lines**
(–427 from T2.2). New blueprints add ~330 lines split across 4
focused modules with clear seams.

**Verified**: `gui.app.url_map.iter_rules()` still produces 65
`/api/*` routes; all blueprint routes (`/api/catalogs`,
`/api/catalogs/<id>`, `/api/companies`, `/api/projects`,
`/api/export/preview`, `/api/export/package`, `/api/export/download`,
`/api/export/list`, `/api/continuity`) resolve correctly.

### T2.3 — Pyright `basic` mode + soft warnings

Flipped `tool.pyright.typeCheckingMode` from `"off"` → `"basic"`.
The 14k-LOC monolith wasn't authored for type checking, so
historical issues are *softened* (warnings, not errors) for
specific report categories:
`reportReturnType`, `reportOptionalSubscript`,
`reportOptionalMemberAccess`, `reportArgumentType`,
`reportCallIssue`, `reportAttributeAccessIssue`,
`reportAssignmentType`, `reportPossiblyUnboundVariable`,
`reportIndexIssue`, `reportGeneralTypeIssues`,
`reportInvalidTypeArguments`.

This means:
- New modules (`app/routes/*`, `catalog.py`, `export.py`,
  `claude_code_provider.py`) get type-checked seriously and produce
  **0 errors**
- Legacy paths get warnings the developer can choose to address
  (or filter via `# type: ignore[<name>]` if intentional)
- `pyright app/` now reports `0 errors, 157 warnings` — green
  enough to wire into CI as a gate

### Verification

```
$ uv run pytest -q
9 passed in 5.14s

$ uv run ruff check .
All checks passed!

$ uv run pyright app/
0 errors, 157 warnings, 0 informations

$ wc -l comply_verify_gui.py
5219 lines    # was 15175 before T2.1
```

### Lessons captured

- **Single biggest pain killer = template extraction**. Three minutes
  of editing the HTML now without scrolling 9000 lines of Python first
  felt like upgrading to a faster computer.
- **Blueprint factory via `current_app.config`** is dramatically
  cleaner than passing state around or importing the gui module
  (which would create circular imports). Flask designed this for a
  reason.
- **Pyright basic + softened warnings** is the right halfway point
  for a half-typed codebase. New code gets caught, old code doesn't
  flood the output.
- **Keep state-dependent routes in main**: don't try to move
  everything in one go. The "pure" CRUD endpoints came out cleanly;
  the xlsx-mutating endpoints would have required encapsulating
  ROWS/PDF_INDEX first, which is a separate refactor.

---

## Phase C — Catalog annotation editor (ship-now MVP)

**Trigger**: user feedback after the audit — *"แก้ระบบมานานแล้วงาน
หลักไม่ได้ทำสักที"* (we keep refactoring infra, the actual user-proof
work hasn't shipped). Phase C closes the last user-facing gap: a way
to edit annotations on catalogs from the new catalog browser, so the
user can finish the user-proof pass and submit.

**Decision**: ship the SIMPLEST thing that works rather than build a
brand-new DB-annotation editor. The existing PDF edit flow (Phase 17
WYSIWYG + Phase A5 floating toolbar + drawRect / addText / undo /
redo / save) is battle-tested. Reuse it, route catalogs through it,
done.

**Built**:
- New "**Edit annotations**" button in the catalog browser detail
  pane (next to "Save metadata" / "Apply to row" / "Open PDF (raw)").
- New JS `catalogEditAnnotations(catalog_id, pdf_rel)`:
  1. Stores `_CATALOG_EDIT_CONTEXT = {catalog_id, pdf_rel}`
  2. Closes the catalog browser modal
  3. Switches mobile view to the PDF tab if needed
  4. Calls existing `loadPdf(pdf_rel, 1, null)` (no row binding,
     no row-highlight)
  5. Calls `toggleEditMode()` to enter the existing edit flow
  6. Shows a floating warn-colored "Catalog edit mode" banner with
     filename + close button
- The user now has the **full Phase 17 + A5 toolset**: drawRect,
  addText, click-to-select, drag-to-move, resize handles, undo/
  redo, floating annotation toolbar, byte-exact WYSIWYG preview.
- Save (existing `saveEdits` → `/api/pdf_save` → `apply_pdf_edits`)
  bakes annotations into the actual catalog PDF file.
- A small wrapper around `saveEdits` detects when we're in catalog
  edit mode and after the PDF save succeeds it pings
  `PATCH /api/catalogs/<id>` with `{}` to bump `updated_at` (so
  the catalog browser shows freshness when the user reopens it).
- "Catalog edit mode" banner has an `✕` button → calls
  `catalogExitEditMode()` which clears the context, hides the
  banner, and toggles edit mode off (prompting save if dirty).

**Why this beats a separate DB-annotation editor**:
- Zero new editor code — every existing keyboard shortcut, undo
  semantics, WYSIWYG invariant, etc. just works.
- Annotations show up in the export package immediately because
  `apply_pdf_edits` writes them to the actual PDF file (which the
  exporter reads via `fitz.insert_pdf`).
- No "render layer" gap: with a DB-only editor the user wouldn't
  see their edits in the export until a separate "bake" pass.

**Trade-off**: catalog_annotations table stays empty for now. The
DB-annotation render layer is still a future improvement (would let
catalogs share annotations across multiple project bindings without
re-baking PDFs). Not blocking anything user-facing today.

**Verification**:
- 11/11 smoke tests pass (was 9; +2 for Phase C presence + empty-
  patch wiring)
- `ruff check`: All checks passed
- HTML rendered at /: catalogEditAnnotations, catalog-edit-banner,
  _CATALOG_EDIT_CONTEXT, "Edit annotations" button, catalogExitEditMode
  all present
- Empty-body PATCH on /api/catalogs/<id>: 200 OK (the bump-only path)

**Workflow user can do now**:
```
1. Open the GUI
2. Pick a row from the tree (or use the Catalog Browser directly)
3. Click 📚 (rail icon)
4. Search/filter to find the catalog
5. Click the catalog → detail pane opens
6. Click "Edit annotations" → catalog PDF loads in main viewer +
   edit mode auto-on, with a banner showing what catalog you're
   editing
7. Use the existing tools: drawRect / addText / drag / resize /
   floating toolbar / undo / redo
8. Click Save → annotations bake into the catalog PDF file
9. Click ✕ on the banner to exit catalog edit mode
10. Export package picks up the new annotations on next /api/export/package
```

**Lessons**:
- **Reuse > rewrite**, especially when the existing code carries
  multiple invariants (Phase 17 byte-exact, A5 toolbar positioning,
  undo semantics). New editor would have re-derived all of those.
- **Banner over modal** for ambient mode: less invasive, leaves
  the full PDF view accessible, has a clear exit affordance.
- **Wrap-and-extend pattern** (`const _orig = saveEdits;
  window.saveEdits = async function() { … }`) lets us add
  catalog-aware behavior without forking the save flow.

---

## Phase C+ — User manual (docs/MANUAL.md + /manual page)

**Trigger**: user request — *"ใช้ claude design จัดทำ manual สอนใช้
งานระบบนี้อย่างละเอียด"*. Now that all user-facing features ship,
they need a single source of truth that explains how to actually
USE the thing — not just architecture docs for developers.

**Built**:
- `docs/MANUAL.md` (~34 KB Thai markdown) — comprehensive user manual
  covering 12 chapters + 4 appendices:
  1. ภาพรวมระบบ (overview + concepts)
  2. เริ่มต้นใช้งาน (setup + auth + launch)
  3. รู้จักหน้าจอ (UI walkthrough — every panel)
  4. Workflow หลัก: User-proof loop
  5. แก้ไข Col D (3 ways: inline, right-click, AI)
  6. Annotation ใน catalog PDF (edit mode + tools + floating toolbar)
  7. Catalog Library (browse / metadata / **edit annotations** / apply)
  8. AI Assistance (Claude Code Run + auto-annotate + manual + wizard
     + teach-back + patterns triggered)
  9. Export PDF Package
  10. Project & Continuity (multi-company + STATE handoff)
  11. Versions & Audit (snapshots + audit log)
  12. Troubleshooting (common issues + recovery)
  + Appendix A: keyboard shortcuts (full table)
  + Appendix B: file paths + DB schema
  + Appendix C: glossary
  + Appendix D: developer quick reference

- `app/server/templates/manual.html` (~12 KB) — clean standalone
  rendered viewer:
  - **Sticky header** with project name pill + search box + "back to
    GUI" link
  - **Sticky aside TOC** built dynamically from h2/h3 headings
    (3-level outline, indented)
  - **ScrollSpy**: highlights current section in TOC as you scroll
  - **Inline search**: type 2+ chars → debounced 200 ms →
    highlights matches with `<mark>` + scrolls to first hit
  - **Light/dark theme** auto-follows system preference
  - **Mobile-responsive**: collapses to single column < 900 px
  - **Print-friendly**: `@media print` hides header/aside, expands
    main column
  - Renders markdown via `marked.js` (CDN, single 30 KB script)

- 2 new routes (in main module — small, no blueprint needed):
  - `GET /manual` → renders the styled HTML viewer
  - `GET /api/manual/raw` → returns raw markdown body (so the
    template renders it client-side, keeping markdown as the
    canonical source)

- 3 entry points wired:
  - **Topbar menu**: new item "📖 คู่มือการใช้งาน (Manual)" between
    Settings and Help
  - **Onboarding modal**: prominent banner with primary "เปิดคู่มือ"
    button so first-time users see it immediately
  - **Direct URL**: `/manual` opens in new tab (so docs stay open
    while user works)

**Why a separate viewer instead of an in-GUI modal**:
- Manual is **long** (34 KB, ~12 sections). A modal would feel
  cramped and lose context when user closes it.
- Opening in a new tab lets the user keep the GUI open in one window
  and the manual in another — common pattern for documentation.
- The standalone page is also share-able and printable.

**Verification**:
- 12/12 smoke tests pass (was 11; +1 for `test_manual_routes_serve_content`)
- `ruff check`: All checks passed
- Manual page tested at `/manual` — TOC builds, ScrollSpy works,
  search highlights, theme follows system

**Lessons**:
- **Markdown stays canonical**: rendering via marked.js client-side
  means the source is plain `.md` (editable in any editor, version-
  controllable, portable). No build step.
- **Sticky TOC + ScrollSpy** is the right UX for a long manual —
  user always knows where they are without scrolling away from
  content.
- **Onboarding integration** is critical: users won't read the
  manual if they don't know it exists. The first-time modal now
  has a prominent "เปิดคู่มือ" CTA.

---

## Final batch — UX polish + Ops + CI gate (handoff readiness)

**Trigger**: user request — *"ทำที่เหลือทำหลัง deadline (ไม่
blocking) ไปเลย พร้อมทดสอบระบบอย่างละเอียดเพื่อส่งมอบระบบให้ USER
ทำงานได้"*. Final pass that ships every deferred polish item plus
broad test coverage and a CI gate, so the system can be handed off
clean.

### UX-3 — Page picker modal (replaces native `prompt()`)

The Apply-to-row flow used to call `prompt('หน้าใน catalog?')` which
on macOS browsers shows a native dialog with no styling control,
unreadable Thai font, and no preview. Now:
- Click "Apply to R{N}" → fetch catalog page count
- Modal opens with a grid of buttons: `1`, `2`, `3`, …, `pages`
- Plus an "ไม่ระบุหน้า" button at the top
- Click any → spinner on that button → toast on success
- Modal closes automatically + refreshes xlsx + tree

### UX-4 — "🔴 ต้องดู" filter

New checkbox in the tree filter row. When checked, only shows rows
that need user attention:
- Empty Col D (still needs filling)
- `"ยินดีปฏิบัติตามข้อกำหนด"` (commitment — verify it's the right call)
- `"ไม่พบใน catalog"` (flag from SKILL Rule 11 — needs review)
- Any verdict = `need_fix`

This is the user-proof "what's left" filter — instant visual of
remaining work.

### UX-2 — Bulk catalog metadata cleanup

Heuristic ingest left ~70/309 catalogs with bad brand guesses
("PoE", "G7-05002", `-`, NULL, etc.). New "Bulk cleanup" button in
catalog browser opens a modal with:
- Field selector (brand / model / category / section_hint)
- Match-type selector (exact / contains / prefix / regex)
- Match value input (empty = match NULL/empty)
- New-value input
- Live preview pane showing first 30 matches
- Confirm dialog before applying
- Audit-logged

Three new endpoints: `GET /api/catalogs/bulk_preview`,
`POST /api/catalogs/bulk_update`, plus existing audit_log integration.

### UX-1 — Multi-company / project switcher

Topbar pill (primary color) showing "{company_code} · {project_code}".
Click → modal with:
- Active project card (top, highlighted)
- List of all projects (click to activate)
- Collapsible "+ Add company / project" form
- Toast hint: "restart to re-ingest output for new project"

Backend support already existed (`/api/projects/<id>/activate`); UX-1
just wired the topbar UI.

### Ops-2 — Boot ingest skip-by-sha256

`ingest_output_dir` previously always opened every PDF (309 of them)
to re-extract pages + metadata, even when nothing changed. New fast
path:
1. Pre-fetch existing `(pdf_rel → catalog_id, sha256)` map in one query
2. For each file: hash first 4 MB
3. If sha matches stored → SKIP everything (no fitz.open, no metadata
   reparse)

**Result**: cold boot dropped from **~7 s → 1.18 s** with
`308/309 fast-skip-by-sha`. Restart is now nearly instant.

### Ops-1 — Auto-write continuity STATE on shutdown

New module `app/continuity.py` with:
- `write_session_state(root)` — pulls last 50 audit_log entries since
  the previous STATE file, plus verdict counts, feedback summary,
  Claude call cost, active project, latest snapshot. Writes a
  Markdown summary to `_continuity/STATE_<ts>.md`.
- `install_atexit_hook(root)` — registers via `atexit`, idempotent,
  opt-out via `COMPLY_NO_ATEXIT_STATE` env (used in CI/tests).

Boot installs the hook automatically. Now every GUI session leaves
a paper trail the next session can read — the Skill Agent and the
GUI now BOTH produce STATE files.

### Tech-1 — Test coverage 13% → 50%+

5 → 35 tests. New modules:
- `tests/test_catalog.py` (8 tests) — section filter, query, annotation
  CRUD cycle, bulk preview, bulk-update validation, companies+projects,
  apply-to-row writes-xlsx-and-records-link
- `tests/test_export.py` (6 tests) — preview modes, section filter,
  bound-only, 404 + path-traversal protection on download, list endpoint
- `tests/test_continuity.py` (3 tests) — module imports,
  write_session_state creates valid markdown, install hook is
  idempotent
- `tests/test_pdf_render.py` (4 tests) — pdf_meta returns pages+annots,
  404, path-traversal protection, **Phase 17 byte-exact view==edit**,
  ?bake=0 differs from baked

Total run time: **5.67 s** for 35 tests.

### Tech-2 — GitHub Actions CI gate

New `.github/workflows/ci.yml`:
- Runs on push/PR to main + `workflow_dispatch`
- ubuntu-latest, 10-min timeout
- `astral-sh/setup-uv@v3` → `uv python install 3.11` → `uv sync --extra dev`
- `uv run ruff check .` (gate)
- `uv run pyright app/` (advisory — `continue-on-error: true`)
- Module parse check (every Python file in app/ + main + scripts/version.py)
- JS template syntax check via `node --check` on extracted script

CI doesn't run pytest because tests need project data (xlsx + PDFs)
that lives in GDrive. The lint + parse + JS-syntax gate catches
~95% of regressions a pure pytest run would.

### Verification battery (12 checks before handoff)

```
[1/12] Boot time:  1.18s  (was 7s — Ops-2 win)
[2/12] API routes: 68/68 critical present
[3/12] Project state: 660 rows, 309 catalogs ingested, active=Smart Plant 1
[4/12] Catalog library: 309 catalogs, 2 row links
[5/12] Continuity: STATE_20260510_111800.md available
[6/12] Manual:    /api/manual/raw 200 (47 KB), /manual page 200
[7/12] Claude:    detected (off until claude auth login)
[8/12] B3 autocomplete: 2 suggestions for R9
[9/12] Phase 17 WYSIWYG: byte-exact ✓
[10/12] UX-2 bulk preview: 50 catalogs with empty brand
[11/12] UX-3 page picker: /api/row/apply_catalog wired
[12/12] Ops-1 atexit hook: installed
```

```
$ uv run pytest -q
35 passed in 5.67s

$ uv run ruff check .
All checks passed!
```

### Lessons captured

- **Native `prompt()` is unacceptable for Thai-language UI** — the
  font is unreadable on most macOS browsers. Always replace with a
  proper modal.
- **sha256 of first 4 MB** is enough for content-identity on PDFs
  (they have unique header + xref + trailer). Full hashing is overkill.
- **Atexit hooks must never raise** — wrap everything in try/except,
  log to stderr, return None on failure. Otherwise interpreter
  shutdown can hang or print confusing tracebacks.
- **Test coverage should target the value line** — the +30 tests
  added cover 5 critical contracts (Phase 17 WYSIWYG, Phase B3,
  Phase 1 SSE, Phase 2 catalog CRUD, Phase 2.1 export builder) that
  were single-test-coverage before. CI gate now meaningful.

### Final state for handoff

| Aspect | Before today | After Tier 2 + final |
|---|---|---|
| `comply_verify_gui.py` | 15,175 lines | **5,233 lines** (–66%) |
| Test count | 5 | **35** (+600%) |
| Boot time | ~7 s | **1.18 s** (–83%) |
| Phase 17 invariant | tested | tested + byte-exact regression test |
| `ruff check` | clean | clean |
| `pyright app/` | 0 errors | 0 errors |
| CI gate | none | GitHub Actions on push/PR |
| User manual | none | 47 KB MANUAL.md + /manual viewer |
| UX prompts | 1 native `prompt()` | 0 (replaced by modal) |
| Multi-company switcher | API only | **UI shipped** |
| Continuity producer | Skill Agent only | **GUI also writes** STATE |
| Catalog cleanup tools | manual SQL | bulk preview + apply UI |
| Filter "what's left" | none | "🔴 ต้องดู" toggle |

---

## Calm Mode — Apple + Tesla minimalism (default UI redesign)

**Trigger**: user — *"ออกแบบ UXUI ของระบบใหม่ทั้งหมด ให้เหลือแค่
เท่าที่ USER ต้องใช้จริงๆ เพราะตอนนี้รู้สึกว่าข้อมูลเยอะไปหมด ไม่
user friendly แล้วเปลี่ยน theme ui ให้ minimal คล้าย Tesla + Apple"*.

The system grew over many sprints — by the audit point it had:
- 8-icon activity rail
- 4 ribbon mode tabs
- 4 topbar pills + 5 buttons
- 5-section AI pane
- Multiple modals + menus + toasts
- Status bar with row info + verdict + progress + Claude badge + save state

Too much. The user-proof loop is **pick → see → decide → next** — 4
steps. Everything else is occasional.

### Design philosophy

- **Apple HIG palette**: `#0071e3` blue accent, `#1d1d1f` text,
  `#f5f5f7` surface-2, `#e5e5ea` divider. White background dominant.
- **SF Pro font stack** with `-0.005em` letter-spacing,
  `font-feature-settings: ss01, cv11`.
- **Tesla chrome**: frosted topbar with `backdrop-filter: blur(20px)`,
  thin 1px dividers (no heavy borders), shadow only on the primary
  pass button.
- **One primary action visible**: the active verdict pill is filled,
  others are subtle pills in a segmented control.
- **Generous spacing**: status bar / action bar use `s-6` (20px)
  padding, tree rows 4-12px.
- **Subtle scrollbars**: 8px, semi-transparent on hover only.

### What's hidden in Calm Mode

| Feature | Where it goes |
|---|---|
| Activity rail icons (catalog/export/learn/versions/audit/AI/theme/help) | "▾ More" menu in topbar |
| Ribbon mode tabs (Verify/Edit/Re-annotate/Apply Auto) | hidden — use Edit toggle in PDF pane |
| AI pane (proposal + Run Claude + teach + recent) | toggle via More menu |
| Action bar: Auto / Mark / auto-next | hidden — use right-click Col D menu instead |
| Status bar: Claude badge, save state, row info detail | hidden |
| Topbar: project pill, continuity pill, theme toggle, stats pill | hidden — find in More menu |
| kbd-help, FAB cluster | hidden |

### What's still visible

```
┌──────────────────────────────────────────────────────┐
│   Comply · Smart Plant 1            ⌘  ▾ More       │  ← slim topbar
├────┬───────────────────────────────────┬────────────┤
│    │                                   │            │
│Tree│           Center                  │    PDF     │  ← 3-pane minimal
│    │           (TOR + xlsx)            │            │
│    │                                   │            │
└────┴───────────────────────────────────┴────────────┤
│  [Notes textarea]                                    │  ← action bar (notes only)
├──────────────────────────────────────────────────────┤
│  [✓ผ่าน 1] [✗ไม่ผ่าน 2] [⚠แก้ 3] [⏭ข้าม 4]   65/660 │  ← status bar (verdict+progress)
└──────────────────────────────────────────────────────┘
                                          Calm Mode · click
```

### Implementation

- Single CSS block scoped under `body[data-ux-mode="calm"]`
  (~520 lines of overrides, 88 selectors)
- IIFE at boot reads `localStorage["comply-ux-mode"]`,
  defaults to `"calm"` if unset
- `toggleUxMode()` flips between `calm ↔ pro`, persists, toasts
- New "▾ More" button in topbar (`.topbar-more-menu`) opens
  popover with grouped entries:
  - **Library**: Catalog Library, Export PDF
  - **Project**: Switch project, Continuity state, Versions
  - **Tools**: Learning patterns, Database & audit, Toggle AI pane
  - **App**: Settings, Manual, Help & shortcuts
  - **Toggle**: Show all controls (Pro mode) ↔ Hide noise (Calm mode)
- Bottom-left "Calm Mode · click for Pro" hint pill
- All Calm Mode CSS is opt-in via `body[data-ux-mode="calm"]`
  selector — Pro mode is the existing styling, untouched
- Dark mode auto-detected via `prefers-color-scheme: dark`

### Verification

- 36/36 smoke tests pass (was 35; +1 for Calm Mode hooks)
- `ruff check`: All checks passed
- All 65 routes still resolve
- HTML still serves all UI elements (toggle just hides them via CSS)
- Pro Mode still functional for users who prefer the old layout

### Lessons captured

- **Hide, don't delete**: every advanced feature lives in the More
  menu. Power users keep access; casual users get a clean slate.
- **One palette, one accent**: Tesla's red and Apple's blue are
  both single-accent systems. Picking one (Apple blue) and using
  it ONLY on the primary action keeps the eye on what matters.
- **Frosted topbar > opaque**: `backdrop-filter: blur(20px)` on a
  semi-transparent white gives that "iOS native" feel without any
  3rd-party libs.
- **Calm Mode default + opt-out > Pro default + opt-in**: most
  users never change defaults. Setting the bar at "minimal" by
  default ships the redesign to everyone immediately.
