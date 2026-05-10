# Architecture — Comply Verify Tool

> **Read this first** if you (or an AI agent) are picking up this project
> mid-flight. It captures the design decisions that aren't obvious from
> the code alone.

---

## 0. The product, in one sentence

A desktop-class web tool that lets a human do **visual proof-checking**
of AI-generated annotations on Comply spec spreadsheets + catalog PDFs,
with a continuous learning loop that turns every user correction into
patterns the AI re-uses next time.

---

## 1. The three contracts

The whole system follows three contracts. Every line of code must respect them.

### Contract A — "xlsx is the single source of truth"

`output/Comply spec Smart Plant 1.xlsx` is the **canonical** state of the
project. The DB, the in-memory caches, the FTS index, the audit log are all
**derived** from xlsx + filesystem. They can be rebuilt by deleting `_db/`
and rebooting.

**Implication**: writes go to xlsx **first** (with O_TRUNC workaround for
Google Drive's permission quirks), then mirror into DB. Never write to DB
without writing xlsx — divergence is the worst failure mode.

### Contract B — "AI proposes, user verifies, system learns"

Every change to xlsx flows through this triangle:

```
   Core Generator         User Action            Feedback Loop
   ─────────────         ─────────────          ─────────────
   rules + learned       click verdict /        record_feedback() →
   patterns + (LLM)  →   edit Col D /       →   retrain_patterns() →
   = suggestion +        Mark catalog           promote rules in
   confidence                                   learned_patterns
```

**Implication**: never auto-apply without recording feedback. Even silent
fallbacks ("ยินดีปฏิบัติ" because AI couldn't find content) are recorded
when the user later overrides them — that's the strongest training
signal.

### Contract C — "always work on the latest version"

`scripts/version.py` produces snapshots. On boot, the system checks
whether working files match the latest snapshot:

- **in_sync** → nothing to do
- **working_ahead** → auto-take a `boot-auto` snap (latest catches up)
- **working_behind / divergent** → BANNER, never auto-restore

**Implication**: destructive operations always pre-snapshot. Restore is
a user action, never automatic.

---

## 2. Module boundaries

```
comply-module/
├── comply_verify_gui.py    ← Flask app + UI (1 file, ~282 KB)
│                              The "monolith glue" — entry point, routes,
│                              HTML/CSS/JS template. Contains both the
│                              legacy module-level functions AND the
│                              integration with the OOP layer.
│
└── app/                    ← Application package
    ├── __init__.py
    ├── core.py             ← Domain model (Row, CatalogPDF, Project)
    ├── database.py         ← SQLite layer (12 tables + FTS5)
    └── learning.py         ← HITL pipeline (suggest/feedback/retrain/LLM)
```

### Why is `comply_verify_gui.py` so big?
Historical — it grew turn-by-turn during exploration. The HTML/CSS/JS
template alone is ~140 KB. **Don't refactor it speculatively.** Split
ONLY when adding a feature naturally lives elsewhere. Specifically, the
Flask routes are good candidates to move into `app/server/` someday but
nothing in there is "wrong" today.

### What lives where

| Concept | Module |
|---|---|
| Reading xlsx, parsing Col D, inferring sections | `comply_verify_gui.py` (legacy fns) |
| Row/CatalogPDF/Project classes | `app/core.py` |
| SQLite schema, sync helpers, FTS, audit | `app/database.py` |
| Learning feedback, pattern mining, LLM hook | `app/learning.py` |
| Flask routes, HTML/CSS/JS template | `comply_verify_gui.py` |

### Co-existence pattern (gradual OOP migration)

Legacy globals (`ROWS`, `PDF_INDEX`, `SECTION_INDEX`) still exist in
`comply_verify_gui.py`. After every `load_rows()` call, they're mirrored
into the OOP layer via `PROJECT_OBJ.set_rows([core.from_legacy_dict(r) ...])`.
**New code should prefer `PROJECT_OBJ.row(N)` over scanning ROWS**, but
old code paths keep working.

---

## 3. Data flows

### 3.1 Boot

```
boot()
  ├─ db.init_db(_db/comply.db)
  ├─ migrate verification_status.json → DB (one-time)
  ├─ build_pdf_index()       ← scan output/ → PDF_INDEX + SECTION_INDEX
  ├─ load_rows()             ← read xlsx → ROWS (4 inference passes)
  ├─ collect_extra_refs()    ← TOR/BOQ paths
  ├─ build_tor_section_index()  ← scan TOR for section markers
  ├─ index_tor_text()        ← per-page normalised text
  ├─ sync_db_from_memory()   ← mirror everything into SQLite
  └─ boot_sync_check()       ← contract C invariant
```

### 3.2 The 4 row-inference passes (ORDER MATTERS)

```python
# Pass 1: parse Col D into structured fields, resolve_ref_to_pdf
#   → "เอกสาร 5.1.2-2 …" parsed; pdf_rel set if direct match works.

# Pass 2: brand_model parents inherit pdf_rel from next sub-row OR from
#         _candidate_pdfs_for_section(row.section)
#   → "ยี่ห้อ Lenovo รุ่น …" rows get a catalog without a Col D ref.

# Pass 3: section inheritance for sub-items
#   PREFER MORE-SPECIFIC: between last_section ("5.1.1.2") and D-ref-derived
#   ("5.1.1"), keep the deeper one. Why: dot-form refs ("5.1.1-2") are
#   3-level but the actual section is 4-level — D-ref alone is too shallow.

# Pass 4: empty + commitment Col D → folder convention
#   _candidate_pdfs_for_section() handles dot/dash/wildcard/parent fallback.
#   Commitment rows get a catalog so the user can use 📍 Mark.
```

If you reorder these, things break in subtle ways. Document the new
order if you do.

### 3.3 Auto-annotate

```
preview = auto_annotate_plan(row_num)
  ├─ detect_row_role()                    ← header/item/sub_item + parent
  ├─ if section_header:
  │    ├─ apply_learned_brand(filename)   ← user-validated brand wins
  │    └─ else parse_brand_model_from_filename()
  ├─ if item/sub_item:
  │    ├─ find_text_match_in_pdf()        ← token-search Col B in PDF
  │    └─ compute Square + FreeText label rects
  ├─ confidence_score()                   ← 0.65 rules / 0.82 +pattern / 0.92 learned
  └─ provenance trail                     ← what fired the suggestion

apply_auto_annotate_plan(plan)
  ├─ pre-snap via version.py
  ├─ apply_pdf_edits()                    ← Square + FreeText into PDF
  ├─ write Col C/D in xlsx (O_TRUNC)
  ├─ refresh ROWS + sync_db_from_memory()
  ├─ db.log_audit(action='auto_annotate_apply')
  └─ learning.record_feedback(user_action='accepted'|'edited')
```

### 3.4 Manual-annotate (for "ยินดีปฏิบัติ" rows)

```
GUI: user clicks 📍 Mark on commitment row
  ├─ /api/manual_annotate/context returns {role, suggested_label, candidates}
  ├─ if multiple candidate PDFs → prompt user to pick
  ├─ force edit mode + drawRect tool
  └─ banner with target row info

  user draws a rect in catalog
  ├─ Square added by addRectAt
  ├─ auto-pair FreeText label nearby with the suggested_label content
  └─ user can drag/resize either annot

  user clicks ✓ Save & update Col D
  ├─ pre-snap
  ├─ apply_pdf_edits([Square, FreeText])
  ├─ Col D = make_col_d_for_row(row, role_info, pdf_path, page)
  ├─ write xlsx
  ├─ refresh + audit
  └─ record_feedback(user_action='edited', generator='commitment_fallback')
```

### 3.5 The "every 5 events → retrain" cadence

The frontend calls `tickRetrain()` after each:

- status verdict (1/2/3/4)
- inline Col D edit
- auto-annotate apply (handled server-side)
- manual annotate apply
- Col D revert from dropdown menu

Every 5 ticks → background fetch `/api/learn/retrain` → toast if patterns
were promoted. Threshold (`RETRAIN_THRESHOLD = 5`) is in the JS at the
top of the toast/auto-retrain block.

---

## 4. Section number conventions (the most error-prone area)

This caused **most** of the bugs we hit. Memorise these.

### 4.1 Folder conventions

```
output/5.1.1. ระบบ.../5.1.1.-2/5.1.1.2. เครื่อง...Lenovo.pdf
                    │  │              │
                    │  │              section number in filename (depth 4)
                    │  └─ "-N" suffix in folder name (parent rack convention)
                    └─ section number in folder name (depth 3)
```

### 4.2 Col D ref formats — TWO conventions co-exist

| Convention | Example | Generator |
|---|---|---|
| **Dash form** | `เอกสาร 5.1.2-2 ตู้ควบคุม… หน้า 1` | older agent + SKILL.md |
| **Dot form (full filename)** | `เอกสาร 5.1.1.2 เครื่องคอม...Lenovo... หน้า 1` | newer agent + clone_*.py |

**Both must resolve.** `_ref_to_folder_keys()` and `_candidate_pdfs_for_section()`
do bidirectional translation. Do **not** standardise on one without
migrating all existing Col D values.

### 4.3 Section depth tells you what to look up

| Depth | Section | Maps to |
|---|---|---|
| 1 | `5` | chapter root (TOR boundary) |
| 2 | `5.1` | top-level group |
| 3 | `5.1.2` | section with sub-catalogs (`5.1.2-1` through `-4`) |
| 4 | `5.1.1.2` | sub-section with ONE catalog (in `5.1.1.-2/`) |

For depth-4 sections, folder is `{parent}.-{N}/` and ref form `{parent}-{N}`.
For depth-3 sections with sub-catalogs, ref form is `{section}-{N}`.

When the candidate-finder wildcard-matches, **sort by trailing N
ascending** so the parent (`-1`) wins. (Was a bug — R8 picked `-4`.)

---

## 5. Highlight matching — also surprisingly subtle

### 5.1 Thai text gotchas

- **SARA AM** (`ำ` U+0E33) ↔ NIKHAHIT + SARA AA (`ํา`): TOR uses decomposed,
  xlsx uses precomposed. **NFD doesn't fix it** because U+0E33 is only a
  compatibility decomposition. We use NFKD plus an explicit replace.
- **Combining-mark order**: `น้ำ` (xlsx) vs `นํ้า` (TOR) — NIKHAHIT and
  MAI THO swap because NIKHAHIT has class 0. Manual reorder regex needed.
- → All text matching uses `normalize_thai_text()` first, never raw substring.

### 5.2 Section-bounded TOR search

When a phrase repeats across TOR sections (e.g. "Power Supply Redundant"
appears in Server, NGFW, AND switches), the matcher filters rects to
within the section's y-bounds on the page. The bounds are computed by
finding the next-sibling section header on the same page.

`_section_y_bounds_on_page()` is the function. **Don't skip the bounds
filter** — it's why R21 doesn't highlight NGFW item 11 anymore.

### 5.3 Anchor-on-rarest-token

When token-fallback fires (full phrase didn't match), generic words like
"Firewall" appear 3+ times → all get highlighted → scattered. Fix: pick
the **rarest** matched token as anchor, keep only rects on the same line
(±14pt y) as that anchor.

`_anchor_cluster_by_rarest()` is the function. Tested on R24 — was 7
hits scattered, now 3 on one line.

---

## 6. Database schema (12 tables, all in one .db file)

| Table | Owner | Purpose |
|---|---|---|
| `rows` + `rows_fts` | xlsx | mirrored Comply spec rows + FTS5 search |
| `pdfs` + `pdf_annotations` | filesystem | catalog files + Square/FreeText annots |
| `tor_sections` + `tor_pages` | TOR pdf | section→page index + normalised text |
| `tor_row_matches` | (cache) | per-row TOR lookup results |
| `verification_status` | user | canonical pass/fail/skip + notes |
| `snapshots` | filesystem | mirror of `_versions/snapshots/` |
| `pdf_history` | filesystem | mirror of `output/_pdf_history/` |
| `auto_annotate_plans` | system | every plan generated + apply result |
| `audit_log` | append-only | every meaningful change |
| `learning_feedback` | append-only | every (suggestion → user action) triple |
| `learned_patterns` | system | distilled rules with confidence + on/off |

**Safe to delete** the entire `_db/` directory — it'll rebuild on next
boot. The legacy `verification_status.json` is preserved as a backup.

---

## 7. Learning loop internals

### Pattern types currently mined

| Type | Trigger | Output | Example |
|---|---|---|---|
| `filename_brand` | Latin-token in filename stem | brand string | `ruijie → Ruijie` |
| `section_vendor` | section root (`5.1`, `5.2`) | vendor in Col E | `5.1 → SMART` |
| `row_format_d` | (role, section_root) tuple | Col D shape | `(section_header, 5.2) → brand_model` |
| `annot_position` | catalog folder name | label rel position | (stub — full impl pending) |

### Promotion threshold

Set to `2` in `comply_learning.PROMOTION_THRESHOLD`. With ≥2 agreeing
samples, a pattern goes from "noise" to "rule". Ties are broken by
sample count and recency.

### LLM provider hook

```python
import comply_learning as learning
def my_llm(prompt, context): ...
learning.set_llm_provider(my_llm, name="claude-3-5-sonnet")
```

Off by default. When set, only fires when rules + learned patterns yield
low confidence. The result is just another "generator" type in
`learning_feedback` — no special-casing in the rest of the pipeline.

---

## 8. Frontend architecture

### Layout (responsive)

```
>1280px:  3 columns  (Tree | Center | Catalog)
1280-900: 2 columns  (Tree | Center; Catalog folds with toggle)
900-700:  stacked    (Tree on top, Center below)
<700px:   tabs       (one pane visible at a time)
```

### Key state objects (in JS)

| Variable | What it holds |
|---|---|
| `DATA` | `/api/index` response — rows, sections, tree, status, sync, stats |
| `ROWS_BY_NUM` | quick-lookup row dict |
| `TREE_ROOT` | the section tree |
| `SELECTED_ROW` | currently inspected row |
| `EDIT_MODE`, `EDIT_ANNOTS`, `UNDO_STACK` | edit-mode state |
| `MANUAL_MODE`, `MANUAL_TARGET_ROW` | manual-annotate state |
| `_PENDING_RETRAIN_COUNT` | feedback counter for auto-retrain |

### Action bar (floating, bottom-center)

```
[R65 · 5.1.2 · Schneider QO116…] [SMART]    flags…
[✓ผ่าน 1] [✗ไม่ผ่าน 2] [⚠แก้ 3] [⏭ข้าม 4]   [✨Auto] [📍Mark]   ☑auto-next
[notes textarea…]                                                 [↺reset]
```

The Col D dropdown menu (single-click) is **separate** from the action
bar — it lives next to the Col D cell in xlsx preview.

---

## 9. Known constraints / limitations

1. **xlsx mutex**: only one process can write xlsx at a time. The Flask
   app holds an implicit lock during write. Don't open Excel while the
   app is running — both will write and one will lose.
2. **Google Drive sync**: file mtime/hash can change without content
   changes (Drive re-uploads). The version-sync logic uses sha256, which
   is correct, but expect occasional "working_ahead" prompts.
3. **PyMuPDF inline annotations**: some PDFs (เสา catalogs) embed
   annotations as inline dicts (xref=0). They render correctly but can't
   be edited via `apply_pdf_edits` — only viewed/highlighted. Migration
   to indirect annots is possible but not implemented.
4. **No multi-user**: single-user assumption. The DB has WAL mode but
   the app doesn't reconcile concurrent edits.
5. **TOR text-search ceiling**: ~98% accuracy on chapter 5.1; some
   commitment rows that genuinely don't appear in TOR will land on
   `section_start` with 0 hits.

---

## 10. Extension points

If you need to add:

| Feature | Where to plug in |
|---|---|
| New pattern type for the learner | Add a mining loop in `comply_learning.retrain_patterns()` |
| LLM-based suggestion | `learning.set_llm_provider(callable)` |
| New auto-annotate strategy | Wrap or replace `auto_annotate_plan()` |
| New PDF format support | `safe_iter_annots()` + `parse_inline_annots()` in `comply_verify_gui.py` |
| New verdict / status type | Update `verification_status` table + UI buttons |
| New audit action | Just call `db.log_audit(action='your_action', ...)` from any endpoint |
| New keyboard shortcut | `document.addEventListener('keydown', …)` block in JS |

---

## 11. Bootstrap a new agent / new contributor

If you (human or AI) just walked into this project:

1. Read `README.md` (5 min) — features + quick start
2. Read this file (`docs/ARCHITECTURE.md`) (15 min) — design rationale
3. Skim `docs/CHANGELOG.md` — what was tried + lessons
4. Run the launcher, click around, see how it feels (10 min)
5. Open the audit log (`📊 Audit`) to see recent activity
6. **Before changing anything**: `python3 scripts/version.py snap "before-<your-work>"`
7. Make changes incrementally, run end-to-end after each, verify in
   browser

If you're an AI and the user asks you to fix a bug, **ask which row
number / file / function** — don't rely on conversation memory across
context-window compaction.

---

## 12. Phase 13–17 additions (Acrobat-style UI + Claude integration)

The five most recent phases reshaped both the UI and the AI plumbing.
Read this section together with `docs/CHANGELOG.md` Phases 13–17 if
you're picking up after these changes.

### 12.1 Module layout — what's new

```
comply-module/
├── comply_verify_gui.py        ← still the monolith glue (now ~12K LOC)
├── .env / .env.example         ← API key + model + budget (gitignored)
└── app/
    ├── __init__.py
    ├── core.py
    ├── database.py             ← + `llm_calls` table (Phase 15)
    ├── learning.py             ← + Anthropic bridge via install_into_learning()
    └── anthropic_provider.py   ← NEW (Phase 15) — Claude wrapper
```

`anthropic_provider.py` is the **only** file that imports the Anthropic
SDK. Everything else talks to learning's existing `set_llm_provider()`
seam. This keeps the LLM swap-out point single-sourced.

### 12.2 Anthropic provider (`app/anthropic_provider.py`)

```
AnthropicProvider
  ├─ __init__(api_key, model, budget_usd_per_day)
  ├─ tools = [
  │     propose_col_d(ref, page, label),
  │     propose_brand_model(brand, model),
  │     escalate_to_user(question, options),
  │   ]
  ├─ system blocks (4 ephemeral cache_control entries):
  │     1. role + invariants
  │     2. SKILL.md (verbatim)
  │     3. project knowledge base summary
  │     4. last N learned_patterns
  ├─ propose(prompt, context, row_num) → ToolUseResult
  ├─ BudgetExceededError when day spend ≥ cap
  └─ records every call into `llm_calls` (db.log_llm_call)

get_provider()                 ← module-level singleton
install_into_learning()        ← bridges to learning.set_llm_provider
```

Lifecycle:

1. Boot reads `ANTHROPIC_API_KEY` from `.env` (or env). If present and
   non-empty, `install_into_learning()` wires it.
2. `learning.suggest_for_row()` calls Claude **only** when rules +
   learned patterns yield confidence < 0.85.
3. Each call:
   - prompt + system blocks sent
   - Anthropic returns tool_use blocks
   - we record `(input_tokens, output_tokens, cache_write, cache_read,
     cost_usd, elapsed_ms)` to `llm_calls`
   - return value funnels back through `learning_feedback` so user
     verdicts still produce training signal

The provider intentionally has **no fallback**: on `BudgetExceededError`
or network failure we return None and the caller falls back to rules.
Don't add silent retries — that masks both bugs and runaway spend.

### 12.3 New DB table — `llm_calls`

```sql
CREATE TABLE llm_calls (
    call_id INTEGER PRIMARY KEY AUTOINCREMENT,
    ts TIMESTAMP NOT NULL,
    row_num INTEGER,
    model TEXT,
    stop_reason TEXT,
    input_tokens INTEGER,
    output_tokens INTEGER,
    cache_write_tokens INTEGER,
    cache_read_tokens INTEGER,
    cost_usd REAL,
    elapsed_ms INTEGER,
    tool_calls_json TEXT,
    response_text TEXT,
    prompt_size_chars INTEGER
);
```

Used for:
- Daily budget enforcement (sum cost_usd WHERE date(ts) = today)
- AI panel "today's spend" badge
- Debugging tool-use behaviour after the fact (tool_calls_json)

Like all derived tables, it's safe to delete; rebuilds empty on next
boot.

### 12.4 Frontend layout — 5 columns × 6 rows grid (Phase 16)

The legacy "topbar + 3-pane content + floating action bar" was
replaced with a CSS Grid that gives every persistent surface a named
slot. This is the **Phase A** Acrobat-style frame.

```
                ┌────── grid-template-columns ──────┐
                rail   tree   center   pdf     ai
              ┌───────┬──────┬────────┬──────┬──────┐
   safetop   │       (--safe-top spacer, full-width) │   ← row 1
              ├───────┴──────┴────────┴──────┴──────┤
   topbar    │ logo · project · sync · settings     │   ← row 2 (z=10)
              ├──────────────────────────────────────┤
   ribbon    │ mode-tabs · tools (mode-aware)       │   ← row 3
              ├───────┬──────┬────────┬──────┬──────┤
   content   │ rail  │ tree │ row    │ pdf  │ ai   │   ← row 4 (1fr)
              │       │      │ inspect│ view │ pane │
              ├───────┴──────┴────────┴──────┴──────┤
   action    │ verdict / notes / auto / mark        │   ← row 5 (76px)
              ├──────────────────────────────────────┤
   status    │ row · pdf · sync · claude badge      │   ← row 6 (28px)
              └──────────────────────────────────────┘
```

CSS template:
```
grid-template-columns: var(--rail-w) var(--tree-w) 1fr var(--pdf-w) var(--ai-w);
grid-template-rows:    var(--safe-top) auto auto 1fr var(--action-bar-h) var(--status-bar-h);
grid-template-areas:
  "safetop safetop safetop safetop safetop"
  "topbar  topbar  topbar  topbar  topbar"
  "ribbon  ribbon  ribbon  ribbon  ribbon"
  "rail    tree    center  pdf     ai"
  "action  action  action  action  action"
  "status  status  status  status  status";
```

### 12.5 Mode tabs (Verify / Edit / Re-annotate / Apply Auto)

Mode is now an explicit ribbon control, not a scattered set of toggles.

```
setMode(name)
  ├─ updates --mode dataset attr on root
  ├─ for back-compat:
  │    'edit'        → EDIT_MODE = true
  │    'reannotate'  → enters wizard (calls openReannotateWizard)
  │    'auto-apply'  → triggers auto-annotate flow
  │    'verify'      → EDIT_MODE = false, exits wizards
  ├─ ribbon shows the matching tool group (ribbon-mode-* divs)
  └─ _syncRibbonState() updates buttons enabled/disabled
```

Critical rule: **mode is the source of truth for which tools render**.
Don't add new tools that don't belong to a mode. If a tool is
universally useful, put it in the topbar (settings, theme, layout) or
status bar (sync, badges).

### 12.6 AI pane (right column)

```
.ai-pane
  ├─ header (badge, refresh, close)
  ├─ proposal card    — `aiPaneRefresh()` populates from
  │                     /api/auto_annotate/preview for selected row
  ├─ inline actions   — Accept / Edit / Reject buttons trigger
  │                     aiAccept(), aiEdit(), aiReject()
  ├─ tags             — toggleAiTag() adds positive/negative training tags
  └─ teach-back       — aiTeachSend() POSTs free-form correction to
                        /api/learn/feedback with kind='teach'
```

`toggleAiPane()` collapses the column to 0 width via CSS variable
swap (`--ai-w: 0`); the grid template re-evaluates and the pdf column
reclaims space — no JS layout work needed.

### 12.7 `--safe-top` + embedded mode

The top of the viewport is regularly covered by:
- Claude Preview MCP toolbar (when in Claude Code preview iframe)
- Safari / browser toolbars on macOS

We ship a CSS variable `--safe-top` (default 0) that offsets the entire
grid. The first row is reserved as a transparent spacer of that
height.

Auto-detection:
```js
if (window.self !== window.top) {
  document.documentElement.style.setProperty('--safe-top', '56px');
}
```

User override: settings menu "Safe top offset" slider (0–120 px) writes
to `localStorage` and applies on boot.

Backup access: a fixed FAB at bottom-left exposes `Settings` and
`Layout` so the user is never locked out even if `--safe-top` is wrong.

### 12.8 Boot speedup

Old `_build_pdf_records_for_db()` called `list_pdf_annots()` for every
PDF (~101 of them × Google Drive on-demand fetch ≈ 60–90 s).

The DB's `pdf_annotations` table was **write-only** at boot — no read
path needed it warmed. Removed the per-PDF annot scan; annots fetched
live via `/api/pdf_meta?pdf=…` when the user opens a catalog.

Result: cold boot dropped from ~70 s to ~7 s on Google Drive.

### 12.9 WYSIWYG annotation render (Phase 17)

Edit mode used to render the page with `no_annots=True` and re-paint
all annots from JSON via SVG. This was systemically off because
PyMuPDF SVG ≠ PyMuPDF baked-PDF (font hinting, banner backgrounds,
border rounding all differ).

New render contract:
- **Always bake all existing annots into the page bytes**, even in
  edit mode. The base image is byte-identical to preview.
- The SVG overlay paints **only** annots with `_isNew=true` or
  `id === SELECTED_ANN_ID`. New ones aren't in the baked PDF yet;
  the selected one needs handles.
- Saving a new annot: `apply_pdf_edits` mints it into the PDF, then
  the page is re-rendered (now baked), and the SVG drops `_isNew`.

This is enforced in `comply_verify_gui.py` `render_page()` and the JS
overlay in `refreshOverlay()`.

### 12.10 Two filesystems, one git workflow

The user's working directory is on Google Drive (iCloud-style
on-demand sync). `.git/` operations there silently time out.

Workflow (already established, document it for future agents):

1. Real git work happens in `/tmp/comply-git-work/smart-bos/`
2. After commit + push, `cp` modified source files back to the Google
   Drive working directory so the running Flask app picks them up
3. `.env` is **only** in Google Drive (never copied to /tmp clone) —
   it's gitignored and would leak the API key

Don't try to make `.git/` work on Google Drive. macOS's iCloud
Files-On-Demand can't satisfy git's tight latency requirements.

### 12.10b Verdict in status bar (Phase A6)

The action bar lost its 4-button verdict segmented control. Verdict
is **state**, so it now lives in the status bar as `.sb-verdict`
pills. The action bar is now strictly *content actions* (Auto / Mark
/ auto-next / notes).

```
┌─ status bar ─────────────────────────────────────────────────────┐
│ R65 · 5.1.2 · summary │ [✓ผ่าน1][✗ไม่ผ่าน2][⚠แก้3][⏭ข้าม4][↺] │ progress │ Claude $ │ saved │
└──────────────────────────────────────────────────────────────────┘
```

`statusBarUpdate()` is the single source of truth — it sets
`aria-checked` on the matching pill and toggles
`body[data-row-selected="0|1"]` to grey out the pills when no row
is selected. The legacy `setStatus`/`renderActionBar` queries
against `.ab-btn.pass` are no-ops (guarded by `if (btn)`) and kept
in place so the existing wrapper chain stays intact.

### 12.10c FAB / kbd-help cleanup (Phase A7)

The activity rail covers Search / Theme / Settings / Help in normal
use. The FAB cluster (bottom-left) and kbd-help strip (bottom-right)
are now CSS-gated:

- `.floating-actions { display: none }` by default;
  `body[data-embedded="1"] .floating-actions { display: flex }` for
  iframe / Claude Preview where the topbar might be covered.
- `.kbd-help` collapses to a 38 px ⌨ badge that expands on hover
  to ~600 px. Hidden entirely in embedded mode.

CSS-only — the DOM stays in place so accessibility/screen-reader
exposure of those affordances is unchanged.

### 12.10d Col D autocomplete (Phase B3)

`GET /api/row/col_d/suggest?row=N&q=text` returns ≤6 ranked
suggestions when the user inline-edits a Col D cell:

| Kind | Source | Weight |
|---|---|---|
| `ai` | `auto_annotate_plan(row).proposed_d` (rules + patterns, no LLM) | 100 |
| `neighbor` | Col D from same-section-root rows where verdict ∈ {pass, need_fix} | 50 |
| `shape` | canonical templates (`เอกสาร {section} ... หน้า ?`, commitment) | 10 |

Frontend (`_colDAcOpen` etc.) mounts a single dropdown on `<body>`
so it escapes any clipping context (same lesson as Phase 13 topbar
menu). Debounced 250 ms input; Tab/Enter accepts, ArrowUp/Down
navigates. The panel closes when the editor blurs (with a 120 ms
delay so a click on a suggestion is processed first).

### 12.10i Module split (Tier 2 — 2026-05-10)

`comply_verify_gui.py` shrank from 15 K → 5.2 K lines via three
coordinated changes. The new structure:

```
~/Code/smart-bos/
├── comply_verify_gui.py            ← 5,219 lines (was 15,175)
│                                     boot, indexing, xlsx routes,
│                                     auto_annotate, claude_stream,
│                                     /api/row/apply_catalog
├── app/
│   ├── server/
│   │   ├── templates/
│   │   │   └── index.html           ← T2.1: 395 KB HTML/CSS/JS
│   │   └── static/                  ← reserved for future split
│   ├── routes/
│   │   ├── __init__.py              ← register_all(app, root, output)
│   │   ├── catalog_api.py           ← T2.2: /api/catalogs/*,
│   │   │                              /api/companies, /api/projects/*
│   │   ├── export_api.py            ← T2.2: /api/export/*
│   │   └── continuity_api.py        ← T2.2: /api/continuity
│   ├── catalog.py
│   ├── claude_code_provider.py
│   ├── core.py
│   ├── database.py
│   ├── export.py
│   └── learning.py
└── tests/
    └── test_smoke.py                ← 9 tests (was 5)
```

State plumbing for blueprints uses `app.config`:
```python
# in main file
from app.routes import register_all
register_all(app, root=ROOT, output_root=OUTPUT)
# now app.config["COMPLY_ROOT"] and ["COMPLY_OUTPUT"] are set

# in blueprint
from flask import current_app
output = current_app.config["COMPLY_OUTPUT"]
```

This avoids circular imports (the alternative was importing the gui
module from a blueprint, which fails because the gui module imports
the blueprints).

Pyright is now at `typeCheckingMode = "basic"` with several legacy
report categories softened to warnings — strict on new code, lenient
on historical. CI can gate on `pyright app/` returning 0 errors.

### 12.10h Print-ready PDF export (Phase 2.1)

`app/export.py` builds a deliverable PDF combining Cover + TOC +
Comply Spec sheet + grouped catalogs + optional Audit appendix.

The TOC-before-catalogs ordering requires a placeholder-and-replace
trick:

```
1. Insert N estimated TOC placeholder pages
2. Build all downstream content, accumulating (level, title, page1)
3. Render the real TOC into a temp doc (returns actual_pages)
4. delete_pages(placeholders); insert_pdf(tmp, start_at=placeholder_pos)
5. Apply Δ = (actual_pages - placeholder_pages) to every bookmark
   whose page > TOC region
6. doc.set_toc(bookmarks) → reader navigation tree
7. Per-page footer: project name + "Page N of M"
```

Annotations bake transparently because `fitz.insert_pdf` preserves
the source PDF's appearance streams — Phase 17's WYSIWYG layer
flows into the export with zero special-casing.

Endpoints: `/api/export/preview`, `/api/export/package`,
`/api/export/download`, `/api/export/list`.
Output dir: `_db/exports/` (gitignored, per-machine).

### 12.10g Multi-company / catalog library (Phase 2)

Adds an *additive* layer over the xlsx-canonical model: a catalog
library that's reusable across projects, with metadata + annotations
editable in the DB. The xlsx is still the single source of truth for
project state (Contract A), but catalogs are now first-class entities.

```
                ┌─────────────────────────────────┐
                │  Catalog Library (DB)           │
                │  ─────────────────────────────  │
                │  catalogs                       │
                │   ├─ pdf_rel, sha256, pages     │
                │   ├─ brand, model, category     │
                │   ├─ section_hint, description  │
                │   └─ metadata_json              │
                │                                 │
                │  catalog_annotations  (DB-side  │
                │                       templates)│
                │  catalog_pages        (FTS text)│
                └────────────┬────────────────────┘
                             │
                             ▼
            ┌────────────────┴────────────────┐
            │                                 │
   ┌────────┴─────────┐             ┌─────────┴────────┐
   │  Project A       │             │  Project B       │
   │  (Pattaya/SP1)   │             │  (Pattaya/SP2)   │
   │                  │             │                  │
   │  rows            │             │  rows            │
   │   └─ row_catalog │             │   └─ row_catalog │
   │      _links ─────┼─── SHARES ──┼──── _links       │
   └──────────────────┘             └──────────────────┘
```

**Key design decisions**:
- **Catalog == PDF + metadata**, identified by `pdf_rel` (relative
  to OUTPUT). De-dup key `pdf_sha256` for future "same file in
  different folder" detection.
- **Annotations in DB** going forward. For backwards compat the
  PDF-baked annotations are still rendered (Phase 17 WYSIWYG).
  Phase C will overlay DB annotations on top.
- **Bindings are project-scoped**: `row_catalog_links (project_id,
  row_num) → catalog_id`. Same row in different projects can use
  different catalogs.
- **Active project is implicit** in single-user UI — boot elects
  one, `get_active_project()` is the only scope check needed for
  MVP. UI multi-project switcher = Phase D.

**Apply flow** (`POST /api/row/apply_catalog`):
1. snapshot before mutation
2. synthesize Col D (or use user-provided text)
3. write Col D into xlsx (O_TRUNC dance for Google Drive)
4. `load_rows() → sync_db_from_memory()`
5. `bind_row_to_catalog(...)` records the link
6. `db.log_audit(action="apply_catalog", ...)`

**Module layout** (after Phase 2):
```
app/
├── catalog.py             ← NEW: companies/projects/catalogs/links
├── claude_code_provider.py
├── core.py
├── database.py            ← DB_VERSION=2 (added Phase 2 tables)
├── learning.py
└── anthropic_provider.py
```

### 12.10f Claude Code as core provider (Phase 1)

The module `app/claude_code_provider.py` is the new primary LLM
boundary. It uses the **Claude Agent SDK** (`claude-agent-sdk`
package, requires Python 3.10+) which spawns `claude` CLI as a
subprocess transport — same engine that powers Claude Code in IDE.

```
   Flask request /api/claude/stream?row=N
              │
              ▼
   ClaudeCodeProvider.propose_streaming(row_context)
              │
              ▼  (claude_agent_sdk.query)
   subprocess: claude --print --input-format stream-json …
              │
              │  uses ~/.claude.json OAuth token (Claude Max)
              │
              ▼
   AssistantMessage / UserMessage / ResultMessage events
              │
              ▼  (mapped to {type, …})
   yield → Flask Response stream → SSE
              │
              ▼
   Browser EventSource → AI pane chips
```

**Auth modes** (auto-detected by provider):
| Mode | Trigger | Cost |
|---|---|---|
| `claude_max` | `~/.claude.json` exists (after `claude auth login`) | Subscription |
| `api_key`    | `ANTHROPIC_API_KEY` set                              | Metered |
| `none`       | neither                                              | Provider unavailable |

**Provider toggles** via env:
- `COMPLY_LLM=claude_code` (default) → use new provider
- `COMPLY_LLM=anthropic` → use legacy `anthropic_provider.py` (API direct)
- `COMPLY_LLM=` (empty) → falls through `claude_code` first then `anthropic`

**Tools Claude can invoke**:
- `Read` (filesystem read — bounded to project cwd)
- `Grep` (regex search)
- `mcp__comply__propose_col_d`        — structured Col D proposal
- `mcp__comply__propose_brand_model`  — brand+model decomposition
- `mcp__comply__escalate_to_user`     — clarifying question

**Permission mode**: `default` (Claude asks before tool not in
`allowed_tools`). `Edit`/`Write`/`Bash` are NOT in the allowlist —
all xlsx/PDF mutations go through the GUI's user-confirm flow.

**Streaming is the default**: even for Phase 1, the legacy sync
`propose()` is kept for the `learning.set_llm_provider` bridge so
the existing low-confidence-refinement path still works. New UI
calls go through `/api/claude/stream`.

### 12.10e Patterns-triggered subsection (Phase B5)

`aiPaneRender()` reads `plan.provenance` (already shaped by
`apply_learned_brand` / `apply_learned_vendor` etc. into
`{kind: "learned", pattern_type, trigger, confidence, samples}`)
and renders an `.ai-patterns` subsection inside the Proposal section
listing which patterns fired and at what confidence. Pure-rules
proposals (no learned hits) skip the subsection entirely.

This is observability — it makes the AI's reasoning auditable so
users can debug a misfire by seeing which pattern caused it, then
disable that pattern from the rail's Learn panel.

### 12.11 Floating annotation toolbar (Phase A5)

Phase A5 adds a small toolbar that floats above the currently selected
annotation, Acrobat-style:

```
                  ┌──────────────────────────────────┐
                  │ Square · color · width · 🗑 · ⎘ │
                  └──────────────────────────────────┘
                  ┌─────────────┐
                  │ (the rect)  │
                  └─────────────┘
```

Implementation:
- DOM: single `#float-annot-toolbar` element appended to body
  (escapes any stacking context). Visibility tied to
  `SELECTED_ANN_ID`.
- Position: computed from the SVG group's `getBoundingClientRect()`,
  flipped below the annot if it would overflow the top.
- Reposition triggers: catalog scroll, page change, zoom, window
  resize, annot drag/resize. All call the same `repositionFloatingToolbar()`.
- Actions: `Delete` (existing `deleteAnnot`), `Duplicate` (clone with
  +12pt offset, mark `_isNew`).
- Properties: color (red/orange) + border width — write through to
  the annot dict, refresh overlay, no PDF edit until Save.

Hides automatically when:
- selection cleared (Esc, click empty space)
- mode switches away from `edit` or `reannotate`
- the SVG element scrolls out of view

---

## 13. Bootstrap (updated)

If you're an AI agent restarting after context compaction:

1. Re-read `MEMORY.md` and any user/feedback memories.
2. Skim `docs/CHANGELOG.md` Phase 13–17 entries for **what** changed.
3. Skim this file's section 12 for **why** + module map.
4. Re-orient on the user's last message before continuing.
5. Don't re-implement anything described as "done" unless explicitly
   asked — verify with `git log --oneline -20` from the /tmp clone
   first.
