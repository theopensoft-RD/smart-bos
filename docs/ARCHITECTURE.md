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
