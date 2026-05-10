# Comply Verify Tool — Smart Plant 1

GUI สำหรับช่วย user ตรวจ visual proof ระหว่าง **Comply spec (xlsx) ↔ TOR
↔ Catalog PDFs** ที่ AI annotate มาแล้ว พร้อม HITL learning loop ที่
ระบบฉลาดขึ้นทุกครั้งที่ user แก้ไข

```
┌─────────── AI Core ───────────┐
│  rules + learned patterns     │  ─────►  Suggestion + Confidence
│  + (optional LLM)             │
└──────────────┬────────────────┘
               │
               ▼
┌──────── User (Visual Proof) ─────────┐
│  ✓ Pass  ✗ Fail  ⚠ Fix  ⏭ Skip      │
│  📍 Mark in catalog                  │
│  Inline Col D edit                   │
└──────────────┬───────────────────────┘
               │ feedback
               ▼
┌────────── Retrain Loop ──────────────┐
│  every 5 events → mine patterns      │
│  → toast: "🧠 Learned 2 patterns"    │
└──────────────────────────────────────┘
```

## Quick start

```bash
# Double-click in Finder, or:
./start_verify_gui.command
```

Browser opens at <http://127.0.0.1:5173>. First run uses **uv** to install
Python 3.10+ + all deps from `pyproject.toml` (~75 s); subsequent runs
boot in ~0.1 s.

## Repo layout

```
~/Code/smart-bos/                      ← canonical (local SSD, fast .git)
├── app/, scripts/, docs/              ← code + docs
├── _db/                               ← per-machine state (SQLite, audit, llm_calls)
├── tests/                             ← pytest smoke suite
├── pyproject.toml + uv.lock           ← uv-managed deps
├── .venv/                             ← uv-created virtualenv (gitignored)
└── (symlinks to GDrive)
    ├── output      → project xlsx + 101 PDFs (canonical, shared)
    ├── _versions   → 43 snapshots (1.4 GB — kept in GDrive for backup)
    ├── BOQ, TOR    → project docs
```

The launcher resolves these symlinks transparently. Project data stays in
GDrive (so other team members can see it); only the dev machine's `.git/`
+ `.venv/` + `_db/` are local.

## Dev workflow

```bash
uv run pytest          # smoke tests (~1.3 s)
uv run ruff check .    # linter (warnings)
uv run pyright app/    # type checker (advisory)
uv sync                # refresh deps from pyproject.toml
```

## Documentation

- **[docs/ARCHITECTURE.md](docs/ARCHITECTURE.md)** — design contracts, module
  boundaries, the 4 inference passes, Thai-text gotchas, DB schema, learning
  loop internals, bootstrap guide for new agents.
- **[docs/CHANGELOG.md](docs/CHANGELOG.md)** — phase-by-phase evolution with
  bug post-mortems (R71 / R21 / R24 / R8 / R19), things considered and
  rejected, and open questions.
- **[SKILL.md](SKILL.md)** — domain knowledge for the verification workflow.

## Project layout

```
comply-module/
├── README.md                  ← you are here
├── SKILL.md                   ← domain knowledge for the workflow
├── start_verify_gui.command   ← macOS launcher (double-click)
├── comply_verify_gui.py       ← Flask app + UI (entry point)
│
├── docs/                      ← deep-dive documentation
│   ├── ARCHITECTURE.md        ← design decisions, contracts, data flows
│   └── CHANGELOG.md           ← phase-by-phase evolution + bug post-mortems
│
├── app/                       ← application package
│   ├── __init__.py
│   ├── core.py                ← OOP domain model (Row, CatalogPDF, Project)
│   ├── database.py            ← SQLite store (rows + audit + FTS5 + learning)
│   └── learning.py            ← HITL loop (suggest / record / retrain / LLM hook)
│
├── BOQ/                       ← input: Bill of Quantities (xlsx)
├── TOR/                       ← input: Terms of Reference (pdf + docx)
├── output/                    ← working files
│   ├── Comply spec Smart Plant 1.xlsx     ← THE spreadsheet (single source of truth)
│   ├── 5.1.x/.../*.pdf                    ← catalog PDFs (annotated)
│   ├── _archive/                          ← old versions
│   ├── _pdf_history/                      ← per-PDF edit snapshots
│   └── verification_status.json           ← legacy backup of pass/fail
│
├── scripts/                   ← CLI utilities
│   ├── version.py             ← snapshot tool (snap / restore / diff / prune)
│   ├── pdf_header.py          ← add page headers to catalogs
│   ├── fix_uv_headers.py      ← one-off header fix
│   └── clone_*.py             ← clone annotated PDFs to sister sections
│
├── knowledge_base/            ← reference docs
│   ├── KB.md
│   ├── pipelines.md
│   ├── pitfalls.md
│   ├── catalogs.json
│   ├── rect_coords.json
│   └── sections.json
│
├── _db/                       ← generated SQLite (rebuild-able, safe to delete)
│   └── comply.db              ← rows + annotations + status + audit + FTS5 + patterns
│
└── _versions/                 ← project-level snapshots
    └── snapshots/<YYYY-MM-DD_HHMMSS_tag>/
        ├── manifest.json
        ├── Comply spec Smart Plant 1.xlsx
        ├── SKILL.md
        └── output.tar.gz       (only for `snap-full`)
```

## Features (current)

### Three-pane responsive UI
- **Left** — folder tree of all 660 rows by section (filterable + searchable)
- **Center top** — TOR PDF preview (auto-scrolls to highlighted phrase)
- **Center bottom** — comply.xlsx context table (±6 rows around selected, double-click Col D to edit inline)
- **Right** — catalog PDF with annotation overlay (edit mode = drag rects + edit labels)
- **Floating action bar** — verdict (1/2/3/4) + ✨ Auto-annotate + 📍 Mark + auto-next
- **Mobile**: tabs collapse to single-pane view at <700px

### HITL workflows
| Workflow | Trigger | Outcome |
|---|---|---|
| **Visual verify** | click row | see TOR + xlsx + catalog highlighted |
| **Verdict** | `1`/`2`/`3`/`4` keys | recorded in DB + audit log |
| **Auto-annotate** | ✨ Auto button | AI proposes Col C/D + PDF annotations with confidence |
| **Manual mark** | 📍 Mark button | when AI fell back to "ยินดีปฏิบัติ" — draw rect → auto Col D |
| **Inline edit** | double-click Col D | quick edit + records as feedback |
| **Smart nav** | `N` key | jump to next uncertain row (low confidence + flags + has PDF) |

### Continuous learning
- Every user action → `learning_feedback` (append-only)
- Every 5 events → background `retrain_patterns()` → distil into `learned_patterns`
- Toast notifications when new patterns are learned
- 🧠 Learn modal shows accuracy + pattern list with on/off toggle
- LLM provider hook (off by default — pluggable Anthropic/OpenAI/Ollama)

### Versioning
- `📚 Versions` modal wraps `scripts/version.py`
- Auto-snap on boot if working files differ from latest snapshot
- "Always-load-latest" invariant — surfaces divergence as a banner
- Per-PDF edit history in `output/_pdf_history/`

## Database

`_db/comply.db` — SQLite, ~3 MB.  Tables:

| Table | Purpose |
|---|---|
| `rows` + `rows_fts` | mirrored xlsx + FTS5 search |
| `pdfs` + `pdf_annotations` | catalog files + every Square/FreeText |
| `tor_sections` + `tor_pages` | TOR section index + per-page normalized text |
| `verification_status` | canonical pass/fail/skip + notes |
| `snapshots` + `pdf_history` | mirror of `_versions/` and `output/_pdf_history/` |
| `audit_log` | append-only log of every change (status / pdf / restore / etc.) |
| `auto_annotate_plans` | every plan generated + applied result |
| **`learning_feedback`** | every (suggestion → user action) triple |
| **`learned_patterns`** | distilled rules from feedback (filename_brand / section_vendor / row_format) |

Safe to delete `_db/` — it rebuilds from xlsx + filesystem on next boot.

## Keyboard shortcuts

| Key | Action |
|---|---|
| `J` / `K` or ↓ / ↑ | next / prev row |
| `N` | next **uncertain** row (smart nav) |
| `1` / `2` / `3` / `4` | verdict pass / fail / fix / skip |
| `[` / `]` | catalog PDF page |
| `,` / `.` | TOR PDF page |
| `+` / `-` | zoom |
| `⌘S` | save edits in PDF edit mode |
| `⌘Z` / `⇧⌘Z` | undo / redo (edit mode) |

## Troubleshooting

- **"PDF ไม่ขึ้น"** — check `_db/comply.db` is recent; force refresh via
  `GET /api/refresh`.
- **Rows ไม่มี catalog** — open 📊 Audit → check pattern resolution; commitment
  rows now show fallback catalog so 📍 Mark works.
- **Import errors** — make sure you boot from the project root so
  `from app import ...` resolves.
- **Catalog edits don't stick** — `output/_pdf_history/<flat>/<ts>.pdf`
  has snapshots; restore via the History modal.
