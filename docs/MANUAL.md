# คู่มือการใช้งาน — Comply Verify Tool

> **สำหรับใคร**: User ที่จะตรวจ proof ของเอกสาร comply (Smart Plant 1, Plant 2, โครงการอื่นๆ ในอนาคต)
>
> **สถานะ**: เป็น HITL (Human In The Loop) tool — AI ช่วยเสนอ, user ตัดสิน, ระบบเรียนรู้จาก correction
>
> **เวอร์ชันคู่มือ**: 2026-05-10 · ตามระบบหลัง Phase C

---

## 📑 สารบัญ

- [บทที่ 1 — ภาพรวมระบบ](#บทที่-1--ภาพรวมระบบ)
- [บทที่ 2 — เริ่มต้นใช้งาน](#บทที่-2--เริ่มต้นใช้งาน)
- [บทที่ 3 — รู้จักหน้าจอ](#บทที่-3--รู้จักหน้าจอ)
- [บทที่ 4 — Workflow หลัก: User-proof loop](#บทที่-4--workflow-หลัก-user-proof-loop)
- [บทที่ 5 — แก้ไข Col D](#บทที่-5--แก้ไข-col-d)
- [บทที่ 6 — Annotation ใน catalog PDF](#บทที่-6--annotation-ใน-catalog-pdf)
- [บทที่ 7 — Catalog Library](#บทที่-7--catalog-library)
- [บทที่ 8 — AI Assistance (Claude)](#บทที่-8--ai-assistance-claude)
- [บทที่ 9 — Export PDF Package](#บทที่-9--export-pdf-package)
- [บทที่ 10 — Project & Continuity](#บทที่-10--project--continuity)
- [บทที่ 11 — Versions & Audit](#บทที่-11--versions--audit)
- [บทที่ 12 — Troubleshooting](#บทที่-12--troubleshooting)
- [ภาคผนวก A — Keyboard Shortcuts ทั้งหมด](#ภาคผนวก-a--keyboard-shortcuts-ทั้งหมด)
- [ภาคผนวก B — File Paths & Storage](#ภาคผนวก-b--file-paths--storage)
- [ภาคผนวก C — Glossary](#ภาคผนวก-c--glossary)

---

## บทที่ 1 — ภาพรวมระบบ

### 1.1 ระบบนี้ทำอะไร

Comply Verify Tool เป็นเครื่องมือ desktop-class web app ช่วย user ทำ **visual proof-checking** ของ comply spec spreadsheet (xlsx) เทียบกับ catalog PDFs ที่ AI annotate มาแล้ว — พร้อม HITL learning loop ที่ระบบฉลาดขึ้นทุกครั้งที่ user แก้ไข

```
   ┌─────────── AI Core ───────────┐
   │  rules + learned patterns     │  ─────►  Suggestion + Confidence
   │  + Claude Code (Agent SDK)    │
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
   ┌──────── Retrain Loop ──────────────┐
   │  every 5 events → mine patterns      │
   │  → toast: "🧠 Learned 2 patterns"    │
   └──────────────────────────────────────┘
```

### 1.2 Concepts สำคัญ

| คำ | ความหมาย |
|---|---|
| **Project** | โครงการหนึ่ง เช่น "Smart Plant 1" — มี xlsx + ชุด catalog PDFs |
| **Company** | บริษัทเจ้าของ project (เช่น Smart Solution) |
| **Comply spec** | Excel ไฟล์ที่มี 660 rows ของ TOR + Col C/D ที่ต้องเติม |
| **Row** | แถวหนึ่งใน xlsx — มีหมายเลข, Col A (section ID), Col B (TOR text), Col C, Col D (reference) |
| **Catalog** | PDF เอกสารผลิตภัณฑ์ของ vendor ที่ใช้ comply เช่น datasheet ของ Lenovo SR630 |
| **Annotation** | กรอบ (Square) + label (FreeText) บน catalog PDF ที่ชี้จุด spec ตรงกับ TOR |
| **Verdict** | ผลตัดสินของ user ต่อแต่ละ row: pass / fail / need_fix / skip |
| **HITL** | Human-In-The-Loop — AI เสนอ, user ตัดสิน, ระบบเรียน |
| **Snapshot** | ภาพถ่าย project ตอนหนึ่งๆ เก็บใน `_versions/snapshots/` (restore ได้) |
| **Continuity STATE** | เอกสาร handoff ระหว่าง session (`_continuity/STATE_*.md`) |

### 1.3 ส่วนประกอบของระบบ

```
~/Code/smart-bos/                       ← canonical (local SSD)
├── comply_verify_gui.py                 ← Flask app entry point
├── app/
│   ├── routes/                          ← Blueprints (catalog, export, continuity)
│   ├── server/templates/                ← HTML/CSS/JS (manual.html, index.html)
│   ├── catalog.py                       ← catalog DB layer
│   ├── claude_code_provider.py          ← Claude Max integration
│   ├── core.py / database.py / learning.py / export.py
├── _db/
│   ├── comply.db                        ← SQLite (rows, learned_patterns, audit, …)
│   └── exports/                         ← built PDF packages
├── _continuity/                         ← STATE handoff files
├── docs/                                ← MANUAL, CHANGELOG, ARCHITECTURE
└── (symlinks → GDrive)
    ├── output/                          ← xlsx + 309 catalog PDFs
    ├── _versions/                       ← snapshots
    ├── BOQ/, TOR/                       ← project docs
```

---

## บทที่ 2 — เริ่มต้นใช้งาน

### 2.1 Setup ครั้งแรก (one-time)

#### 2.1.1 ตรวจ dependencies

ระบบต้องการ:
- **Python 3.10+** (Phase 1 Claude SDK บังคับ)
- **uv** (Python package manager — เร็วกว่า pip 50×)
- **Node.js + npm** (สำหรับ Claude Code CLI)
- **Claude Code CLI** (สำหรับ Claude Max OAuth)

```bash
# macOS — ติดตั้งทั้งหมดใน 1 บรรทัด
brew install python@3.11 uv node
npm install -g @anthropic-ai/claude-code
```

#### 2.1.2 Authenticate Claude Max

```bash
claude auth login
# เปิด browser → login → กลับมา terminal
claude auth status   # ควรเห็น "loggedIn: true"
```

#### 2.1.3 ตรวจ output symlink

```bash
ls -la ~/Code/smart-bos/output
# ต้องเห็น: output -> /Users/.../GDrive/.../co-work/claude-code/output
```

ถ้า output ไม่ใช่ symlink หรือชี้ผิดที่ — ดูบทที่ 12.4

### 2.2 Launch (ทุกครั้งที่ใช้)

```bash
cd ~/Code/smart-bos
./start_verify_gui.command
```

หรือ double-click `start_verify_gui.command` ใน Finder

Browser เปิดอัตโนมัติที่ <http://127.0.0.1:5173>. ครั้งแรก uv sync ใช้เวลา ~75 วินาที (ติดตั้ง deps + Python). ครั้งถัดไป ~0.1 วินาที

### 2.3 ตรวจสถานะ

หลังเปิด browser:

| ส่วน | ดูที่ | สถานะที่ถูกต้อง |
|---|---|---|
| Boot log | terminal | ตอนท้าย: "version sync: state=in_sync" |
| Claude provider | Status bar (ล่างขวา) | "claude-sonnet-4-5 · Max · 0 calls" สีเขียว |
| Continuity badge | Topbar (ขวาบน, สีส้ม) | แสดง headline ของ STATE ล่าสุด |
| Catalog count | คลิก 📚 → stats pill | "309 catalogs · 1 co · 1 proj" |
| Tree | Pane ซ้าย | แสดง 660 rows ใต้ section tree |

---

## บทที่ 2.5 — Calm Mode vs Pro Mode

ระบบมี 2 โหมด UI — **Calm Mode** เป็น default (เห็นเฉพาะที่ใช้บ่อย),
**Pro Mode** เห็นเครื่องมือทั้งหมด

### Calm Mode (default)

ออกแบบสไตล์ Apple + Tesla — minimalism + พื้นที่หายใจ:

- หน้าจอเหลือ 4 ส่วน: Tree (ซ้าย) · Center (TOR + xlsx) · PDF (ขวา) · Status bar (ล่าง)
- Activity rail ซ่อนหมดยกเว้น Tree
- Ribbon mode tabs ซ่อน
- AI pane ซ่อน (เปิดได้จาก More menu)
- Action bar เหลือแค่ notes textarea
- Status bar เหลือ verdict pills + progress
- Topbar slim, frosted, blur background
- Apple blue `#0071e3` เป็น accent เดียว
- SF Pro typography
- Generous whitespace

### Pro Mode

เห็น UI เดิมทั้งหมด — ใช้เมื่อต้องการ:
- AI pane พร้อม proposal/teach-back/Run with Claude
- Activity rail ครบทุก icon
- Ribbon mode tabs
- Action bar Auto/Mark/auto-next
- Status bar Claude badge + saved indicator
- ทุก topbar pill (project switcher, continuity, theme toggle, ฯลฯ)

### สลับโหมด

3 ทาง:

1. **More menu** (▾ ใน topbar) → "Show all controls (Pro mode)" / "Hide noise (Calm mode)"
2. **Calm Mode hint** (มุมล่างซ้าย, สีเทาจางๆ) → คลิกเพื่อสลับ
3. **Console**: `localStorage.setItem('comply-ux-mode', 'pro')` แล้ว reload

### เมื่อไหร่ใช้โหมดไหน

| Workflow | Mode แนะนำ |
|---|---|
| User-proof rapid (J/K + 1-4) | **Calm** |
| ตรวจ catalog + แก้ annotations | **Calm** (Catalog เปิดผ่าน More menu) |
| ดู AI live stream (Run with Claude) | **Pro** (เห็น AI pane เต็มจอ) |
| Teach Claude, ดู learned patterns | **Pro** |
| Re-annotate wizard + advanced edit | **Pro** (เห็น ribbon mode tabs) |
| Multi-company switching | **Calm** (Switch project ผ่าน More menu) |
| Export PDF + tools occasional | **Calm** (More menu) |

**ความแตกต่างใน "More menu"**: Calm Mode ใช้ More menu (▾) ในมุม topbar เพื่อเข้าถึงฟีเจอร์ที่ซ่อนไว้.
Pro Mode ใช้ Activity rail (icon ฝั่งซ้าย) แทน

---

## บทที่ 3 — รู้จักหน้าจอ

ระบบใช้ **Acrobat-style layout** แบ่งเป็น 6 surface หลัก:

```
┌────────────────────────────────────────────────────────────┐
│  Topbar                                  [Continuity 🔔]  │  ← row 1
├────────────────────────────────────────────────────────────┤
│  Ribbon (mode tabs: Verify / Edit / Re-annotate / Apply)  │  ← row 2
├──────┬───────┬─────────────────────┬────────┬─────────────┤
│      │       │                     │        │             │
│  📂  │ Tree  │  Center             │  PDF   │  AI pane    │
│  🔍  │       │  (TOR + xlsx)       │  view  │  (proposal, │
│  🧠  │       │                     │        │   teach)    │
│  📦  │       │                     │        │             │
│  📊  │       │                     │        │             │
│  📚  │       │                     │        │             │
│  📄  │       │                     │        │             │
│  ✨  │       │                     │        │             │
├──────┴───────┴─────────────────────┴────────┴─────────────┤
│  Action bar (notes / Auto / Mark / auto-next)              │  ← row 5
├────────────────────────────────────────────────────────────┤
│  Status bar [R65 · 5.1.2 · summary][✓1 ✗2 ⚠3 ⏭4][stats]  │  ← row 6
└────────────────────────────────────────────────────────────┘
```

### 3.1 Activity Rail (ซ้ายสุด — แนวตั้ง)

| Icon | ฟังก์ชัน | คีย์ลัด |
|---|---|---|
| 📂 | Row tree (default) | — |
| 🔍 | Search (Command palette) | ⌘K |
| 📚 | Catalog Library | — |
| 📄 | Export print-ready PDF | — |
| 🧠 | Learning patterns | — |
| 📦 | Project versions / snapshots | — |
| 📊 | Database & audit | — |
| ✨ | Toggle AI pane | — |
| ☀️/🌙 | Toggle theme | — |
| ⚙️ | Settings | — |
| ❓ | Help & onboarding | ? |

### 3.2 Ribbon (Mode tabs)

| Mode | ใช้เมื่อไหร่ |
|---|---|
| **Verify** (default) | ตรวจ proof — pass / fail / need_fix / skip |
| **Edit** | แก้ annotations บน PDF (drawRect, addText, save) |
| **Re-annotate** | Wizard 3-step สำหรับเปลี่ยน rect+label ใน brand_model row |
| **Apply Auto** | Apply auto-annotate plans แบบ batch |

### 3.3 Tree (ซ้าย)

แสดง 660 rows ใน tree ตาม section. แต่ละ row มี:
- ตัวเลข row (เช่น R65)
- Status indicator: เขียว (pass) / แดง (fail) / เหลือง (need_fix) / เทา (skip / unverified)
- Section + brand/model เป็นข้อความสั้น

คลิก row → เลือก row นั้น → Center / PDF / AI panes อัปเดตให้

### 3.4 Center (TOR + xlsx)

- บน: TOR text excerpt (highlight ตรงที่ matched)
- ล่าง: xlsx Comply spec rows (Col A/B/C/D) แบบตาราง

**Col D คลิกได้**:
- คลิกซ้าย — เลือก row
- คลิกขวา — context menu (Re-annotate / Auto / Mark / Edit / Revert)
- ดับเบิลคลิก — inline edit (พร้อม **Phase B3 autocomplete**)

### 3.5 PDF Pane (ขวา)

แสดง catalog PDF ของ row ปัจจุบัน. มี:
- Page navigation (← → หรือ `[` `]`)
- Highlight toggle (เปิด/ปิด search highlight)
- Edit toggle (เข้า/ออก edit mode)
- Edit toolbar เมื่อ edit mode เปิด: drawRect / addText / select / undo / redo / save

### 3.6 AI Pane (ขวาสุด — เปิด/ปิดได้)

| Section | เนื้อหา |
|---|---|
| **Proposal** | AI เสนอ Col D + confidence + rationale |
| **Patterns triggered** (B5) | learned_patterns ที่ fire สำหรับ row นี้ |
| **Run with Claude Code** (Phase 1) | คลิก "Run" → Claude Code agent runs live, stream events |
| **Teach Claude** (B4) | tags + textarea ส่ง correction ให้ระบบเรียน |
| **Recent (last 5)** | สถิติ feedbacks + accuracy |

ปิด/เปิด AI pane ได้ที่ rail ✨ icon

### 3.7 Action Bar (ล่างกลาง)

ปุ่ม secondary actions:
- **Auto** — preview auto-annotate plan
- **Mark** — เริ่ม manual annotate (ลาก rect ใน catalog)
- **auto-next** toggle — หลัง verdict ไป row ถัดไปอัตโนมัติ
- **notes** textarea — บันทึกของ row

### 3.8 Status Bar (ล่างสุด)

| ส่วน | แสดง |
|---|---|
| Row info | `R65 · 5.1.2 · summary` |
| **Verdict pills** (Phase A6) | `✓ผ่าน 1` `✗ไม่ผ่าน 2` `⚠แก้ 3` `⏭ข้าม 4` `↺` (reset) |
| Progress | `done/total · progress bar` |
| Claude badge | `claude-sonnet-4-5 · Max · N calls` (เขียวถ้า online) |
| Save state | `● saved` |

---

## บทที่ 4 — Workflow หลัก: User-proof loop

### 4.1 ทำงานปกติทุก row

```
Pick row (J / K หรือคลิก) ───►  ดู TOR + xlsx + PDF
                                     │
                                     ▼
                          AI proposal (มีใน AI pane)
                                     │
                            ┌────────┴────────┐
                            │                 │
                       Col D ตรง?         Col D ผิด/ขาด?
                            │                 │
                            ▼                 ▼
                    [✓ผ่าน 1]            แก้ไขก่อน
                            │                 │
                            │       ┌─────────┼─────────┐
                            │       │         │         │
                            │       ▼         ▼         ▼
                            │   Inline    Mark      Re-annotate
                            │    edit    in PDF      wizard
                            │       │         │         │
                            │       ▼         ▼         ▼
                            │   [⚠แก้ 3]  [✗ไม่ผ่าน 2]  [↺ reset]
                            │       │         │         │
                            └───────┴────┬────┴─────────┘
                                         │
                                         ▼
                                Auto-advance to next row
                                  (toast: "Learned 2 patterns")
```

### 4.2 Keyboard shortcuts (ใช้บ่อยที่สุด)

| Key | Action |
|---|---|
| `J` / `K` | row ก่อนหน้า / ถัดไป |
| `N` | row uncertain ถัดไป (low-confidence) |
| `1` / `2` / `3` / `4` | verdict pass / fail / need_fix / skip |
| `[` / `]` | catalog page ก่อนหน้า / ถัดไป |
| `,` / `.` | TOR page ก่อนหน้า / ถัดไป |
| `⌘K` | Command palette |
| `?` | Help & shortcuts |

ดู shortcuts ทั้งหมดที่ [ภาคผนวก A](#ภาคผนวก-a--keyboard-shortcuts-ทั้งหมด)

### 4.3 ตรวจ row อย่างเร็ว (rapid verify)

1. เปิด `auto-next` ในยุทธ์ action bar
2. ตั้ง row แรกที่ต้องการเริ่ม (คลิกหรือ J/K)
3. กด `1` (pass) ทุกครั้งที่ proof ตรง — ระบบเลื่อนไป row ถัดไปอัตโนมัติ
4. เจอที่ผิด — กด `3` (need_fix) → แก้ → กด `1` (pass)

### 4.4 Notes per row

- Action bar มี notes textarea ที่ auto-save ทุก 400ms
- Notes ติดกับ row, อยู่ใน DB `verification_status` table

---

## บทที่ 5 — แก้ไข Col D

### 5.1 ทาง 1: Inline edit (ดับเบิลคลิก)

1. ดับเบิลคลิกที่ Col D ของ row ที่ต้องการแก้
2. **Autocomplete dropdown** เปิดใต้ cell (Phase B3) แสดง:
   - 🤖 **AI proposal** (จาก auto_annotate_plan)
   - 👥 **Neighbor** (Col D ของ row อื่นใน section เดียวกันที่ verify แล้ว)
   - 📐 **Shape** (template เช่น `เอกสาร 5.1.1 ... หน้า ?`)
3. พิมพ์เพื่อ filter, **Tab** หรือ **Enter** เพื่อ accept ตัวที่ highlight
4. **Esc** ออกโดยไม่ save

### 5.2 ทาง 2: Right-click context menu

1. คลิกขวาที่ Col D
2. เมนูแสดงตัวเลือกตามสถานะ row:

| สถานะ | ตัวเลือก |
|---|---|
| Commitment ("ยินดีปฏิบัติ") | Mark / Auto / Re-annotate / Edit |
| มี reference อยู่แล้ว | Re-annotate / Auto / Mark / Edit / Revert |
| Empty / needs_col_d | Auto / Mark / Edit |

### 5.3 ทาง 3: ใช้ AI (Run with Claude Code)

1. ในมี AI pane เลือก row → คลิก **Run** ใน "Run with Claude Code"
2. Claude อ่าน SKILL.md / KB / row context → เสนอ Col D + rationale
3. ดู streaming chips (thinking → tool_use → text → result)
4. คลิก **Accept** เพื่อ apply, **Reject** เพื่อบันทึกเป็น signal
5. Reject → "Teach" textarea จะเปิด — พิมพ์เหตุผลแล้ว Send

### 5.4 กฎสำคัญที่ต้องรู้ (จาก SKILL.md)

| Pattern | ใช้เมื่อ |
|---|---|
| `เอกสาร 5.1.2-2 ... หน้า N` | Catalog ref dash form (มีใน old data) |
| `เอกสาร 5.1.1.2 ... หน้า N` | Catalog ref dot form (preferred) |
| `ยี่ห้อ Lenovo รุ่น M70q-Tiny` | Brand_model (section_header rows) |
| `ยินดีปฏิบัติตามข้อกำหนด` | **Installation / software/firmware เขียนเอง / commitment** |
| `ไม่พบใน catalog` | **Hardware spec ที่หายังไม่เจอใน catalog — flag เตือน user** |
| (empty) | section header / parent items |

⚠ **กฎใหม่**: distinguish "ยินดีปฏิบัติ" vs "ไม่พบใน catalog" ให้ชัด:
- ติดตั้ง / สาย / ทดสอบ → ยินดีปฏิบัติ
- Software เขียนเอง → ยินดีปฏิบัติ
- Hardware spec ที่ควรมีใน catalog แต่ยังหาไม่เจอ → ไม่พบใน catalog (ตรวจซ้ำ)

---

## บทที่ 6 — Annotation ใน catalog PDF

### 6.1 Edit mode

เข้า edit mode ทำได้ 2 ทาง:
- คลิก **Edit toggle** ใน PDF pane toolbar
- คลิกที่ Mode tab "Edit" ใน ribbon

PDF เปลี่ยนเป็น **WYSIWYG mode** (Phase 17):
- รูปที่เห็น = รูปที่จะ save (byte-exact)
- annotations เก่าทั้งหมดยัง bake-in อยู่ — แก้ได้
- annotations ใหม่ที่ยัง unsaved จะมีกรอบแดงพิเศษ (`is-new`)

### 6.2 Drawing tools

| Tool | คีย์ | ใช้ |
|---|---|---|
| Select (V) | `V` | เลือก / drag / resize annotations |
| Draw rect (R) | `R` | ลาก rect → spec callout |
| Add text (T) | `T` | คลิก → label `5.1.1.2 ข้อย่อย 3` |

### 6.3 Floating annotation toolbar (Phase A5)

เมื่อ select annotation จะมี toolbar เด้งขึ้นเหนือ rect:

```
  ┌────────────────────────────────────────────┐
  │ Square · 🔴 · 1pt red · 📋 Duplicate · 🗑 │
  └────────────────────────────────────────────┘
              ↓
   ┌──────────────┐
   │ (annotation) │
   └──────────────┘
```

| ปุ่ม | คีย์ | ทำอะไร |
|---|---|---|
| **Duplicate** | `D` | clone +12pt offset (มาร์ค `_isNew`) |
| **Delete** | `Del` / `Backspace` | ลบ |

### 6.4 Save

- คลิก **Save** ใน edit toolbar — เขียนลง PDF file ตรง
- Auto-snapshot ก่อน save — restore ได้
- ผ่าน `apply_pdf_edits()` รักษา appearance streams ของ PyMuPDF

### 6.5 Undo / Redo

- `⌘Z` / `Ctrl+Z` — undo
- `⌘⇧Z` / `Ctrl+Shift+Z` — redo
- Stack เก็บ 200 step ล่าสุด

---

## บทที่ 7 — Catalog Library

### 7.1 เปิด Catalog Browser

คลิก 📚 ใน rail (icon ที่ 4)

```
┌──────────────────────────────────────────────────────────┐
│  Catalog Library  [309 catalogs · 1 co · 1 proj · 0 bound]│
├──────────────────────────────────────────────────────────┤
│  [search box]  [section ▾]  [Re-scan]                    │
├──────────────┬───────────────────────────────────────────┤
│  List (left) │  Detail (right)                           │
│              │                                           │
│  5.1.1 Lenovo│  [Section pill] filename                  │
│  M70q-Tiny   │                                           │
│  p.3 ...     │  Brand: [Lenovo]                          │
│              │  Model: [ThinkSystem SR630]               │
│  5.1.1 Apple │  Category: [Server]                       │
│  iPad A16    │  Section: [5.1.1.2]                       │
│  ...         │  Description: [...]                       │
│              │                                           │
│              │  [Save metadata] [✏ Edit annotations]     │
│              │  [Apply to R65] [Open PDF (raw)]          │
│              │                                           │
│              │  Annotations (DB-stored, per page): 0     │
│              │  Used by (0 rows)                         │
└──────────────┴───────────────────────────────────────────┘
```

### 7.2 Browse / search / filter

- **Search box**: filter ตาม brand / model / section / filename
- **Section dropdown**: เลือก section root (เช่น 5.1.1, 5.2)
- **Re-scan** button: scan output/ ใหม่ — ใช้เมื่อเพิ่ม PDF ใหม่ๆ

### 7.3 แก้ Metadata

1. คลิก catalog ทางซ้าย
2. แก้ Brand / Model / Category / Section / Description ในฟอร์มทางขวา
3. คลิก **Save metadata**
4. List ทางซ้าย refresh — สะท้อน metadata ใหม่ + sort ใหม่

### 7.4 Edit annotations (Phase C — ฟีเจอร์ใหม่)

1. ในรายละเอียด catalog → คลิก **✏ Edit annotations**
2. Modal ปิดอัตโนมัติ
3. PDF โหลดเข้า main viewer + edit mode เปิด
4. Banner สีส้มขึ้นบน: "Catalog edit mode · {filename} · ✕"
5. ใช้ tools เดิมทั้งหมด (drawRect / addText / floating toolbar / undo)
6. กด **Save** — annotations เซฟลง PDF file โดยตรง
7. Export package ครั้งถัดไป — เห็น annotations ใหม่ทันที
8. คลิก **✕** บน banner เพื่อออก catalog edit mode

### 7.5 Apply catalog → row

1. เลือก row ใน tree ก่อน (สำคัญ)
2. เปิด Catalog Browser → คลิก catalog
3. คลิก **Apply to R{N}** (เปลี่ยนสีตาม row ที่เลือก)
4. Prompt: "หน้าใน catalog?" — กรอก page (ว่างได้)
5. ระบบ:
   - Snapshot ก่อนเขียน
   - Sythesize Col D string (ตาม convention)
   - เขียนลง xlsx
   - บันทึก link ใน `row_catalog_links`
   - Audit log
6. Toast: "✓ Applied — R65 → เอกสาร 5.1.1.2 ... หน้า 3"

### 7.6 Annotations บน catalog ที่ใช้ใน DB

ส่วน "Annotations (DB-stored)" แสดง annotations ที่อยู่ใน `catalog_annotations` table

> **หมายเหตุ Phase C**: ตอนนี้เมื่อแก้ผ่าน "Edit annotations" จะเขียนลง PDF file ตรง — DB table ยังว่าง. Render layer สำหรับ DB annotations เป็น future work

### 7.7 Used by (rows ที่ใช้ catalog นี้)

แสดง project_id + row_num + col_d_text สำหรับทุก row ที่ apply catalog นี้แล้ว
- คลิก row number → กระโดดไป row นั้นในตาราง

---

## บทที่ 8 — AI Assistance (Claude)

### 8.1 ติดตั้ง Claude (one-time)

ทำตามบทที่ 2.1.2

ระบบจะ detect auto:
- ถ้า login OAuth → ใช้ **Claude Max** (ไม่จำกัด — ใช้ subscription)
- ถ้ามี `ANTHROPIC_API_KEY` ใน `.env` → ใช้ **API key** (metered, $5/day cap)
- ถ้าไม่มี → ใช้ rules-only mode (ไม่มี LLM)

### 8.2 Run with Claude Code (Phase 1 — แนะนำ)

ใน AI pane:

1. เลือก row จาก tree
2. Section "Run with Claude Code" → คลิก **Run**
3. ระบบ stream events live:
   - 🟦 **Thinking**: Claude คิดอะไร
   - 🟧 **Tool use**: Read PDF / Grep KB / call propose tool
   - 🟩 **Tool result**: ผลของ tool
   - 🟪 **Text**: narration
4. Result card ขึ้นล่างสุด:
   ```
   ✓ propose_col_d                       8.3s · $0.0000
     เอกสาร 5.1.1.2 อุปกรณ์... หน้า 3
     conf: 92%
     rationale: ...
     [Accept] [Reject]
   ```
5. **Accept** → เขียน Col D ลง xlsx + bump retrain counter
6. **Reject** → record เป็น learning_feedback signal

### 8.3 Auto-annotate (rule-based, no LLM)

1. คลิก **Auto** ในย action bar
2. Modal preview แสดง:
   - proposed_d (จาก rule-based generator)
   - confidence (0-1)
   - generator (rules / rules+pattern / claude+rules)
   - annotations จะเพิ่ม
3. คลิก **Apply** เพื่อ commit

### 8.4 Manual annotate (📍 Mark)

ใช้เมื่อ rule-based + AI หาไม่เจอ — user ต้อง locate spec เอง

1. เลือก row → คลิก **📍 Mark** ในย action bar
2. ระบบ:
   - Force enter edit mode + drawRect tool
   - Banner ขึ้นบอก row ที่ target
3. ลาก rect ใน catalog เพื่อระบุตำแหน่ง spec
4. Auto-pair FreeText label ขึ้นใกล้ๆ (เช่น `5.1.1.2 ข้อย่อย 3.`)
5. คลิก **✓ Save & update Col D**
6. ระบบเขียน Col D + annotations + record feedback

### 8.5 Re-annotate Wizard (Phase 14)

สำหรับ brand_model rows ที่ต้องเปลี่ยน rect+label

1. คลิก Mode tab "Re-annotate" หรือ right-click Col D → Re-annotate
2. Wizard 3 steps:
   - Step 1: Pick PDF (ถ้ามีหลาย candidate)
   - Step 2: Draw rect + label ใหม่
   - Step 3: Confirm + save

### 8.6 Teach Claude (Phase B4)

ใต้ AI pane Proposal:

1. คลิก tags ที่เกี่ยวข้อง (#wrong-page / #brand-wrong / #missing-spec / #typo / #format / #commitment)
2. พิมพ์เหตุผลใน textarea
3. คลิก **Send to Claude**
4. ระบบ record เป็น `learning_feedback` row
5. หลัง 5 events → auto-retrain → toast "🧠 Learned 2 patterns"

### 8.7 Patterns triggered (Phase B5)

ใน AI pane Proposal section จะมีกล่อง "Patterns triggered":

```
🧠 Patterns triggered                               2
  filename_brand    ruijie         95% · 12 samples
  section_vendor    5.1            87% · 8 samples
```

แสดงให้รู้ว่า AI ใช้ pattern ไหนเสนอ → audit ได้ → ถ้า pattern ผิด ปิดได้ใน Learn panel

---

## บทที่ 9 — Export PDF Package

### 9.1 ทำไมต้อง export

หลัง user-proof ครบ — ต้องส่ง deliverable เป็น PDF ที่:
- รวม comply spec sheet + ทุก catalog
- มี cover page + TOC + bookmarks
- Page numbers + footer
- Print-ready (A4 portrait)

### 9.2 เปิด Export modal

คลิก 📄 ใน rail (icon ที่ 5)

### 9.3 ตัวเลือก

| Option | Default | คำอธิบาย |
|---|---|---|
| **Mode** | Full package | Cover + TOC + Comply + ทุก catalog |
| | Comply spec sheet only | แค่ Comply PDF |
| | Catalogs only | catalogs + cover + TOC (ไม่มี comply sheet) |
| **Section filter** | (empty = all) | ระบุ section root เช่น `5.1.1` |
| **Bound only** | off | เฉพาะ catalogs ที่ apply ลง row แล้ว |
| **Include audit log** | off | append audit_log appendix |

### 9.4 Live preview

แต่ละครั้งที่เปลี่ยน option, preview pane ที่ล่างอัปเดต (debounced 200ms):
```
Project:        Smart Plant 1 (Smart Solution)
Comply sheet:   ✓ Comply spec Smart Plant 1.pdf
Catalogs:       309 catalogs · 5.1.1: 27 · 5.1.2: 12 · ...
Audit:          200 entries
```

### 9.5 Build

1. คลิก **Build PDF** → spinner
2. ระบบใช้เวลา ~0.3s ต่อ section, ~3s ต่อ full package
3. PDF เซฟใน `_db/exports/`
4. Download trigger อัตโนมัติ
5. Toast: "✓ Export ready — 311 pages · 491 KB"

### 9.6 Recent exports

ใน modal มี collapsible **Recent exports** — list ของ PDFs ที่ build ก่อนหน้า, คลิกดาวน์โหลดซ้ำได้

### 9.7 ตรวจ PDF

หลัง download — เปิดด้วย Preview / Acrobat:
- **Bookmarks panel**: ตรวจ tree (Cover / TOC / Comply / Catalogs / sections / individual catalogs)
- **TOC page**: คลิก page number → กระโดดไปหน้า
- **Footer**: ทุกหน้ามี "Project Name | Page N of M"

---

## บทที่ 10 — Project & Continuity

### 10.1 Multi-company / multi-project

ระบบรองรับหลายบริษัท / หลาย project — แต่ตอนนี้ใช้ implicit "active project" คือ Smart Plant 1

API ที่มี (UI switcher อยู่ใน roadmap):
```
GET  /api/companies
POST /api/companies                      {name, code}
GET  /api/projects[?company_id=]
POST /api/projects                       {company_id, name, code, ...}
POST /api/projects/<id>/activate
```

### 10.2 Continuity STATE document

```
_continuity/
└── STATE_20260510_111800.md           ← handoff document
```

#### ใครเขียน

- **Skill Agent** ใน session อื่นเขียนตอนจบ session
- ตอนนี้ GUI ไม่ได้เขียนเอง (future work)

#### มีอะไรบ้าง

```markdown
# Continuity State — 2026-05-10 11:18 UTC

## Last completed task
- ทำอะไรเสร็จล่าสุด

## Open in-progress
- มีอะไรค้างอยู่

## Pending user decisions
- ที่ถาม user แต่ยังไม่ตอบ

## Critical recent context
- User corrections (verbatim)
- Bugs ที่เจอ
- Path conventions

## Next planned action
- step ถัดไปที่จะทำ
```

#### ใช้ที่ไหน

1. **Topbar badge** (ส้ม): แสดง headline ของ STATE ล่าสุด
2. คลิก badge → modal เปิด แสดง markdown เต็ม + history
3. **Claude provider**: โหลดเข้า system prompt ก่อนทุก call → Claude รู้ context

### 10.3 SKILL.md / KB / pitfalls / sr_pattern

ระบบโหลดอัตโนมัติเข้า Claude prompt:
- **SKILL.md** (75 KB) — กฎ + convention + 12 rules
- **KB.md** — knowledge base
- **pitfalls.md** — bugs ที่เคยเจอ
- **sr_pattern.md** — SR annotation conventions
- **pipelines.md** — workflow pipelines
- **Top 30 learned_patterns** — pattern ที่ระบบเรียนได้

แก้ไฟล์เหล่านี้ → cache 60s → Claude เห็นค่าใหม่

---

## บทที่ 11 — Versions & Audit

### 11.1 Snapshots

ทุก destructive op จะ snapshot ก่อน:
- `auto_annotate_apply` → `pre-auto-annotate-row-N`
- `manual_annotate_save` → `pre-manual-row-N`
- `apply_catalog` → `pre-apply-catalog-row-N`
- `export_package` → no snap (read-only)

#### เปิด Versions panel

คลิก 📦 ใน rail

#### ดู snapshots

แสดงทุก snapshot ใน `_versions/snapshots/`:
```
2026-05-10_122125_before-co-work-output-relink
2026-05-10_115048_Update-SKILL-KB-pitfalls-sr_pattern---add-continui
2026-05-10_111357_Add--ไม-พบใน-catalog--rule...
...
```

#### Restore

1. คลิก snapshot
2. Confirm dialog
3. ระบบ restore xlsx + PDFs + annotations
4. Reload page

### 11.2 Audit log

คลิก 📊 ใน rail — แสดง:

```
Stats:
  rows_total: 660 · rows_with_pdf: 597
  pdfs: 309 · annotations: 0
  snapshots: 44 · audit_entries: 27
  ...
```

```
Recent audit:
  2026-05-10 06:17 export_package        project 1
  2026-05-10 06:17 catalog_update        catalog 102
  ...
```

#### Search audit

มี search bar — query string → FTS5 over `audit_log` table

---

## บทที่ 12 — Troubleshooting

### 12.1 Boot ไม่ขึ้น / port 5173 in use

```bash
lsof -ti tcp:5173 | xargs kill -9
./start_verify_gui.command
```

### 12.2 "Not logged in" ใน Claude Code stream

```bash
claude auth login
claude auth status   # ตรวจ "loggedIn: true"
# รีสตาร์ท GUI
```

### 12.3 Output PDFs ไม่ขึ้น

ตรวจ symlink:
```bash
ls -la ~/Code/smart-bos/output
# ต้องเป็น symlink → /GDrive/.../co-work/.../output
```

ถ้า symlink เสีย:
```bash
ln -sfn "/Users/.../GDrive/.../co-work/claude-code/output" ~/Code/smart-bos/output
```

### 12.4 GDrive iCloud sync timeout

อาการ: `git status` ค้าง / `cp` timeout

✅ ป้องกัน: code อยู่ใน `~/Code/` เสมอ, ไม่อยู่ใน GDrive โดยตรง
ถ้าเจอ — ตรวจว่า `.git/` ไม่อยู่ใน GDrive folder

### 12.5 Snapshot เก่าไม่เคย restore — ทำ?

```bash
cd ~/Code/smart-bos
.venv/bin/python scripts/version.py list
.venv/bin/python scripts/version.py restore <snapshot-id>
```

หรือใน UI: Versions panel → คลิก snapshot → Restore

### 12.6 DB corrupt / ต้อง rebuild

```bash
mv ~/Code/smart-bos/_db/comply.db ~/Code/smart-bos/_db/comply.db.bak
./start_verify_gui.command
# DB rebuild auto จาก xlsx + filesystem
```

### 12.7 Catalog metadata เพี้ยน (brand เป็น "PoE", "G7-05002")

ใน Catalog Browser:
1. ค้น brand ที่อยากแก้
2. คลิก catalog → แก้ Brand / Model fields → Save
3. ทำซ้ำจน clean

หรือใช้ SQL ตรงๆ:
```bash
sqlite3 ~/Code/smart-bos/_db/comply.db
> UPDATE catalogs SET brand='Lenovo' WHERE brand='-';
```

### 12.8 Tests แตก

```bash
cd ~/Code/smart-bos
uv run pytest -v 2>&1 | grep FAIL
```

ดูว่า test ไหน fail → อ่าน output → ดู `tests/test_smoke.py`

### 12.9 Logs

| ที่ไหน | มีอะไร |
|---|---|
| Terminal stdout (ตอนรัน) | boot log + Flask access log + errors |
| `_db/comply.db` audit_log | every meaningful change |
| `_db/comply.db` llm_calls | every Claude call (tokens, cost, response) |
| `_db/comply.db` learning_feedback | every (suggestion → user action) tuple |
| `_versions/snapshots/*/` | full state snapshots |

---

## ภาคผนวก A — Keyboard Shortcuts ทั้งหมด

### Navigation

| Key | Action | Context |
|---|---|---|
| `J` | row ก่อนหน้า | tree focused |
| `K` | row ถัดไป | tree focused |
| `N` | row uncertain ถัดไป | (low-confidence) |
| `[` | catalog page ก่อนหน้า | PDF pane |
| `]` | catalog page ถัดไป | PDF pane |
| `,` | TOR page ก่อนหน้า | TOR pane |
| `.` | TOR page ถัดไป | TOR pane |

### Verdict (Phase A6)

| Key | Action |
|---|---|
| `1` | ✓ pass |
| `2` | ✗ fail |
| `3` | ⚠ need_fix |
| `4` | ⏭ skip |

### Edit mode

| Key | Action |
|---|---|
| `V` | Select tool |
| `R` | Draw rectangle (Square) |
| `T` | Add text (FreeText) |
| `D` | Duplicate selected (Phase A5) |
| `Del` / `Backspace` | Delete selected |
| `↑` / `↓` / `←` / `→` | Nudge selected (1pt; +Shift = 5pt) |
| `Esc` | Exit edit mode (or close text editor) |

### General

| Key | Action |
|---|---|
| `⌘K` / `Ctrl+K` | Command palette (search rows / actions) |
| `⌘Z` / `Ctrl+Z` | Undo |
| `⌘⇧Z` / `Ctrl+Shift+Z` | Redo |
| `?` | Help & shortcuts |
| `Esc` | Close modal / dropdown / wizard |

### Col D inline edit (Phase B3)

| Key | Action | Context |
|---|---|---|
| `↑` / `↓` | Navigate suggestions | autocomplete open |
| `Tab` | Accept highlighted suggestion | autocomplete open |
| `Enter` | Accept (if suggestion highlighted) | autocomplete open |
| `Esc` | Close autocomplete (keep editing) | autocomplete open |
| `Esc` | Cancel edit (no save) | no autocomplete |

---

## ภาคผนวก B — File Paths & Storage

```
~/Code/smart-bos/
├── comply_verify_gui.py                 ← Flask entry (~5,200 lines)
├── start_verify_gui.command             ← launcher script
├── pyproject.toml + uv.lock             ← deps
├── .venv/                                ← Python 3.11 + deps (gitignored)
│
├── app/
│   ├── server/
│   │   ├── templates/
│   │   │   ├── index.html               ← UI template (~9,500 lines)
│   │   │   └── manual.html              ← this manual rendered
│   │   └── static/                       ← reserved for asset split
│   ├── routes/
│   │   ├── catalog_api.py               ← /api/catalogs/*
│   │   ├── export_api.py                ← /api/export/*
│   │   └── continuity_api.py            ← /api/continuity
│   ├── catalog.py                        ← catalog DB layer
│   ├── claude_code_provider.py          ← Claude Max integration
│   ├── core.py / database.py / learning.py / export.py
│   └── anthropic_provider.py            ← API-key fallback
│
├── _db/
│   ├── comply.db                         ← SQLite (20 tables)
│   └── exports/                          ← built PDF packages
│
├── _continuity/                          ← STATE_*.md handoff docs
│
├── docs/
│   ├── MANUAL.md                         ← THIS FILE
│   ├── CHANGELOG.md
│   └── ARCHITECTURE.md
│
├── tests/
│   ├── conftest.py
│   └── test_smoke.py                     ← 11 smoke tests
│
├── SKILL.md                              ← project domain knowledge (75 KB)
├── knowledge_base/
│   ├── KB.md
│   ├── pitfalls.md
│   ├── sr_pattern.md
│   ├── pipelines.md
│   ├── catalogs.json
│   ├── rect_coords.json
│   └── sections.json
│
└── (symlinks to GDrive)
    ├── output/                           ← xlsx + 309 catalog PDFs
    ├── _versions/                        ← 44 snapshots
    ├── BOQ/                              ← project xlsx
    └── TOR/                              ← project docs
```

### DB tables (sqlite `_db/comply.db`)

| Table | Owner | Purpose |
|---|---|---|
| `rows` + `rows_fts` | xlsx | mirrored Comply rows + FTS5 |
| `pdfs` + `pdf_annotations` | filesystem | catalog PDFs + annot index |
| `tor_sections` + `tor_pages` | TOR | section→page index + text |
| `tor_row_matches` | cache | per-row TOR match results |
| `verification_status` | user | pass/fail/skip + notes |
| `snapshots` | filesystem | mirror of `_versions/snapshots/` |
| `pdf_history` | filesystem | mirror of `output/_pdf_history/` |
| `auto_annotate_plans` | system | dry-run + apply records |
| `audit_log` | append-only | every meaningful change |
| `learning_feedback` | append-only | every (suggestion → user action) |
| `learned_patterns` | system | distilled rules with confidence |
| `llm_calls` | append-only | every Claude call (tokens + cost) |
| `catalogs` | catalog ingest | Phase 2 catalog metadata |
| `catalog_pages` | catalog ingest | per-page text excerpt (FTS) |
| `catalog_annotations` | UI / future | DB-stored annotations |
| `companies` + `projects` | UI | multi-company schema |
| `row_catalog_links` | UI Apply | row → catalog binding |
| `schema_version` | system | DB version (currently 2) |

---

## ภาคผนวก C — Glossary

| คำ | ภาษาไทย | ความหมาย |
|---|---|---|
| Annotation | คำอธิบายภาพ | กรอบ + label บน PDF page ที่ชี้จุด spec |
| Brand_model | ยี่ห้อ + รุ่น | Col D pattern สำหรับ section_header rows |
| Catalog | แค็ตตาล็อก | PDF เอกสารผลิตภัณฑ์ของ vendor |
| Col B | คอลัมน์ B | TOR text ใน xlsx |
| Col C | คอลัมน์ C | Comply text — สิ่งที่บริษัทเสนอ |
| Col D | คอลัมน์ D | Reference document — ที่ comply ไปอ้างอิง |
| Comply | คอมพลาย | การตรวจ spec ผลิตภัณฑ์เทียบกับ TOR |
| Confidence | ความเชื่อมั่น | คะแนน 0-1 ของ AI ต่อ proposal |
| Continuity | ต่อเนื่อง | การส่งต่อ context ระหว่าง session |
| Edit mode | โหมดแก้ไข | สถานะที่ user แก้ annotations ได้ |
| FreeText | ข้อความ | annotation ประเภท text label |
| FTS5 | Full-Text Search 5 | SQLite full-text search engine |
| HITL | Human In The Loop | mode ที่ AI ทำงานร่วมกับ user |
| Provenance | ที่มา | tracker ว่า output มาจาก rule/pattern/LLM ตัวไหน |
| Re-annotate | ทำ annotation ใหม่ | wizard 3-step สำหรับ rect+label |
| Row_catalog_links | binding | ตารางที่บอกว่า row ไหนใช้ catalog ไหน |
| Section | section number | ตัวเลขเช่น 5.1.1.2 ใน TOR |
| Section_header | หัว section | row ที่มีแค่ Col A กับ section number |
| Snapshot | snapshot | full state backup ใน `_versions/` |
| Square | สี่เหลี่ยม | annotation ประเภท rectangle |
| TOR | Terms of Reference | เอกสารกำหนดสเปก |
| Verdict | คำตัดสิน | pass / fail / need_fix / skip |
| WYSIWYG | What You See Is What You Get | edit mode = preview byte-exact |
| xref | cross-reference | PyMuPDF object reference |

---

## ภาคผนวก D — สำหรับ developer

### Quick links

- **Architecture**: [docs/ARCHITECTURE.md](./ARCHITECTURE.md)
- **Changelog**: [docs/CHANGELOG.md](./CHANGELOG.md)
- **Tests**: `tests/test_smoke.py`
- **Routes audit**:
  ```bash
  cd ~/Code/smart-bos
  .venv/bin/python -c "import comply_verify_gui as g; g.boot(); print('\n'.join(sorted(str(r) for r in g.app.url_map.iter_rules())))"
  ```

### Dev commands

```bash
uv sync                # refresh deps
uv run pytest -q       # smoke tests (11 tests, ~5s)
uv run ruff check .    # linter
uv run pyright app/    # type checker (basic mode)
```

### Adding a new feature

1. Snap state: `python scripts/version.py snap "before-<feature>"`
2. Code change
3. `uv run pytest -q && uv run ruff check .`
4. Add a smoke test if endpoint added
5. Commit + push

### Backup before risky op

```bash
cd ~/Code/smart-bos
python scripts/version.py snap "before-<description>"
# work
# if bad → python scripts/version.py restore <id>
```

---

> **คู่มือนี้อัปเดตล่าสุด**: 2026-05-10 (Phase C)
> **GitHub**: <https://github.com/theopensoft-RD/smart-bos>
> **License**: Proprietary (Smart Solution Co., Ltd.)
