# Knowledge Base — Smart Plant 1 Comply Spec Workflow

> KB สำหรับ agent + ผู้ทำงานในอนาคต — ข้อมูลที่ derive จาก code/PDF ไม่ได้ + บริบทที่จำเป็น

---

## 📂 KB Structure

| File | Purpose |
|---|---|
| `KB.md` (this file) | Human-readable overview + index |
| `catalogs.json` | 35 catalogs inventory: brand, model, file path, used_in sections |
| `sections.json` | 92 sub-sections + 530 rows status, Col D pattern distribution |
| `rect_coords.json` | 14 catalog rect coord templates (reusable for cloning) |
| `pipelines.md` | Workflow recipes (Opus + Sonnet Agent strategy) |
| `pitfalls.md` | Lessons learned + common issues + solutions |

---

## 🎯 Core Concepts

### Vendors (Consortium Members)
| Vendor | Scope | Notes |
|---|---|---|
| **TRIO** | Network/IT (Server, NGFW, L2/L3 Switch, NAS, Tablet, UPS, NVR, AP, CCTV, PC, Monitor) | Lead vendor for IT |
| **SMART** | UV Cabinet, sensors, cables, conduits, civil work, Fiber Optic equipment | Smart Solution own equipment |
| **SR** | HDPE pipe only (5 entries in 5.2.X.12 / 5.2.X.8) | SR Co Ltd partner |

### Comply Spec Structure (xlsx columns)
- **A** = Section ID (5, 5.1, 5.2 ...)
- **B** = TOR text (preserved verbatim — including typos)
- **C** = Proposed spec (catalog values, comparison words removed)
- **D** = Reference document/format (6 patterns)
- **E** = Vendor (SMART/TRIO/SR)
- **F** = Status (รอ user ตรวจสอบ / รอ catalog)

### Col D Patterns (6 + 2 sub-patterns)
1. **เทียบเท่าข้อกำหนด** — catalog spec matches TOR
2. **สูงกว่าข้อกำหนด** — catalog spec exceeds TOR ✨ (added 2026-05-09)
3. **ยี่ห้อ X รุ่น Y** — parent row with explicit brand+model
4. **ยี่ห้อ - รุ่น Y** — fabricate item (no brand, e.g., Vibration sensor)
5. **ยินดีปฏิบัติตามข้อกำหนด** — Software, install work, generic capability
6. **(empty)** — section header rows
7. *(extra)* **filename-format** — single-row item: `{section} {Col B desc minus จำนวน} {model}`
8. *(extra)* **model-only** — nested ข้อย่อย under multi-row install parent (e.g., `US-9106LSZH`)

---

## 📊 Project Status (2026-05-09)

```
Total comply rows:       660
Status รอ user ตรวจสอบ:  637 (97%) ✅
Section header rows:      23 (3%)
Status รอ catalog:         0 ✅
```

### Annotated PDFs in output/
- **101 catalog PDFs** annotated with header + brand/model + spec rects
- Spread across **5.1.1 — 5.1.8 + 5.2.1 — 5.2.6** (all sections)

---

## 🛠 Tools & Pipelines

### Scripts available
| Script | Purpose |
|---|---|
| `scripts/pdf_header.py` | Bulk add header to PDFs |
| `scripts/fix_uv_headers.py` | Fix UV cabinet headers |
| `scripts/version.py` | Snapshot-based version control (snap/list/diff/restore/prune) |

### Pipeline Strategy (Opus + Sonnet Agent)
For sister sections that share a catalog → use multi-model pipeline (40-50% cost savings):

```
[Opus 4.7] Inspect catalog + design rect coords + annotate REFERENCE file (1 ref/catalog)
[Sonnet]   Spawn Agent → clone reference annotations to all sister files + bulk xlsx update
```

**See:** `pipelines.md` for prompt templates + recipes

---

## 📋 Quick Reference

### Section vs catalog mapping (top 20)
| Section | Catalog | Vendor |
|---|---|---|
| 5.1.1.1, 5.2.1.1 | G3N-61142 (rack 42U) | SMART |
| 5.1.1.2, 5.2.1.3 | Lenovo SR630 V4 | TRIO |
| 5.1.1.3 | FortiGate 120G | TRIO |
| 5.1.1.4 | Ruijie RG-NBS5100 | TRIO |
| 5.1.1.5 | QNAP TS-433-4G | TRIO |
| 5.1.1.6 | Apple iPad A16 (HTML) | TRIO |
| 5.1.1.7 | Cleanline T-10K33LV2 | TRIO |
| 5.1.2.x — 5.1.6.x | UV cabinet + sensors + LED + เสา | SMART |
| 5.1.7 | SCADA Software | (commitment) |
| 5.1.8 | งานเดินสาย | (commitment + cables) |
| 5.2.1.4 | HP Pro Tower 280 G9 + P27 G5 Monitor | TRIO |
| 5.2.1.5 | H3C S6520X-16ST-SI Core Switch | TRIO |
| 5.2.1.6, 5.2.2.2, 5.2.3.2, 5.2.4.2 | Ruijie RG-CS85 L3 Switch | TRIO |
| 5.2.1.7 | Dahua NVR5432-EI | TRIO |
| 5.2.1.8 + 5 sisters | Ruijie CS4220 PoE L2 | TRIO |
| 5.2.1.9 + 3 sisters | LINK UF-2010A FO Drawer | SMART |
| 5.2.1.11 + 5 sisters | Dahua DH-IPC-HFW4231T CCTV | TRIO |
| 5.2.1.12 + 3 sisters | TP-Link EAP660 HD AP | TRIO |
| 5.2.1.13/14, 5.2.X.9-11 | LINK cables + Union EMT | SMART |
| 5.2.X.12, 5.2.X.8 | SR HDPE Conduit | SR |

**Full inventory:** `catalogs.json`

---

## 🔍 Lookup Queries

### "I need annotation coords for X catalog"
→ See `rect_coords.json[catalogs][catalog_key]`

### "What sections use catalog Y?"
→ See `catalogs.json[catalogs][Y].used_in`

### "What's the row range for section 5.X.Y.Z?"
→ See `sections.json[sections][5.X.Y.Z].subitems[].row`

### "How do I clone a catalog to sister sections?"
→ See `pipelines.md` → "Pipeline 2: Clone reference to sister sections"

### "What are common annotation pitfalls?"
→ See `pitfalls.md`

---

## 📝 Glossary

- **ข้อย่อย** (sub-item) — numbered list `1) 2) 3)` in TOR Col B
- **เทียบเท่า** — equivalent (catalog matches TOR exactly)
- **สูงกว่า** — exceeds (catalog spec better than TOR)
- **ยินดีปฏิบัติฯ** — full form `ยินดีปฏิบัติตามข้อกำหนด` = "willing to comply" — used for commitments without specific catalog spec
- **Sister section** — different sections that use the same catalog (e.g., CCTV used in 6 sections)
- **Reference file** — first annotated PDF for a catalog; used as template to clone into sisters
- **ICT TOR** — Thailand Information & Communication Technology government TOR template (catalogs with `(ICT)` suffix have pre-baked annotations)

---

## 🚦 Verification Checkpoints (before user proof)

1. **Status counts**: รอ user ตรวจสอบ = N, รอ catalog = 0
2. **Vendor coverage**: All rows with status have Col E
3. **Col C clean**: No comparison words (หรือดีกว่า, ไม่น้อยกว่า, ต้องสามารถ, จะต้อง)
4. **PDF headers**: Every page of every output PDF has red header
5. **PDF rects**: Brand+Model labels (ยี่ห้อ/รุ่น) present except for vibration sensors and pole fabrications
6. **Cross-ref**: Col D references resolve to actual PDF files

**Run verification script:** see `pipelines.md` → "Verification Checkpoint Script"

---

## 🆕 Adding a New Section/Catalog

**Steps for future agent:**

1. **Read TOR** — verify Col B is preserved verbatim (incl. typos)
2. **Identify catalog** — find file in `catalog/` matching the section's equipment
3. **Inspect catalog** — render with PyMuPDF, find brand logo + model + spec page positions
4. **Design rects** — record in `rect_coords.json` for reuse
5. **Annotate reference PDF**:
   - Header on every page
   - Brand rect + `ยี่ห้อ` label
   - Model rect + `รุ่น` label
   - Per ข้อย่อย: spec rect + `{section} ข้อย่อย N.` label
6. **For sister sections** — spawn Sonnet Agent with clone prompt (see `pipelines.md`)
7. **Update xlsx** — Col D parent + ข้อย่อย, Col E vendor (inherit to children), Col F = "รอ user ตรวจสอบ"
8. **Add to KB**:
   - `catalogs.json` — add catalog entry
   - `sections.json` — auto-regenerate by re-running script
   - `rect_coords.json` — add coord template
9. **Verify** + snapshot via `python3 scripts/version.py snap-full "<tag>"`

---

## 🔗 Related Files

- **`SKILL.md`** (root) — workflow rules + conventions (read first)
- **`output/Comply spec Smart Plant 1.xlsx`** — main work file
- **`TOR/`** — source Terms of Reference
- **`BOQ/`** — Bill of Quantities (verify quantities match)
- **`_versions/snapshots/`** — backup history
