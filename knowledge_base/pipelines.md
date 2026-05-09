# Pipelines — Workflow Recipes

> Copy-paste-ready recipes for common comply spec tasks. ใช้ได้ทันทีโดย agent หรือ user

---

## Pipeline 1: New section with catalog (Single section)

**Use when:** TOR section has 1 catalog, no sister sections.

**Steps:**

```python
# Stage 1: Inspect catalog (Opus)
import fitz
fitz.TOOLS.mupdf_display_errors(False)
doc = fitz.open("catalog/SOURCE.pdf")
for pi in range(doc.page_count):
    pix = doc[pi].get_pixmap(matrix=fitz.Matrix(0.5, 0.5))
    pix.save(f"/tmp/inspect_p{pi+1}.png")
# View images via Read tool, identify brand logo, model, spec page

# Stage 2: Search text positions
for q in ["BrandName", "ModelName", "Spec keyword 1", ...]:
    rs = doc[pi].search_for(q)
    print(f"{q}: {[(r.x0, r.y0, r.x1, r.y1) for r in rs[:2]]}")

# Stage 3: Annotate output
import shutil, tempfile, os
SECTION = "5.X.Y.Z"
OUT_DIR = "output/.../[section_name]/[sub_folder]"
OUT_NAME = f"{SECTION}. [section_name] [Brand] [Model].pdf"
OUT_PATH = os.path.join(OUT_DIR, OUT_NAME)

os.makedirs(OUT_DIR, exist_ok=True)
shutil.copy2("catalog/SOURCE.pdf", OUT_PATH)
doc = fitz.open(OUT_PATH)
filename_base = OUT_NAME[:-4]

# Header on every page
for pi in range(len(doc)):
    page = doc[pi]
    pw = page.rect.width
    h = page.add_freetext_annot(
        fitz.Rect(15, 10, pw-15, 50),
        f"{filename_base} หน้า {pi+1}",
        fontsize=14 if pw < 1500 else 60,  # 60pt for large pages
        fontname="hebo", text_color=(1, 0, 0),
        align=fitz.TEXT_ALIGN_CENTER)
    h.set_border(width=0); h.update()

# Brand + model rects on cover page
def add_rect(page, rect, content, label, label_rect):
    sq = page.add_rect_annot(rect)
    sq.set_colors(stroke=(1, 0, 0))
    sq.set_border(width=0.8)
    sq.update(opacity=1.0)
    sq.set_info(content=content)
    ft = page.add_freetext_annot(label_rect, label, fontsize=9, fontname="hebo",
                                  text_color=(1, 0, 0), align=fitz.TEXT_ALIGN_LEFT)
    ft.set_border(width=0); ft.update()

p1 = doc[0]
add_rect(p1, fitz.Rect(...), "(brand-X)", "ยี่ห้อ", fitz.Rect(...))
add_rect(p1, fitz.Rect(...), "(model)", "รุ่น", fitz.Rect(...))

# Per ข้อย่อย: rect + label
for i, (rect, label_rect) in enumerate(spec_rects, 1):
    add_rect(spec_page, rect, f"(s{i}-...)", f"{SECTION} ข้อย่อย {i}.", label_rect)

# Save (CRITICAL: garbage=4, clean=True for Google Drive)
tmp = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False); tmp.close()
doc.save(tmp.name, garbage=4, clean=True, deflate=True)
doc.close()
data = open(tmp.name, 'rb').read()
fd = os.open(OUT_PATH, os.O_WRONLY | os.O_TRUNC)
os.write(fd, data); os.close(fd)
os.unlink(tmp.name)

# Stage 4: Update xlsx
from openpyxl import load_workbook
wb = load_workbook("output/Comply spec Smart Plant 1.xlsx")
ws = wb["Trio_SR_Solution"]
CAT = f"{SECTION} [section_name] [Brand] [Model]"

updates = [
    (parent_row, "ยี่ห้อ [Brand] รุ่น [Model]"),
    (sub1_row, f"เทียบเท่าข้อกำหนด เอกสาร {CAT} หน้า 1 ข้อ {SECTION} ข้อย่อย 1."),
    # ... more sub-items
    (commit_row, "ยินดีปฏิบัติตามข้อกำหนด"),
]
for r, d in updates:
    ws.cell(row=r, column=4).value = d
    ws.cell(row=r, column=5).value = "TRIO"  # or SMART/SR
    ws.cell(row=r, column=6).value = "รอ user ตรวจสอบ"
wb.save("output/Comply spec Smart Plant 1.xlsx")
```

---

## Pipeline 2: Clone reference → sister sections

**Use when:** Same catalog used in multiple sister sections (e.g., CCTV in 6 rooms, L3 Switch in 4 rooms).

**Strategy:** Opus annotates ONE reference; Sonnet Agent clones to all sisters.

### 2a. Annotate reference (Opus)

Same as Pipeline 1, but the FIRST sister section becomes the reference.

### 2b. Spawn Sonnet Agent for cloning

**Sonnet Agent prompt template** (use `Agent` tool with `subagent_type: general-purpose, model: sonnet`):

````markdown
Working dir: `/path/to/co-work/claude-code/`

**Task:** Clone reference annotated PDF to N sister sections + update xlsx for all (1+N) sister rows.

**Reference (already annotated):**
`output/.../[REF_SECTION]/.../[REF_SECTION]. [name] [Brand] [Model].pdf`

**Source catalog:** `catalog/[catalog_filename].pdf`

**Sister sections to create:**
- [SISTER1] → `output/.../sister1_dir/`
- [SISTER2] → ...

Filename pattern: `{section}. [name] [Brand] [Model].pdf`

**Cloning approach (PyMuPDF):**
1. Read all annotations from reference PDF (preserve subtype, rect, /Contents tag, FreeText content)
2. For each sister: copy source catalog → add header per page (`{filename} หน้า {page+1}`, [N]pt red center) + clone all rect/label annotations
3. Replace any FreeText content equal to "[REF_SECTION]" with sister section
4. Use `fitz.TOOLS.mupdf_display_errors(False)` + save with `garbage=4, clean=True, deflate=True` + `os.O_TRUNC` write workaround
5. **VERIFY: header is added on EVERY page** (Sonnet has been observed to skip headers — check after)

**xlsx updates (`output/Comply spec Smart Plant 1.xlsx`, sheet "Trio_SR_Solution"):**

For each section, update parent row + N ข้อย่อย rows.

Row ranges (look up by scanning Col B for section pattern):
- [REF_SECTION]: R{parent} parent + R{p+1}-R{p+N} ข้อย่อย
- [SISTER1]: scan, then nested rows
- ...

**Col D format (apply with section# substitution):**
- Parent: `ยี่ห้อ [Brand] รุ่น [Model]`
- ข้อย่อย 1 ([desc]): `เทียบเท่าข้อกำหนด เอกสาร {CAT} หน้า {P} ข้อ {section} ข้อย่อย 1.`
- ... etc per ข้อย่อย
- Generic: `ยินดีปฏิบัติตามข้อกำหนด`

Where `{CAT}` = `{section} [section_name] [Brand] [Model]`

**Col E:** "TRIO" / "SMART" / "SR" (specify) all rows
**Col F:** "รอ user ตรวจสอบ" all rows

Report concise: PDFs created (N), xlsx rows updated, any issues.
````

### 2c. Verify Sonnet output

```python
# Check ALL cloned files have header on every page
for path in cloned_paths:
    doc = fitz.open(path)
    for pi in range(doc.page_count):
        page = doc[pi]
        a = page.first_annot
        has_hdr = False
        while a:
            try:
                c = a.info.get('content', '')
                if a.type[1] == 'FreeText' and 'หน้า' in c and a.rect.y0 < 60:
                    has_hdr = True; break
                a = a.next
            except: break
        assert has_hdr, f"Missing header: {path} page {pi+1}"
```

### 2d. Re-add missing headers (if Sonnet skipped)

```python
# Common Sonnet bug: clone misses headers — re-add
for path in cloned_paths:
    fname = os.path.basename(path)[:-4]
    doc = fitz.open(path)
    needs = []
    for pi in range(doc.page_count):
        page = doc[pi]
        a = page.first_annot
        has_hdr = False
        while a:
            try:
                c = a.info.get('content','')
                if a.type[1]=='FreeText' and 'หน้า' in c and a.rect.y0<60:
                    has_hdr=True; break
                a = a.next
            except: break
        if not has_hdr: needs.append(pi)
    for pi in needs:
        page = doc[pi]
        h = page.add_freetext_annot(
            fitz.Rect(15, 10, page.rect.width-15, 50),
            f"{fname} หน้า {pi+1}",
            fontsize=10, fontname="hebo", text_color=(1, 0, 0),
            align=fitz.TEXT_ALIGN_CENTER)
        h.set_border(width=0); h.update()
    # save with O_TRUNC workaround
```

---

## Pipeline 3: Single-row comply item (no nested ข้อย่อย)

**Use when:** TOR row has no `1) 2) 3)` sub-items (e.g., "5.X.Y.Z. ท่อ HDPE 32mm จำนวน 500 เมตร").

**Col D format** (per SKILL.md กฎข้อ 8):
```
{section} {Col B description minus "จำนวน N <unit>"} {model}
```

**Example:**
- Col B: `5.2.1.13. สายสัญญาณ Twisted Pair Shield แบบใช้ภายในอาคาร จำนวน 500 เมตร`
- Col D: `5.2.1.13 สายสัญญาณ Twisted Pair Shield แบบใช้ภายในอาคาร US-9106LSZH`

**Filename sanitization:** replace `/` with `-` in filenames (e.g., `union emt 3/4"` → `union emt 3-4"` in filename, but Col D keeps original)

**Annotation pattern:**
- Brand rect + `ยี่ห้อ`
- Model rect + `รุ่น`  
- Spec rect(s) + `{section}` (no ข้อย่อย because no nested)
- Custom Thai labels OK (e.g., `ใช้ภายในอาคาร` instead of section number)

---

## Pipeline 4: HTML → PDF catalog conversion

**Use when:** User provides spec as HTML (e.g., Apple iPad spec page from apple.com).

```bash
# Install if needed
pip3 install weasyprint  # already on macOS

# Convert
weasyprint "catalog/spec.html" "/tmp/converted.pdf"
```

**Caveats:**
- Image references from local cache may break (warnings, not errors)
- Text content + key spec render OK
- Number of pages ≈ scrolling page length (could be 20+ for long pages)

**Then proceed with Pipeline 1** (single section with catalog).

---

## Pipeline 5: Verification checkpoint

**Use before backup snapshot or user verify:**

```python
"""Comprehensive verification per SKILL.md rules."""
import fitz, os, re
fitz.TOOLS.mupdf_display_errors(False)
from openpyxl import load_workbook
from collections import Counter

XLSX = "output/Comply spec Smart Plant 1.xlsx"
wb = load_workbook(XLSX, data_only=True)
ws = wb["Trio_SR_Solution"]

# 1. Status counts
status = Counter(ws.cell(row=r, column=6).value for r in range(5, ws.max_row+1))
print("Status:", dict(status))

# 2. Vendor coverage
no_vendor = sum(1 for r in range(5, ws.max_row+1)
                if not ws.cell(row=r, column=5).value
                and ws.cell(row=r, column=6).value)
print(f"Missing vendor: {no_vendor} (should be 0)")

# 3. Col D pattern
PATS = [
    ("เทียบเท่า", re.compile(r'^เทียบเท่าข้อกำหนด')),
    ("สูงกว่า", re.compile(r'^สูงกว่าข้อกำหนด')),
    ("brand-model", re.compile(r'^ยี่ห้อ\s+(?!-)(.+)\s+รุ่น')),
    ("dash-brand", re.compile(r'^ยี่ห้อ\s+-\s+รุ่น')),
    ("commitment", re.compile(r'^ยินดีปฏิบัติตามข้อกำหนด\s*$')),
    ("filename", re.compile(r'^5(?:\.\d+){2,3}[\s\-\.]')),
]
bucket = Counter()
for r in range(5, ws.max_row+1):
    d = ws.cell(row=r, column=4).value
    if not d: bucket["empty"] += 1; continue
    matched = next((n for n, p in PATS if p.match(str(d).strip())), None)
    bucket[matched or "UNKNOWN"] += 1
print("Col D patterns:", dict(bucket))

# 4. Col C cleanliness
compare = ['หรือดีกว่า', 'ไม่น้อยกว่า', 'ต้องสามารถ', 'จะต้อง']
viol = [r for r in range(5, ws.max_row+1)
        if ws.cell(row=r, column=3).value
        and any(w in str(ws.cell(row=r, column=3).value) for w in compare)]
print(f"Col C violations: {len(viol)} (should be 0)")

# 5. PDF integrity
pdf_issues = []
for root, dirs, files in os.walk("output"):
    if "_archive" in root or "_versions" in root: continue
    for f in files:
        if not f.endswith(".pdf") or f.startswith("~$") or "Comply spec Smart Plant" in f: continue
        path = os.path.join(root, f)
        doc = fitz.open(path)
        n_pages = doc.page_count
        pages_ok = 0; n_sq = 0; brand = model = False
        for pi in range(n_pages):
            page = doc[pi]
            ph = page.rect.height
            a = page.first_annot
            page_hdr = False
            cnt = 0
            while a:
                try:
                    cnt += 1
                    if cnt > 200: break
                    c = a.info.get('content','')
                    s = a.type[1]
                    if s == 'FreeText':
                        if 'หน้า' in c and (a.rect.y0 < 200 or a.rect.y1 > ph - 100):
                            page_hdr = True
                        elif c.strip() == 'ยี่ห้อ': brand = True
                        elif c.strip() == 'รุ่น': model = True
                    elif s == 'Square': n_sq += 1
                    a = a.next
                except: break
            if page_hdr: pages_ok += 1
        doc.close()
        # Acceptable exceptions: vibration sensor, pole, install commitments
        if pages_ok < n_pages: pdf_issues.append((f[:60], f"hdr {pages_ok}/{n_pages}"))

print(f"PDF issues: {len(pdf_issues)}")
for fn, prob in pdf_issues[:10]: print(f"  ⚠ {fn}: {prob}")
```

---

## Pipeline 6: Snapshot management

```bash
# Quick snap (xlsx + SKILL.md, ~150KB) — daily/per work session
python3 scripts/version.py snap "before-section-X-work"

# Full snap (+ output/ tar.gz, ~190MB) — milestones
python3 scripts/version.py snap-full "milestone-section-X-done"

# Auto-snap (only if xlsx changed)
python3 scripts/version.py auto-snap

# List + diff
python3 scripts/version.py list
python3 scripts/version.py diff <id1> <id2>

# Restore
python3 scripts/version.py restore <id>          # quick
python3 scripts/version.py restore-full <id>     # full

# Cleanup (keep last 30)
python3 scripts/version.py prune --keep 30 -y
```

---

## Pipeline 7: Vendor inheritance fix

**Use when:** Sub-item rows have status (Col F) but no vendor (Col E) — common after xlsx edit by sub-section.

```python
from openpyxl import load_workbook
wb = load_workbook("output/Comply spec Smart Plant 1.xlsx")
ws = wb["Trio_SR_Solution"]

current_vendor = None
fixes = 0
for r in range(5, ws.max_row+1):
    e = ws.cell(row=r, column=5).value
    f = ws.cell(row=r, column=6).value
    if e:
        current_vendor = e
    elif f and current_vendor:
        ws.cell(row=r, column=5).value = current_vendor
        fixes += 1
wb.save("output/Comply spec Smart Plant 1.xlsx")
print(f"Fixed {fixes} rows")
```

---

## Pipeline 8: Adding new "สูงกว่าข้อกำหนด" rows

**Detect when catalog spec exceeds TOR:**
- TOR: `ไม่น้อยกว่า X` → catalog: `Y > X` → use `สูงกว่าข้อกำหนด`
- TOR: `ไม่มากกว่า X` → catalog: `Y < X` → use `สูงกว่าข้อกำหนด` (better than spec)
- TOR: `≥ X cores` → catalog: `Y > X cores` → `สูงกว่าข้อกำหนด`

**Format same as เทียบเท่า:**
```
สูงกว่าข้อกำหนด เอกสาร {CAT} หน้า {P} ข้อ {section} ข้อย่อย N.
```

**Examples seen in this project:**
- IP66 → catalog IP67 ✓
- Image Sensor 1/3" → catalog 1/2.8" ✓
- MAC Address ≥ 32K → catalog 64K ✓
- Switching Capacity ≥ 1 Tbps → catalog 1.44 Tbps ✓
- CPU 5 core → catalog 6 core (A16) ✓
- Network Interface 4G → catalog 5G ✓
