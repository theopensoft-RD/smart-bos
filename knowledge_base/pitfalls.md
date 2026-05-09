# Pitfalls — Lessons Learned

> Common issues encountered + their solutions. Read before starting new work to avoid these.

---

## 1. PDF Annotation Pitfalls

### 1.1 Phantom annotations (xref=0)

**Symptom:** Inspecting a PDF shows annotations but `delete_annot()` doesn't remove them.

**Cause:** Stale annotation objects from previous saves without garbage collection.

**Fix:** When writing PDFs, ALWAYS use:
```python
doc.save(path, garbage=4, clean=True, deflate=True)
```

For severely corrupted phantoms — rebuild from source:
```python
shutil.copy2(SOURCE_CATALOG, OUT_PATH)
doc = fitz.open(OUT_PATH)
# add fresh annotations
```

### 1.2 MuPDF "object out of range" warnings

**Symptom:** Lots of `MuPDF error: format error: object out of range...` in stderr.

**Cause:** Source PDF has unusual cross-reference table.

**Fix:** Suppress with:
```python
fitz.TOOLS.mupdf_display_errors(False)
```
or run shell command with `2>/dev/null`. **Output is still saved correctly** — these are non-fatal warnings.

### 1.3 Header font too small for large pages

**Symptom:** Header text barely visible on large image-based PDFs (e.g., 2480×3507).

**Cause:** Default `fontsize=14` works for A4 (595×842) but is too small for 4× larger pages.

**Fix:** Scale font proportionally:
```python
is_large = pw > 1500
fontsize = 60 if is_large else 14  # 60pt for 2480-wide pages
```

### 1.4 Thai text not rendering in label

**Symptom:** Thai labels appear as boxes or missing characters.

**Cause:** Wrong font. PyMuPDF's built-in fonts (`helv`, `cour`) don't have Thai glyphs.

**Fix:** Use `fontname="hebo"` (Helvetica-Bold) — Adobe Reader has Thai font fallback that uses this:
```python
page.add_freetext_annot(rect, "ยี่ห้อ", fontsize=9, fontname="hebo", text_color=(1, 0, 0))
```

PyMuPDF's `pixmap` rendering may show Thai correctly — depends on system fonts.

### 1.5 Header detected at WRONG position (top vs bottom)

**Symptom:** Verification reports "missing header" but PDF visually has one.

**Cause:** Some PDFs (e.g., G7 outlets/fans) place header at BOTTOM (y > ph - 100) due to top title strip.

**Fix:** Verification should check both positions:
```python
if 'หน้า' in c and (a.rect.y0 < 200 or a.rect.y1 > page.rect.height - 100):
    has_header = True
```

### 1.6 Sonnet Agent skips header on cloning

**Symptom:** Sonnet Agent successfully clones rect+labels but forgets to add header.

**Cause:** Sonnet may interpret prompt loosely.

**Fix:** ALWAYS verify header presence after spawning Agent. If missing, re-add via Pipeline 2d (see `pipelines.md`).

### 1.7 CCTV catalog has pre-baked annotations

**Symptom:** Cloning Dahua CCTV catalog brings duplicate Highlight + numbered FreeText labels.

**Cause:** Source catalog has `(ICT)` suffix — pre-annotated by Thai gov ICT TOR. Highlights + "4.1, 4.2..." labels are PRE-EXISTING in source.

**Fix:** **DO NOT REMOVE** pre-existing annotations — ICT auditor may reference them. Just add OUR rect + labels (`{section} ข้อย่อย N.`) on top.

---

## 2. xlsx Pitfalls

### 2.1 Excel lock file blocks edit

**Symptom:** `~$Comply spec Smart Plant 1.xlsx` exists.

**Cause:** Excel currently has the file open.

**Fix:** Close Excel before running Python edits. The lock file is auto-cleaned when Excel closes. **Don't delete it manually** while Excel is running.

### 2.2 Sheet name changes between sessions

**Symptom:** `KeyError: 'SmartPlant1' does not exist`.

**Cause:** User renamed sheet (e.g., `SmartPlant1` → `Trio_SR_Solution`).

**Fix:** Always check sheet names first:
```python
wb = load_workbook(XLSX, data_only=True)
print(wb.sheetnames)  # ['Trio_SR_Solution', 'Take IT']
ws = wb[wb.sheetnames[0]]  # use first sheet, or specify
```

### 2.3 Vendor not inherited to ข้อย่อย

**Symptom:** 175 rows have status but no vendor in Col E.

**Cause:** Updating xlsx by parent only — forgot to set Col E for nested ข้อย่อย.

**Fix:** Run vendor inheritance script (Pipeline 7) — walks rows, carries forward vendor from latest parent with vendor.

### 2.4 Col C contains comparison words

**Symptom:** `จะต้อง`, `ต้องสามารถ`, `หรือดีกว่า`, `ไม่น้อยกว่า` etc. appear in Col C.

**Cause:** Forgot to clean comparison words when copying from Col B (TOR).

**Fix:** Apply replacement table (per SKILL.md):
| Pattern in B | Action for C |
|---|---|
| `หรือดีกว่า` | remove |
| `ไม่น้อยกว่า` | remove (keep value) |
| `ไม่น้อยไปกว่า` | remove |
| `ไม่มากกว่า` | remove |
| `ต้องสามารถ` | replace with `สามารถ` |
| `จะต้อง` | remove |

### 2.5 Col D ref doesn't match PDF filename

**Symptom:** Cross-ref check finds many "missing refs" but PDFs exist.

**Cause:** Naming convention mismatch:
- Old style ref: `5.1.2-1 NAME` (matches `5.1.2. NAME-1.pdf`)
- New style ref: `5.X.Y.Z NAME` (matches `5.X.Y.Z. NAME.pdf`)

**Fix:** Smart matching with multiple patterns (see Pipeline 5):
```python
# Pattern 1: Direct
# Pattern 2: "5.1.2-1 NAME" → "5.1.2. NAME-1"
m = re.match(r'^(5(?:\.\d+){1,3})-(\d+)\s+(.+)$', ref)
if m:
    sec, num, name = m.groups()
    target = f"{sec}. {name}-{num}"
```

---

## 3. File System Pitfalls

### 3.1 Forbidden chars in filenames

**Symptom:** Cannot create file with `union emt 3/4"`.

**Cause:** `/` is path separator on Linux/macOS — forbidden in filenames.

**Fix:** Replace `/` with `-` in filename ONLY (Col D in xlsx keeps original):
```python
filename = name.replace("/", "-")
# union emt 3/4" → union emt 3-4"
```

### 3.2 Long filenames exceed NAME_MAX

**Symptom:** Cannot create file with very long name.

**Cause:** Linux NAME_MAX = 255 bytes. Thai chars are 3 bytes UTF-8 → ~85 Thai chars max.

**Fix:** Keep folder names SHORT (e.g., `5.2.1.-9` not full section name) and put long descriptions in actual filenames at the bottom.

### 3.3 Google Drive sync conflicts with Python writes

**Symptom:** File seems to write but Excel reports "out of date".

**Cause:** Google Drive locks file during sync.

**Fix:** Use `os.O_TRUNC` write workaround:
```python
fd = os.open(path, os.O_WRONLY | os.O_TRUNC)
os.write(fd, data)
os.close(fd)
```

### 3.4 .git folder + Google Drive incompatibility

**Symptom:** Considering git for version control.

**Issue:** `.git/objects/*` files conflict with Google Drive sync (lock files, malware false-positives).

**Fix:** Use `scripts/version.py` (snapshot-based, file-only) instead of git. Regular folders sync fine with Drive.

---

## 4. PyMuPDF API Pitfalls

### 4.1 `page.annots()` generator can fail

**Symptom:** `AttributeError: 'NoneType' object has no attribute 'm_internal'` on `page.annots()`.

**Cause:** Some PyMuPDF versions have bug in iterator.

**Fix:** Use linked-list iteration:
```python
a = page.first_annot
while a:
    try:
        # process a
        a = a.next
    except Exception:
        break
```

### 4.2 `delete_annot()` during iteration

**Symptom:** Iterator becomes invalid after delete.

**Fix:** Collect first, then delete:
```python
to_delete = []
a = page.first_annot
while a:
    if should_delete(a):
        to_delete.append(a)
    a = a.next
for a in to_delete:
    page.delete_annot(a)
```

### 4.3 `add_freetext_annot()` requires `update()`

**Symptom:** Annotation appears in inspection but not visually.

**Cause:** New annotation needs appearance stream regenerated.

**Fix:** ALWAYS call `update()`:
```python
ft = page.add_freetext_annot(rect, text, ...)
ft.set_border(width=0)
ft.update()  # critical
```

---

## 5. Workflow Pitfalls

### 5.1 Forgetting to verify after Sonnet Agent

**Symptom:** Sonnet Agent reports "Done" but output is incomplete (e.g., missing headers).

**Fix:** ALWAYS verify after spawning Agent:
```python
# After Sonnet says done, check:
- Annotation count matches reference
- Headers on every page
- Section# substituted correctly
```

If issues found, re-fix in Opus directly (faster than re-prompt Agent).

### 5.2 Updating xlsx while Excel is open

**Fix:** Close Excel, then run Python edit, then user re-opens. Or use file watch:
```bash
ls "~$*.xlsx" 2>/dev/null && echo "Excel still open"
```

### 5.3 Snapshot before destructive ops

**Fix:** Before any major change (rebuilding PDFs, mass xlsx edit):
```bash
python3 scripts/version.py snap-full "before-X-change"
```

### 5.4 Trusting "ยินดีปฏิบัติ" without verification

**Symptom:** Used `ยินดีปฏิบัติ` for items that catalog DOES support → should be `เทียบเท่า` or `สูงกว่า`.

**Fix:** When marking ยินดีปฏิบัติ, briefly check if catalog has the spec. If catalog explicitly mentions it (even without precise value), use `เทียบเท่า` with rect.

`ยินดีปฏิบัติ` should be reserved for:
1. Software/install commitments (no spec needed)
2. Catalog truly doesn't mention the requirement
3. Generic capabilities not measurable

---

## 6. Vendor-specific Pitfalls

### 6.1 LINK UTP/STP confusion

**Symptom:** TOR says "Twisted Pair Shield" (STP) but vendor selected `US-9106LSZH` which is U/UTP (unshielded).

**Issue:** Mismatch between TOR requirement and selected model.

**Action:** Flag for user verify. The user/vendor selected this model — annotation is mechanical task. Mark as `รอ user ตรวจสอบ` for them to confirm.

### 6.2 Ruijie L3 marketed as L2

**Symptom:** TOR says "L2 Switch" but RG-NBS5100 is "Layer 3 Non-PoE Switch".

**Action:** L3 includes L2 capability — generally OK. Annotation should reference Layer 3 cell + label that covers both.

### 6.3 Lenovo SR630 V4 vs SR630

**Symptom:** TOR says "SR630" but catalog is "SR630 V4".

**Action:** V4 is current generation — supports all V1-V3 features + more. Mark filename with V4 designation. User to verify model variant matches their procurement.

---

## 7. Recovery Pitfalls

### 7.1 Restoring from snapshot loses recent edits

**Issue:** `restore` overwrites current xlsx.

**Fix:** Take a snap of CURRENT state first, then restore:
```bash
python3 scripts/version.py snap "before-restore"
python3 scripts/version.py restore <old_id>
# if went wrong:
python3 scripts/version.py restore "before-restore"
```

### 7.2 .bak file overwriting

**Issue:** Multiple `.bak` files accumulate.

**Fix:** Use timestamp/tag in snapshot tag instead of `.bak`. The `version.py` system handles this — use it instead of manual backups.
