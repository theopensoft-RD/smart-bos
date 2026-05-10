# SR's Annotation Pattern — Study Reference

> **Read-only study** of how SR (Smart Solution Co., Ltd.) annotates their proposed catalogs.
> ใช้เป็น reference สำหรับ agent — **ห้ามแก้ catalog ของ SR**
>
> Generated from `catalog/SR/extracted/งานโรงบำบัดน้ำเสีย/` (10 SR-proposed catalogs)

---

## 📊 Annotation Type Distribution per Catalog

| SR Catalog | Highlights | Callouts | Rects | Pattern |
|---|---:|---:|---:|---|
| Server H3C R4900 G7 | 5 | 10 | 7 | mixed (text+image areas) |
| L2 Switch TP-Link | 6 | 6 | 1 | text-heavy |
| NAS Synology DS425+ | 15 | 7 | 0 | **all-highlight** |
| UPS Cleanline T-10K | 8 | 2 | 3 | mixed |
| **9U Rack** | 0 | 5 | 8 | **all-rect** (image-only PDF) |
| NVR Dahua NVR5432-EI2 | 26 | 14 | 0 | **all-highlight** (heavy) |
| **CCTV Dahua HFW2441T** | 37 | 24 | 0 | **all-highlight** (heaviest) |
| Core Switch H3C S5590-EI | 10 | 8 | 3 | mixed |
| PoE Switch Dahua DH-CS4220 | 8 | 8 | 1 | text-heavy |

**Total:** 115 highlights, 84 callouts, 23 rects across 9 readable catalogs

---

## 🎯 Key Conventions

### 1. Annotation Type Selection — by PDF type

```
if pdf_is_text_based:
    use highlight() + callout
elif pdf_is_image_based:
    use rect() + callout
elif mixed:
    use both as appropriate
```

**Detection rule:** average chars per page (using `page.get_text()`)
- **>500 chars/page** = text-based → highlight
- **<100 chars/page** = image-based → rect
- **Between** = mixed, prefer highlight

### 2. Callout format — SHORT

| SR's Format | Example | Meaning |
|---|---|---|
| `N)` | `4)`, `7)` | ข้อย่อย N (most common, 95%) |
| `N) <desc>` | `6) รองรับ IPv6` | When extra description helps reviewer |
| `N` (no paren) | `7` | Rare variant |

**No section number prefix** (e.g., never see "5.1.1.2 ข้อย่อย 4." — always just `4)`)
- Section is **implicit** from folder location
- Catalog is in `5.1.1.-2/` folder → all callouts in this PDF refer to section 5.1.1.2

### 3. Callout Position

- **Right margin** (page width − ~25pt) — most common
- **Inside table** in white space cells
- **Near highlighted text** with arrow line if added
- **Avoid** overlapping with content

### 4. Highlight Scope

**Highlight ONLY value/keyword** — not entire row

✅ Good (SR style):
```
Spec row: "Maximum number of MAC address entries: 16000"
Highlight: only "16000"
```

❌ Avoid (over-highlighting):
```
Spec row: "Maximum number of MAC address entries: 16000"
Highlight: entire row
```

### 5. Multiple instances of same callout

Same `N)` can appear multiple times if same TOR ข้อย่อย matches multiple spec mentions.

Example from L2 Switch:
- `2)` appears at 2 positions in P8 (spec mentioned twice)
- All highlighted

### 6. Rect (Square) usage

Used **only** when:
- PDF is **image-only** (no text layer)
- Or to mark **brand logo / product image** (graphic content)

Border: red, ~1pt width

---

## 💡 Why SR's Approach is Better

| Metric | SR Hybrid (highlight+rect) | Rect-only (old SKILL.md) |
|---|---|---|
| Time per spec | < 1 sec (auto search_for) | 5-30 sec (manual coord design) |
| Coord accuracy | Always exact (from text engine) | Depends on visual judgment |
| Print clarity | Yellow highlight stands out | Red border less noticeable |
| Content blocking | None (highlight overlay) | Border can hide adjacent text |
| Maintenance | Re-highlight on layout change works | Manual recoord on layout change |

---

## 🛠 How to Apply (for our own catalogs — NOT SR's)

```python
import fitz
fitz.TOOLS.mupdf_display_errors(False)

doc = fitz.open("our_catalog.pdf")

# Detect PDF type
text_density = sum(len(doc[pi].get_text()) for pi in range(doc.page_count)) / doc.page_count

for pi in range(doc.page_count):
    page = doc[pi]
    if text_density > 500:
        # Text-based: highlight + short callout
        for keyword, sub_n in spec_map.items():
            rects = page.search_for(keyword)
            for r in rects:
                h = page.add_highlight_annot(r)
                h.set_colors(stroke=(1, 1, 0))
                h.update()
            if rects:
                # Callout (SR-style: short, white bg)
                lbl = fitz.Rect(page.rect.width - 25, rects[0].y0, page.rect.width - 5, rects[0].y0 + 12)
                ft = page.add_freetext_annot(lbl, f"{sub_n})", fontsize=9, fontname="hebo",
                    text_color=(1, 0, 0), fill_color=(1, 1, 1),
                    align=fitz.TEXT_ALIGN_CENTER)
                ft.set_border(width=0); ft.update()
    else:
        # Image-based: rect + callout
        for rect_coord, sub_n in rect_map.items():
            sq = page.add_rect_annot(rect_coord)
            sq.set_colors(stroke=(1, 0, 0))
            sq.set_border(width=1.0)
            sq.update()
            # Same callout style
            ...
```

---

## ⚠️ DO NOT modify SR's catalogs

The SR-provided catalogs in `catalog/SR/extracted/` already have SR's annotations.
- **Read** them to learn the pattern
- **Copy** them to `output/TRIO_SR_Solution/` as-is
- **Add only** our header on top (not on top of SR's annotations)
- **Never** add highlights/rects on top of what SR annotated

If we need to verify a comply spec, the SR-annotated catalog already shows where the spec matches — we just reference it in xlsx Col D.

---

## 🆕 (2026-05-10) Adoption Results — TRIO_SR_Solution

After adopting SR pattern across all 104 PDFs in `output/TRIO_SR_Solution/`:

| Status | Count | Notes |
|---|---|---|
| SR pattern (highlight/rect + short callout) | 75 | 22 SR catalogs + 4 L3 Switch + 1 NVR fix + 48 our converted |
| Brand-marker only (single-row) | 27 | Cable/conduit/UF-2010A — already SR-style |
| Header-only placeholder | 2 | BOD/DO Sensor SR proposals (1-page summary) |
| Empty / duplicate / long label | 0 | All cleaned |

### Where we annotated ourselves using SR pattern

SR did not annotate these — we filled in using SR's pattern:
- **L3 Switch H3C S5170-EI** (5.2.1.6, 5.2.2.2, 5.2.3.2, 5.2.4.2 sisters)
  - Item 3 (`24*10/100/1000BASE-T Ports`) on page 18 → callout `3)`
  - Item 6 (`MAC address entries: 32768`) on page 12 → callout `6)`
- **NVR Dahua DHI-NVR5432-EI2** (5.2.1.7) — SR forgot item 5
  - Item 5 (`1920 × 1080` resolution) on page 2 → callout `5)`

### SR's incomplete coverage (flagged via "ไม่พบใน catalog")

| Issue | Count |
|---|---|
| CCTV ข้อ 19-21 (manufacturer standards) | 6 sisters × 3 = 18 rows |
| PoE Switch / AP ข้อ 7 | 8 sisters × 1 = 8 rows |
| L3 Switch ข้อ 1,2,4,5,7,8,9 (only 3,6 in catalog) | 4 sisters × 7 = 28 rows |
| HDPE pipe sections | 3 sections |
| Other SR partial coverage | ~50 rows |
| **Total flagged for user review** | **115 rows (17.4%)** |

### Key Implementation Note: PyMuPDF Annot Deletion

When converting our long-format labels to SR short callouts, standard `delete_annot` didn't work reliably. **The working approach:**

```python
# Step 1: collect all annot specs
specs = collect_annots(doc)

# Step 2: clear /Annots arrays via direct xref manipulation
for pi in range(doc.page_count):
    page = doc[pi]  # keep page ref alive
    doc.xref_set_key(page.xref, "Annots", "[]")

# Step 3: re-add only the converted annotations
for spec in specs:
    page = doc[spec.page_idx]
    if spec.type == 'FreeText' and is_long_label(spec.content):
        new_content = shorten_to_short_callout(spec.content)
        small_rect = fitz.Rect(spec.rect.x0, spec.rect.y0,
                               spec.rect.x0 + 22, spec.rect.y0 + 12)
        ft = page.add_freetext_annot(small_rect, new_content,
            fontsize=9, fontname="hebo",
            text_color=(1, 0, 0), fill_color=(1, 1, 1),
            align=fitz.TEXT_ALIGN_CENTER)
        ft.set_border(width=0); ft.update()
```

See `pitfalls.md` §1.1b/c for full bug analysis (page-ref-alive + xref_set_key approach).
