#!/usr/bin/env python3
"""
fix_uv_headers.py
-----------------
แก้ header annotation ของ UV cabinet -1 ทุกไฟล์:
  - ลบ header เดิมที่ใส่ด้วย pypdf (annotation ที่ top)
  - stamp ข้อความใหม่ top-center, สีแดง bold เหมือน annotation ล่าง
  - ขนาด font auto-fit ระหว่าง 16-18pt ให้พอดีหน้ากระดาษ

RUN (ครั้งแรกจะ install pymupdf อัตโนมัติ):
    python3 fix_uv_headers.py
"""

import os, sys

try:
    import fitz
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pymupdf", "-q"])
    import fitz

# ── กำหนดไฟล์ ─────────────────────────────────────────────────────────────────
BASE = os.path.join(
    os.path.expanduser("~"),
    "Library", "CloudStorage",
    "GoogleDrive-smartsolution.in.th@gmail.com",
    "My Drive", "Smart Solution", "Project",
    "Pattaya Project", "Smart Plant 1", "co-work", "output",
)

FILES = [
    (
        "5.1.2-1 ตู้ควบคุมเครื่องจักรสำหรับควบคุมเครื่องจักรประเภทต่างๆภายในโรงบำบัดน้ำเสีย หน้า 1",
        BASE + "/5.1.2. ตู้ควบคุมเครื่องจักรสำหรับควบคุมเครื่องจักรประเภทต่างๆภายในโรงบำบัดน้ำเสีย"
               "/5.1.2.-1"
               "/5.1.2. ตู้ควบคุมเครื่องจักรสำหรับควบคุมเครื่องจักรประเภทต่างๆภายในโรงบำบัดน้ำเสีย-1.pdf",
    ),
    (
        "5.1.3.1-1 ตู้ควบคุมเครื่องจักรสำหรับควบคุมเครื่องสูบน้ำเสียต่างๆที่อยู่ภายนอกโรงบำบัดน้ำเสีย หน้า 1",
        BASE + "/5.1.3. ตู้ควบคุมเครื่องจักรสำหรับควบคุมเครื่องสูบน้ำเสียต่างอยู่ภายนอกโรงบำบัดน้ำเสีย"
               "/5.1.3.1. ตู้ควบคุมเครื่องจักรสำหรับควบคุมเครื่องสูบน้ำเสียต่างๆที่อยู่ภายนอกโรงบำบัดน้ำเสีย"
               "/5.1.3.1. -1"
               "/5.1.3.1. ตู้ควบคุมเครื่องจักรสำหรับควบคุมเครื่องสูบน้ำเสียต่างๆที่อยู่ภายนอกโรงบำบัดน้ำเสีย-1.pdf",
    ),
    (
        "5.1.4.1-1 ตู้ประมวลผลอุปกรณ์อ่านค่า Biochemical Oxygen Demand (BOD) หน้า 1",
        BASE + "/5.1.4. ตู้ประมวลผลความต้องการออกซิเจนทางชีวเคมี"
               "/5.1.4.1. ตู้ประมวลผลอุปกรณ์อ่านค่า Biochemical Oxygen Demand (BOD)"
               "/5.1.4.1.-1"
               "/5.1.4.1. ตู้ประมวลผลอุปกรณ์อ่านค่า Biochemical Oxygen Demand (BOD)-1.pdf",
    ),
    (
        "5.1.5.1-1 ตู้ประมวลผลอุปกรณ์อ่านค่า Dissolved Oxygen (DO) หน้า 1",
        BASE + "/5.1.5 ตู้ประมวลผลค่าออกซิเจนละลายน้ำ"
               "/5.1.5.1. ตู้ประมวลผลอุปกรณ์อ่านค่า Dissolved Oxygen (DO)"
               "/5.1.5.1.-1"
               "/5.1.5.1. ตู้ประมวลผลอุปกรณ์อ่านค่า Dissolved Oxygen (DO)-1.pdf",
    ),
    (
        "5.1.6.1-1 ตู้ควบคุมการแสดงผลของการบำบัดน้ำเสีย หน้า 1",
        BASE + "/5.1.6 ตู้ควบคุมสถานีแสดงผลการบำบัดน้ำเสียอัจฉริยะ"
               "/5.1.6.1. ตู้ควบคุมการแสดงผลของการบำบัดน้ำเสีย"
               "/5.1.6.1.-1"
               "/5.1.6.1. ตู้ควบคุมการแสดงผลของการบำบัดน้ำเสีย-1.pdf",
    ),
]

RED  = (1, 0, 0)       # #FF0000 — เหมือน annotation ล่าง
FONT = "hebo"          # Helvetica Bold (built-in fitz font)
TOP_MARGIN   = 10      # pt จากขอบบน
HEADER_PAD   = 8       # pt padding บน-ล่างใน box


def choose_fontsize(page_w: float) -> int:
    """16-18pt ตาม page width:  A4=595 → 17,  A3=842 → 18,  เล็กกว่า A4 → 16"""
    if page_w >= 800:
        return 18
    elif page_w >= 580:
        return 17
    else:
        return 16


def remove_old_header_annots(page: fitz.Page) -> int:
    """ลบ annotation ที่ top ของหน้า (ที่ใส่ด้วย pypdf ก่อนหน้า)"""
    ph = page.rect.height
    removed = 0
    to_del = []
    for annot in page.annots():
        r = annot.rect
        # annotation ที่ y0 < 60pt จากบน (top area)
        if r.y0 < 60:
            to_del.append(annot)
    for annot in to_del:
        page.delete_annot(annot)
        removed += 1
    return removed


def add_header_stamp(page: fitz.Page, text: str):
    """stamp text ที่ top-center สีแดง bold, auto-fit font 16-18pt"""
    pw = page.rect.width

    font_size = choose_fontsize(pw)
    line_h    = font_size + HEADER_PAD

    # rect: full width minus 15pt margin, centered at top
    rect = fitz.Rect(15, TOP_MARGIN, pw - 15, TOP_MARGIN + line_h * 2)

    # ลองใส่ข้อความ — ถ้าไม่พอ (return < 0) ลด font ลงทีละ 1
    for fs in range(font_size, 13, -1):
        rc = page.insert_textbox(
            rect,
            text,
            fontsize  = fs,
            fontname  = FONT,
            color     = RED,
            align     = fitz.TEXT_ALIGN_CENTER,
        )
        if rc >= 0:
            print(f"    font {fs}pt ✓")
            break
    else:
        # ถ้ายังไม่พอ ใช้ขนาดเล็กสุดโดยไม่ตรวจ
        page.insert_textbox(rect, text, fontsize=13, fontname=FONT,
                            color=RED, align=fitz.TEXT_ALIGN_CENTER)
        print(f"    font 13pt (fallback)")


# ── Main ───────────────────────────────────────────────────────────────────────
for header_text, path in FILES:
    print(f"\n{'─'*60}")
    print(f"File: {os.path.basename(path)}")

    if not os.path.exists(path):
        print("  ERROR: file not found")
        continue

    doc  = fitz.open(path)
    page = doc[0]

    removed = remove_old_header_annots(page)
    print(f"  Removed {removed} old header annotation(s)")

    add_header_stamp(page, header_text)

    doc.save(path, garbage=4, deflate=True, incremental=False)
    doc.close()
    print(f"  Saved ✓")

print(f"\n{'='*60}")
print("Done.")
