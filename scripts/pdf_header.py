#!/usr/bin/env python3
"""
pdf_header.py  –  Add header/footer stamps to PDF files (macOS)
================================================================
ใช้แทน Adobe Acrobat's "Add Header & Footer" feature

INSTALL (ครั้งแรกครั้งเดียว):
    pip3 install pymupdf

USAGE:
    python3 pdf_header.py [options] <file_or_glob> [file_or_glob ...]

EXAMPLES:
    # Header อัตโนมัติ: ชื่อไฟล์ + หน้า N (default)
    python3 pdf_header.py "output/**/*-1.pdf"

    # กำหนด template เอง
    python3 pdf_header.py --text "{name} หน้า {page}/{pages}" somefile.pdf

    # เฉพาะ header ไม่มี box รอบ
    python3 pdf_header.py --no-box "output/5.1.2.-1/*.pdf"

    # วาง footer แทน header
    python3 pdf_header.py --position bottom somefile.pdf

    # ดูตัวอย่างก่อนโดยไม่บันทึก
    python3 pdf_header.py --dry-run somefile.pdf

TEXT TEMPLATE VARIABLES:
    {name}   = ชื่อไฟล์ (ไม่มี .pdf)
    {page}   = เลขหน้าปัจจุบัน
    {pages}  = จำนวนหน้าทั้งหมด
    {dir}    = ชื่อโฟลเดอร์ parent

OPTIONS:
    --text TEXT        Template ข้อความ  (default: "{name} หน้า {page}")
    --position top|bottom  ตำแหน่ง      (default: top)
    --align left|center|right  การจัดข้อความ (default: left)
    --font-size N      ขนาดตัวอักษร pt   (default: 9)
    --margin N         ระยะขอบจากขอบหน้า (default: 15)
    --box-height N     ความสูง box pt    (default: 20)
    --no-box           ไม่วาด box รอบข้อความ
    --as-annot         ใส่เป็น annotation แทน content stamp
    --dry-run          แสดงผลลัพธ์โดยไม่บันทึก
    --overwrite        เขียนทับไฟล์เดิม (default: บันทึกเป็น *_header.pdf)
    --suffix SUFFIX    suffix ของไฟล์ output  (default: _header)
"""

import argparse
import glob
import os
import sys

# ── auto-install pymupdf ──────────────────────────────────────────────────────
try:
    import fitz  # PyMuPDF
except ImportError:
    print("Installing PyMuPDF...")
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pymupdf", "-q"])
    import fitz

# ── alignment map ─────────────────────────────────────────────────────────────
ALIGN = {"left": fitz.TEXT_ALIGN_LEFT,
         "center": fitz.TEXT_ALIGN_CENTER,
         "right": fitz.TEXT_ALIGN_RIGHT}


def process_file(path: str, args) -> bool:
    """Add header/footer to a single PDF. Returns True on success."""

    if not os.path.isfile(path):
        print(f"  SKIP (not found): {path}")
        return False

    filename = os.path.splitext(os.path.basename(path))[0]
    parent   = os.path.basename(os.path.dirname(path))

    doc = fitz.open(path)
    total_pages = len(doc)

    print(f"  {filename}.pdf  ({total_pages} page{'s' if total_pages > 1 else ''})")

    for page_num, page in enumerate(doc, start=1):
        text = args.text.format(
            name  = filename,
            page  = page_num,
            pages = total_pages,
            dir   = parent,
        )

        pw = page.rect.width
        ph = page.rect.height
        margin = args.margin
        box_h  = args.box_height

        if args.position == "top":
            y0 = ph - margin - box_h   # fitz: y=0 at bottom, up is positive
            y1 = ph - margin
        else:  # bottom
            y0 = margin
            y1 = margin + box_h

        text_rect = fitz.Rect(margin, y0, pw - margin, y1)

        if args.as_annot:
            # ── FreeText annotation (like what we did manually) ──────────────
            annot = page.add_freetext_annot(
                text_rect,
                text,
                fontsize   = args.font_size,
                text_color = (0, 0, 0),
                fill_color = (1, 1, 0),   # yellow background
                border_color = (0, 0, 0),
            )
        else:
            # ── Content stamp (permanent, like Adobe Acrobat) ─────────────────
            if not args.no_box:
                page.draw_rect(
                    text_rect,
                    color = (0, 0, 0),    # black border
                    fill  = (1, 1, 0.7),  # light yellow fill
                    width = 0.5,
                )
            page.insert_textbox(
                text_rect,
                text,
                fontsize  = args.font_size,
                fontname  = "helv",
                color     = (0, 0, 0),
                align     = ALIGN[args.align],
            )

        if args.dry_run:
            print(f"    p{page_num}: [{text_rect.x0:.0f},{text_rect.y0:.0f},"
                  f"{text_rect.x1:.0f},{text_rect.y1:.0f}] → {text!r}")

    if args.dry_run:
        print("    (dry-run — not saved)")
        doc.close()
        return True

    # ── save ──────────────────────────────────────────────────────────────────
    dir_  = os.path.dirname(path)
    base_ = os.path.splitext(os.path.basename(path))[0]

    if args.overwrite:
        out_path = path
    else:
        out_path = os.path.join(dir_, f"{base_}{args.suffix}.pdf")

    doc.save(out_path, garbage=4, deflate=True)
    doc.close()

    if out_path != path:
        print(f"    → {os.path.basename(out_path)}")
    else:
        print(f"    → overwritten")

    return True


def main():
    parser = argparse.ArgumentParser(
        description="Add header/footer to PDF files (like Adobe Acrobat)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("files", nargs="+",
                        help="PDF files or glob patterns")
    parser.add_argument("--text",      default="{name} หน้า {page}",
                        help="Text template (default: '{name} หน้า {page}')")
    parser.add_argument("--position",  choices=["top", "bottom"], default="top")
    parser.add_argument("--align",     choices=["left", "center", "right"], default="left")
    parser.add_argument("--font-size", type=float, default=9, dest="font_size")
    parser.add_argument("--margin",    type=float, default=15)
    parser.add_argument("--box-height",type=float, default=20, dest="box_height")
    parser.add_argument("--no-box",    action="store_true", dest="no_box")
    parser.add_argument("--as-annot",  action="store_true", dest="as_annot",
                        help="Insert as FreeText annotation instead of content stamp")
    parser.add_argument("--dry-run",   action="store_true", dest="dry_run")
    parser.add_argument("--overwrite", action="store_true")
    parser.add_argument("--suffix",    default="_header")

    args = parser.parse_args()

    # expand globs
    paths = []
    for pattern in args.files:
        matched = sorted(glob.glob(pattern, recursive=True))
        if matched:
            paths.extend(matched)
        else:
            paths.append(pattern)  # will show SKIP

    paths = [p for p in paths if p.endswith(".pdf")]
    if not paths:
        print("No PDF files found.")
        sys.exit(1)

    ok = err = 0
    for p in paths:
        print(f"\n[{ok+err+1}/{len(paths)}] {p}")
        if process_file(p, args):
            ok += 1
        else:
            err += 1

    print(f"\n{'='*50}")
    print(f"Done: {ok} succeeded, {err} failed")


if __name__ == "__main__":
    main()
