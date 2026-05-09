#!/bin/bash
# Comply Verify GUI launcher — double-click to start.
# Works on macOS (Finder treats .command files as executables).

set -e

# cd into the script's directory (handles spaces in path)
cd "$(dirname "$0")"

# clear screen if we have a tty (fails silently when piped/no-TERM)
clear 2>/dev/null || true

cat <<'BANNER'
╭───────────────────────────────────────────────╮
│   Comply Verify GUI — Smart Plant 1           │
│   จะเปิด browser ให้อัตโนมัติ                 │
│   ปิดหน้าต่างนี้ หรือกด Ctrl+C เพื่อหยุด     │
╰───────────────────────────────────────────────╯

BANNER

# 1. ensure Python 3
if ! command -v python3 >/dev/null 2>&1; then
  echo "❌ ไม่พบ python3 — กรุณาติดตั้ง Python 3 ก่อน (https://www.python.org)"
  read -n 1 -s -r -p "กดปุ่มใดๆ เพื่อปิด..."
  exit 1
fi

# 2. ensure dependencies (install on first run)
need_install=()
python3 -c "import flask" 2>/dev/null || need_install+=("flask")
python3 -c "import openpyxl" 2>/dev/null || need_install+=("openpyxl")
python3 -c "import fitz" 2>/dev/null || need_install+=("pymupdf")
python3 -c "import PIL" 2>/dev/null || need_install+=("pillow")

if [ ${#need_install[@]} -gt 0 ]; then
  echo "⚙  ติดตั้ง dependency ครั้งแรก: ${need_install[*]}"
  python3 -m pip install --user --quiet "${need_install[@]}" || {
    echo "❌ ติดตั้ง dependency ไม่สำเร็จ"
    read -n 1 -s -r -p "กดปุ่มใดๆ เพื่อปิด..."
    exit 1
  }
  echo "✓ ติดตั้งเรียบร้อย"
  echo
fi

# 3. kill any old instance on port 5173
old_pid=$(lsof -ti tcp:5173 2>/dev/null || true)
if [ -n "$old_pid" ]; then
  echo "🛑 พบ instance เดิมที่ port 5173 (PID $old_pid) — ปิดก่อน"
  kill "$old_pid" 2>/dev/null || true
  sleep 0.5
fi

# 4. run
echo "▶  starting…"
echo
exec python3 comply_verify_gui.py
