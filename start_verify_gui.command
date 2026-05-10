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

# 1. ensure uv (https://docs.astral.sh/uv/) — fast Python package manager
#    + auto-manages Python 3.10+ for us via `uv python install`
if ! command -v uv >/dev/null 2>&1; then
  echo "ℹ uv (Python package manager) ยังไม่ถูกติดตั้ง"
  if command -v brew >/dev/null 2>&1; then
    echo "   ติดตั้งผ่าน Homebrew: brew install uv"
    brew install uv || {
      echo "❌ brew install uv ล้มเหลว"
      read -n 1 -s -r -p "กดปุ่มใดๆ เพื่อปิด..."
      exit 1
    }
  else
    echo "   ติดตั้งผ่าน: curl -LsSf https://astral.sh/uv/install.sh | sh"
    curl -LsSf https://astral.sh/uv/install.sh | sh || {
      echo "❌ ติดตั้ง uv ล้มเหลว"
      read -n 1 -s -r -p "กดปุ่มใดๆ เพื่อปิด..."
      exit 1
    }
    export PATH="$HOME/.local/bin:$PATH"
  fi
fi

# 2. let uv resolve/install dependencies from pyproject.toml. First run
#    creates .venv with Python 3.11 + all deps. Subsequent runs are ~0.1s.
echo "⚙  uv sync (รวดเร็ว — ใช้ pyproject.toml + lockfile)…"
if ! uv sync --quiet 2>&1; then
  echo "❌ uv sync ล้มเหลว"
  read -n 1 -s -r -p "กดปุ่มใดๆ เพื่อปิด..."
  exit 1
fi
PY=".venv/bin/python"
echo "▶  ใช้ $PY ($($PY --version))"

# 2b. ensure Claude Code CLI (npm package). Optional — only needed if
# user wants Claude Max OAuth path. Skip silently if npm not present.
if command -v npm >/dev/null 2>&1; then
  if ! command -v claude >/dev/null 2>&1; then
    echo "ℹ Claude Code CLI ยังไม่ถูกติดตั้ง — ติดตั้งเพื่อใช้ Claude Max?"
    echo "   (Enter เพื่อข้าม / 'y' เพื่อติดตั้งผ่าน npm -g)"
    read -t 5 -r -n 1 ans || ans=""
    echo
    if [ "$ans" = "y" ] || [ "$ans" = "Y" ]; then
      npm install -g @anthropic-ai/claude-code || echo "⚠ ติดตั้ง claude CLI ไม่สำเร็จ — ระบบจะใช้ rules-only mode"
    fi
  fi
  # Check auth status if CLI present
  if command -v claude >/dev/null 2>&1; then
    auth_status=$(claude auth status 2>/dev/null | grep -o '"loggedIn":[^,]*' | head -1)
    if echo "$auth_status" | grep -q "false"; then
      echo
      echo "⚠ Claude Code CLI ยังไม่ได้ login"
      echo "   รัน 'claude auth login' ใน terminal อีกบานเพื่อใช้ Claude Max"
      echo "   (กด Enter เพื่อข้าม — จะใช้ rules-only mode)"
      read -t 3 -r -n 1 -s _ || true
      echo
    fi
  fi
fi

# 3. kill any old instance on port 5173
old_pid=$(lsof -ti tcp:5173 2>/dev/null || true)
if [ -n "$old_pid" ]; then
  echo "🛑 พบ instance เดิมที่ port 5173 (PID $old_pid) — ปิดก่อน"
  kill "$old_pid" 2>/dev/null || true
  sleep 0.5
fi

# 4. run — default to claude_code provider (Phase 1)
export COMPLY_LLM=${COMPLY_LLM:-claude_code}
echo "▶  starting… (LLM=$COMPLY_LLM)"
echo
exec "$PY" comply_verify_gui.py
