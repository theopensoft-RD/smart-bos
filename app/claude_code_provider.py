"""
claude_code_provider.py — Claude Code (Agent SDK) as the primary suggestion engine.

Uses the user's Claude Max subscription via OAuth (no metered API charges) and
gives Claude **full agentic tool use**: Read, Grep, Bash + custom MCP tools
that produce structured Col D / brand+model proposals.

Design (Phase 1 of the Claude-Code-as-core plan):

    SKILL.md + KB.md + learned_patterns  →  cached system prompt
    row context + few-shot               →  user message (per call)
    Allowed tools:
      Read, Grep                          →  Claude can read xlsx, PDFs (fitz),
                                            knowledge_base files, etc.
      mcp__comply__propose_col_d          →  custom MCP tool: structured proposal
      mcp__comply__propose_brand_model    →  custom MCP tool: brand+model split
      mcp__comply__escalate_to_user       →  custom MCP tool: clarifying question

The provider exposes:
  • ``propose(row_context, few_shot)``           — sync wrapper (drop-in for
                                                    AnthropicProvider.propose)
  • ``propose_streaming(row_context, few_shot)``  — async generator yielding
                                                    events for the SSE endpoint

Cost tracking still goes to ``llm_calls`` for telemetry, but ``cost_usd`` is
0 when running through Claude Max (the SDK reports ``total_cost_usd`` per
``ResultMessage`` — we surface it but don't enforce a budget cap when 0).

Auth modes:
  • Claude Max OAuth — ``claude login`` first; SDK picks up tokens from
    ``~/.claude.json``
  • API key fallback — ``ANTHROPIC_API_KEY`` env if user wants metered

Notes for future maintainers:
  • The Agent SDK is async-only. We bridge via ``asyncio.run`` for the sync
    surface and via an event queue for the streaming surface.
  • ``permission_mode="default"`` makes Claude ask before edits. We disallow
    Edit/Write entirely in Phase 1 — the *proposals* go through the user-
    confirm flow in the GUI before anything mutates xlsx/pdf.
  • Streaming uses ``ClaudeSDKClient`` for fine-grained control if we ever
    need to interrupt mid-task. For Phase 1 we use the simpler ``query()``.
"""

from __future__ import annotations

import asyncio
import json
import os
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, AsyncIterator

try:
    from claude_agent_sdk import (
        query,
        ClaudeAgentOptions,
        AssistantMessage,
        ResultMessage,
        TextBlock,
        ThinkingBlock,
        ToolUseBlock,
        ToolResultBlock,
        UserMessage,
        create_sdk_mcp_server,
        tool,
    )
    SDK_AVAILABLE = True
except ImportError as e:
    SDK_AVAILABLE = False
    _IMPORT_ERR = str(e)

from . import database as db
from . import learning


# ---------------------------------------------------------------------------
# MCP custom tools — what Claude calls to produce structured output.
# Each handler just returns the args back; the caller harvests them from the
# event stream so they're recorded as tool_use blocks.
# ---------------------------------------------------------------------------

if SDK_AVAILABLE:

    @tool(
        "propose_col_d",
        (
            "Propose Col D content for a Comply spec row. Use one of the 6 "
            "patterns from SKILL.md. ALWAYS include rationale citing which "
            "rule fired and a calibrated confidence in [0,1]. Call this exactly "
            "once per row (or call propose_brand_model / escalate_to_user)."
        ),
        {
            "pattern": str,
            "col_d_text": str,
            "col_c_proposed": str,
            "page_in_catalog": int,
            "rationale": str,
            "confidence": float,
        },
    )
    async def _propose_col_d_handler(args: dict) -> dict:
        # Just acknowledge receipt — the caller reads the args from the
        # tool_use event stream. Returning the JSON keeps Claude unblocked.
        return {
            "content": [
                {
                    "type": "text",
                    "text": "Proposal recorded. The user will review.",
                }
            ]
        }

    @tool(
        "propose_brand_model",
        (
            "For section_header rows whose Col D pattern is brand_model, "
            "decompose the catalog filename into brand + model fields per "
            "SKILL.md §Brand/Model Annotation Convention. Use 'dash brand' "
            "(brand='-') for fabricate items without a vendor brand."
        ),
        {
            "brand": str,
            "model": str,
            "is_fabricate": bool,
            "rationale": str,
            "confidence": float,
        },
    )
    async def _propose_brand_model_handler(args: dict) -> dict:
        return {
            "content": [
                {
                    "type": "text",
                    "text": "Brand+model proposal recorded.",
                }
            ]
        }

    @tool(
        "escalate_to_user",
        (
            "Use only when you genuinely cannot decide between 2+ valid "
            "interpretations OR when the row needs human-supplied info "
            "(e.g. catalog mismatch, TOR ambiguity). The system will surface "
            "your question to the user before applying anything."
        ),
        {
            "question": str,
            "options": list,
            "context": str,
        },
    )
    async def _escalate_handler(args: dict) -> dict:
        return {
            "content": [
                {
                    "type": "text",
                    "text": "Escalation noted. Awaiting user input.",
                }
            ]
        }


# ---------------------------------------------------------------------------
# Provider
# ---------------------------------------------------------------------------


class ClaudeCodeProvider:
    """Wraps the Claude Agent SDK with our domain-specific prompt structure.

    Drop-in for ``AnthropicProvider`` — exposes the same ``propose()`` method
    + adds ``propose_streaming()`` for SSE.
    """

    def __init__(
        self,
        *,
        model: str | None = None,
        max_turns: int = 8,
        skill_md_path: Path | None = None,
        kb_root: Path | None = None,
        cwd: Path | None = None,
    ):
        if not SDK_AVAILABLE:
            raise RuntimeError(
                f"claude-agent-sdk not available: {_IMPORT_ERR}. "
                "Run: pip install claude-agent-sdk (Python 3.10+ required)"
            )
        self.model = model or os.environ.get("COMPLY_LLM_MODEL", "claude-sonnet-4-5")
        self.max_turns = max_turns
        self.skill_md_path = skill_md_path
        self.kb_root = kb_root
        self.cwd = cwd or Path.cwd()
        self._cached_system_prompt: str | None = None
        self._cached_at: float = 0
        # Build the MCP server once — reused per query
        self._mcp = create_sdk_mcp_server(
            "comply",
            "1.0.0",
            [
                _propose_col_d_handler,
                _propose_brand_model_handler,
                _escalate_handler,
            ],
        )

    # --------------------------------------------------------------
    # Spend tracking (informational only when on Claude Max)
    # --------------------------------------------------------------
    def spent_today_usd(self) -> float:
        today = datetime.now(timezone.utc).strftime("%Y-%m-%d")
        try:
            with db.conn() as c:
                row = c.execute(
                    """SELECT COALESCE(SUM(cost_usd), 0) AS total
                       FROM llm_calls
                       WHERE substr(ts, 1, 10) = ?""",
                    (today,),
                ).fetchone()
                return float(row["total"] if row else 0)
        except Exception:
            return 0.0

    def budget_status(self) -> dict:
        spent = self.spent_today_usd()
        return {
            "model": self.model,
            "auth_mode": self._auth_mode(),
            "budget_usd_per_day": 0.0,  # Max = unlimited
            "spent_today_usd": round(spent, 4),
            "remaining_usd": 0.0,
            "available": SDK_AVAILABLE and self._cli_available(),
        }

    def _auth_mode(self) -> str:
        """Detect which auth path is active. Claude Max wins if both present."""
        claude_json = Path.home() / ".claude.json"
        if claude_json.exists():
            return "claude_max"
        if os.environ.get("ANTHROPIC_API_KEY"):
            return "api_key"
        return "none"

    def _cli_available(self) -> bool:
        # Quick check: claude binary on PATH?
        import shutil
        return bool(shutil.which("claude")) or self._auth_mode() != "none"

    # --------------------------------------------------------------
    # System prompt assembly (cached for 60s)
    # --------------------------------------------------------------
    def _system_prompt(self) -> str:
        if self._cached_system_prompt and (time.time() - self._cached_at < 60):
            return self._cached_system_prompt

        parts: list[str] = []
        # Header
        parts.append(
            "You are Claude Code, embedded in the Smart Plant comply-spec "
            "verification tool. You help fill in Col D (เอกสารอ้างอิง) and "
            "related fields. The user reviews every proposal you make. "
            "Always cite which SKILL.md rule fired in your rationale.\n\n"
            "## How to respond\n"
            "Call EXACTLY ONE of these MCP tools:\n"
            "  - mcp__comply__propose_col_d        (typical case)\n"
            "  - mcp__comply__propose_brand_model  (section_header brand_model rows)\n"
            "  - mcp__comply__escalate_to_user     (only if genuinely stuck)\n\n"
            "You may use Read / Grep first to inspect SKILL.md, knowledge_base/, "
            "or the row's catalog PDF (rendered text) — but conclude with a tool "
            "call, never free-form prose alone."
        )

        # SKILL.md
        if self.skill_md_path and self.skill_md_path.exists():
            try:
                txt = self.skill_md_path.read_text(encoding="utf-8")
                parts.append("## SKILL.md (project domain knowledge)\n\n" + txt)
            except Exception as e:
                sys.stderr.write(f"[claude_code_provider] SKILL.md load: {e}\n")

        # KB.md + pitfalls.md
        if self.kb_root and self.kb_root.exists():
            for fname in ("KB.md", "pitfalls.md"):
                p = self.kb_root / fname
                if p.exists():
                    try:
                        parts.append(f"## {fname}\n\n" + p.read_text(encoding="utf-8"))
                    except Exception as e:
                        sys.stderr.write(f"[claude_code_provider] {fname}: {e}\n")

        # Top learned patterns (compact)
        try:
            patterns = learning.list_patterns()
            top = sorted(patterns, key=lambda p: (-p.get("confidence", 0),
                                                   -p.get("samples_total", 0)))[:30]
            if top:
                lines = ["## Top learned patterns (user-validated, high priority)"]
                for p in top:
                    lines.append(
                        f"- type={p['type']} trigger={p['trigger_key']} → "
                        f"output={p['output']!r} (n={p['samples_total']}, "
                        f"conf={p['confidence']:.2f})"
                    )
                parts.append("\n".join(lines))
        except Exception as e:
            sys.stderr.write(f"[claude_code_provider] patterns: {e}\n")

        self._cached_system_prompt = "\n\n".join(parts)
        self._cached_at = time.time()
        return self._cached_system_prompt

    # --------------------------------------------------------------
    # User message builder (shared by sync + streaming)
    # --------------------------------------------------------------
    def _build_user_msg(self, row_context: dict, few_shot: list[dict] | None) -> str:
        msg_parts: list[str] = []
        msg_parts.append("# Row to analyze\n")
        msg_parts.append(f"- row: R{row_context.get('row', '?')}")
        msg_parts.append(f"- section: {row_context.get('section', '?')}")
        msg_parts.append(f"- role: {row_context.get('role', 'unknown')}")
        if row_context.get("col_a"):
            msg_parts.append(f"- Col A (section ID): {row_context['col_a']}")
        msg_parts.append(f"- Col B (TOR text): {row_context.get('col_b', '')!r}")
        if row_context.get("col_c_current"):
            msg_parts.append(f"- Col C current: {row_context['col_c_current']!r}")
        if row_context.get("col_d_current"):
            msg_parts.append(f"- Col D current: {row_context['col_d_current']!r}")
        if row_context.get("col_e"):
            msg_parts.append(f"- Col E (vendor): {row_context['col_e']}")
        if row_context.get("pdf_filename"):
            msg_parts.append(f"- catalog PDF filename: {row_context['pdf_filename']!r}")
        if row_context.get("pdf_rel"):
            msg_parts.append(f"- catalog PDF rel path: {row_context['pdf_rel']!r}")
        if row_context.get("tor_excerpt"):
            tor = row_context["tor_excerpt"][:1000]
            msg_parts.append(f"\n## TOR excerpt\n{tor}")
        if row_context.get("rule_proposal"):
            rp = row_context["rule_proposal"]
            msg_parts.append(
                f"\n## Rule-based proposal (for your consideration):\n"
                f"- proposed_d: {rp.get('proposed_d', '')!r}\n"
                f"- generator: {rp.get('generator', '?')}\n"
                f"- confidence: {rp.get('confidence', 0):.2f}\n"
                f"You may agree, refine, or override. Cite SKILL.md."
            )

        if few_shot:
            msg_parts.append(f"\n## {len(few_shot)} relevant past corrections (for in-context learning):")
            for i, fs in enumerate(few_shot[:5], 1):
                msg_parts.append(
                    f"\n### Example {i}\n"
                    f"- input Col B: {(fs.get('input_b') or '')[:200]!r}\n"
                    f"- final Col D: {(fs.get('final_d') or '')[:200]!r}\n"
                    f"- correction kind: {fs.get('correction_kind', 'unknown')}"
                )

        msg_parts.append(
            "\n## Task\n"
            "Inspect the context above. Use Read/Grep on SKILL.md or the "
            "catalog PDF if helpful. Then call exactly ONE MCP tool:\n"
            "  - propose_col_d (typical)\n"
            "  - propose_brand_model (section_header brand_model rows)\n"
            "  - escalate_to_user (cannot decide)"
        )
        return "\n".join(msg_parts)

    def _build_options(self) -> "ClaudeAgentOptions":
        return ClaudeAgentOptions(
            system_prompt=self._system_prompt(),
            mcp_servers={"comply": self._mcp},
            allowed_tools=[
                "mcp__comply__propose_col_d",
                "mcp__comply__propose_brand_model",
                "mcp__comply__escalate_to_user",
                "Read",
                "Grep",
            ],
            cwd=str(self.cwd),
            permission_mode="default",  # ask before any unexpected tool
            max_turns=self.max_turns,
            model=self.model,
        )

    # --------------------------------------------------------------
    # Sync proposal (drop-in for AnthropicProvider.propose)
    # --------------------------------------------------------------
    def propose(self, *, row_context: dict, few_shot: list[dict] | None = None) -> dict:
        """Sync wrapper around the async query. Same return shape as
        AnthropicProvider.propose so the learning bridge is identical."""
        try:
            return asyncio.run(self._propose_async(row_context, few_shot))
        except Exception as e:
            sys.stderr.write(f"[claude_code_provider] propose failed: {e}\n")
            return {"ok": False, "error": str(e), "model": self.model, "elapsed_ms": 0}

    async def _propose_async(self, row_context: dict, few_shot: list[dict] | None) -> dict:
        user_msg = self._build_user_msg(row_context, few_shot)
        opts = self._build_options()

        tool_calls: list[dict] = []
        text_parts: list[str] = []
        thinking_parts: list[str] = []
        result_msg: ResultMessage | None = None
        t0 = time.time()

        async for event in query(prompt=user_msg, options=opts):
            if isinstance(event, AssistantMessage):
                for blk in event.content:
                    if isinstance(blk, TextBlock):
                        text_parts.append(blk.text)
                    elif isinstance(blk, ThinkingBlock):
                        thinking_parts.append(blk.thinking)
                    elif isinstance(blk, ToolUseBlock):
                        # Only record our MCP tools as "the proposal"
                        if blk.name.startswith("mcp__comply__"):
                            tool_calls.append({
                                "name": blk.name.replace("mcp__comply__", ""),
                                "input": blk.input,
                            })
            elif isinstance(event, ResultMessage):
                result_msg = event

        elapsed_ms = int((time.time() - t0) * 1000)
        usage_dict = (result_msg.usage if result_msg else {}) or {}
        cost_usd = float(result_msg.total_cost_usd or 0.0) if result_msg else 0.0
        stop_reason = (result_msg.stop_reason or result_msg.subtype if result_msg else "") or ""

        # Persist call for telemetry
        try:
            self._record_call(
                row_num=row_context.get("row"),
                model=self.model,
                stop_reason=stop_reason,
                usage=usage_dict,
                cost_usd=cost_usd,
                elapsed_ms=elapsed_ms,
                tool_calls=tool_calls,
                text="".join(text_parts),
                user_msg=user_msg,
            )
        except Exception as e:
            sys.stderr.write(f"[claude_code_provider] _record_call: {e}\n")

        return {
            "ok": True,
            "tool_calls": tool_calls,
            "text": "".join(text_parts),
            "thinking": "".join(thinking_parts),
            "usage": usage_dict,
            "cost_usd": cost_usd,
            "model": self.model,
            "stop_reason": stop_reason,
            "elapsed_ms": elapsed_ms,
        }

    # --------------------------------------------------------------
    # Streaming proposal (for SSE — frontend renders Claude's reasoning live)
    # --------------------------------------------------------------
    async def propose_streaming(
        self,
        *,
        row_context: dict,
        few_shot: list[dict] | None = None,
    ) -> AsyncIterator[dict]:
        """Yield events as Claude works:
            {type: "thinking", text: "..."}
            {type: "tool_use",  name: "Read"|"Grep"|"propose_col_d"|..., input: {...}}
            {type: "tool_result", name: "...", text: "..." (truncated)}
            {type: "text",      content: "..."}    # any narration
            {type: "result",    proposal: {...}, usage: {...}, cost_usd: 0, elapsed_ms: ...}
            {type: "error",     error: "..."}
        """
        user_msg = self._build_user_msg(row_context, few_shot)
        opts = self._build_options()

        tool_calls: list[dict] = []
        text_parts: list[str] = []
        thinking_parts: list[str] = []
        # Map tool_use_id → tool name for matching tool_result blocks
        tool_id_to_name: dict[str, str] = {}
        t0 = time.time()
        result_msg: ResultMessage | None = None

        try:
            async for event in query(prompt=user_msg, options=opts):
                if isinstance(event, AssistantMessage):
                    for blk in event.content:
                        if isinstance(blk, TextBlock):
                            text_parts.append(blk.text)
                            yield {"type": "text", "content": blk.text}
                        elif isinstance(blk, ThinkingBlock):
                            thinking_parts.append(blk.thinking)
                            yield {"type": "thinking", "text": blk.thinking}
                        elif isinstance(blk, ToolUseBlock):
                            tool_id_to_name[blk.id] = blk.name
                            short_name = blk.name.replace("mcp__comply__", "")
                            yield {
                                "type": "tool_use",
                                "name": short_name,
                                "raw_name": blk.name,
                                "input": blk.input,
                            }
                            if blk.name.startswith("mcp__comply__"):
                                tool_calls.append({
                                    "name": short_name,
                                    "input": blk.input,
                                })
                elif isinstance(event, UserMessage):
                    # Tool results come back as user messages with ToolResultBlock
                    content = event.content if hasattr(event, "content") else []
                    if isinstance(content, list):
                        for blk in content:
                            if isinstance(blk, ToolResultBlock):
                                name = tool_id_to_name.get(blk.tool_use_id, "tool")
                                text = ""
                                if isinstance(blk.content, str):
                                    text = blk.content
                                elif isinstance(blk.content, list):
                                    text = "".join(
                                        c.get("text", "") for c in blk.content
                                        if isinstance(c, dict)
                                    )
                                yield {
                                    "type": "tool_result",
                                    "name": name.replace("mcp__comply__", ""),
                                    "text": (text or "")[:400],
                                    "is_error": bool(blk.is_error),
                                }
                elif isinstance(event, ResultMessage):
                    result_msg = event
        except Exception as e:
            yield {"type": "error", "error": str(e)}
            return

        elapsed_ms = int((time.time() - t0) * 1000)
        usage_dict = (result_msg.usage if result_msg else {}) or {}
        cost_usd = float(result_msg.total_cost_usd or 0.0) if result_msg else 0.0
        stop_reason = (result_msg.stop_reason or result_msg.subtype if result_msg else "") or ""

        # Pick the proposal — prefer the LAST mcp__comply__ tool call
        proposal = tool_calls[-1] if tool_calls else None

        # Persist
        try:
            self._record_call(
                row_num=row_context.get("row"),
                model=self.model,
                stop_reason=stop_reason,
                usage=usage_dict,
                cost_usd=cost_usd,
                elapsed_ms=elapsed_ms,
                tool_calls=tool_calls,
                text="".join(text_parts),
                user_msg=user_msg,
            )
        except Exception as e:
            sys.stderr.write(f"[claude_code_provider] _record_call: {e}\n")

        yield {
            "type": "result",
            "proposal": proposal,
            "all_tool_calls": tool_calls,
            "text": "".join(text_parts),
            "thinking": "".join(thinking_parts),
            "usage": usage_dict,
            "cost_usd": cost_usd,
            "model": self.model,
            "stop_reason": stop_reason,
            "elapsed_ms": elapsed_ms,
        }

    # --------------------------------------------------------------
    # Audit/persistence
    # --------------------------------------------------------------
    def _record_call(self, *, row_num, model, stop_reason, usage, cost_usd,
                      elapsed_ms, tool_calls, text, user_msg):
        with db.conn() as c:
            c.execute(
                """INSERT INTO llm_calls
                   (ts, row_num, model, stop_reason, input_tokens, output_tokens,
                    cache_write_tokens, cache_read_tokens, cost_usd, elapsed_ms,
                    tool_calls_json, response_text, prompt_size_chars)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (
                    datetime.now(timezone.utc).isoformat(timespec="seconds"),
                    row_num, model, stop_reason,
                    int(usage.get("input_tokens", 0) or 0),
                    int(usage.get("output_tokens", 0) or 0),
                    int(usage.get("cache_creation_input_tokens", 0) or 0),
                    int(usage.get("cache_read_input_tokens", 0) or 0),
                    cost_usd, elapsed_ms,
                    json.dumps(tool_calls, ensure_ascii=False),
                    (text or "")[:4000],
                    len(user_msg),
                ),
            )


# ---------------------------------------------------------------------------
# Module-level singleton + bootstrap
# ---------------------------------------------------------------------------

_provider: ClaudeCodeProvider | None = None


def get_provider() -> ClaudeCodeProvider | None:
    """Return the singleton provider, lazily. Returns None when SDK or CLI
    are not installed."""
    global _provider
    if _provider is not None:
        return _provider
    if not SDK_AVAILABLE:
        return None
    # User must opt in via env var (preserves current default-off behaviour)
    if (os.environ.get("COMPLY_LLM") or "").lower() not in ("claude_code", "claude-code"):
        return None
    try:
        root = Path(__file__).parent.parent
        _provider = ClaudeCodeProvider(
            skill_md_path=root / "SKILL.md",
            kb_root=root / "knowledge_base",
            cwd=root,
        )
    except Exception as e:
        sys.stderr.write(f"[claude_code_provider] init failed: {e}\n")
        return None
    return _provider


def install_into_learning() -> bool:
    """Wire this provider into ``app.learning`` so the existing
    ``auto_annotate_plan`` low-confidence path routes through Claude."""
    p = get_provider()
    if p is None:
        return False

    def _bridge(prompt: str, context: dict) -> dict:
        try:
            return p.propose(row_context=context, few_shot=context.get("_few_shot"))
        except Exception as e:
            return {"ok": False, "error": str(e)}

    learning.set_llm_provider(_bridge, name=f"claude_code:{p.model}")
    return True
