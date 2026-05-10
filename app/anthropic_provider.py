"""
anthropic_provider.py — Claude as the primary suggestion engine.

Hooks into ``app.learning.set_llm_provider`` so the existing
``auto_annotate_plan`` pipeline can route through Claude. Designed for
the Smart Plant comply-spec workflow:

    SKILL.md (cached)        →  system block #1, ephemeral cache
    KB.md + pitfalls.md      →  system block #2, ephemeral cache
    learned_patterns top-N    →  system block #3, ephemeral cache
    row context + few-shot    →  user message (per call)

Claude calls one of three tools to produce structured output:

    propose_col_d              →  the Col D string + pattern + confidence
    propose_brand_model        →  brand + model decomposition
    escalate_to_user           →  surface a clarifying question

The provider tracks tokens and enforces a per-day spend cap.

Design choices:
  • Sonnet 4.5 default (cheap + fast, good rule-follower)
  • Vision OFF for now — page resolution still done by find_text_match_in_pdf;
    we ship text excerpts + filename, not page images
  • Extended thinking OFF for now — re-enable per-call when confidence < 0.6
  • Prompt cache: the 3 system blocks tagged ephemeral (5-min TTL)
  • Budget cap: USD/day, enforced server-side, returns {ok: False, error: 'budget_exceeded'}
"""

from __future__ import annotations

import json
import os
import re
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

try:
    import anthropic
except ImportError:
    anthropic = None  # type: ignore

from . import database as db
from . import learning

# ---------------------------------------------------------------------------
# Pricing — Anthropic public list rates (USD per million tokens), as of
# 2025-Q4. Used for budget enforcement only — billing is whatever the
# Anthropic invoice says. Keep this current when models change.
# ---------------------------------------------------------------------------
_PRICES_PER_M_TOKENS = {
    # model_id_prefix : (input, output, cache_write, cache_read)
    "claude-sonnet-4-5":   (3.00, 15.00, 3.75, 0.30),
    "claude-opus-4-5":    (15.00, 75.00, 18.75, 1.50),
    "claude-haiku-4-5":   (0.80,  4.00, 1.00, 0.08),
    # Aliases that resolve to the latest of each family
    "claude-sonnet":       (3.00, 15.00, 3.75, 0.30),
    "claude-opus":        (15.00, 75.00, 18.75, 1.50),
    "claude-haiku":       (0.80,  4.00, 1.00, 0.08),
}


def _price_for(model: str) -> tuple[float, float, float, float]:
    """Return (input, output, cache_write, cache_read) USD per million tokens."""
    for prefix, prices in _PRICES_PER_M_TOKENS.items():
        if model.startswith(prefix):
            return prices
    # Default to Sonnet pricing if unknown — conservative
    return _PRICES_PER_M_TOKENS["claude-sonnet-4-5"]


def estimate_cost_usd(model: str, input_tokens: int, output_tokens: int,
                       cache_write_tokens: int = 0,
                       cache_read_tokens: int = 0) -> float:
    """USD estimate for one API call. Cache reads are 90% off vs uncached input."""
    p_in, p_out, p_cw, p_cr = _price_for(model)
    return (
        input_tokens / 1_000_000 * p_in +
        output_tokens / 1_000_000 * p_out +
        cache_write_tokens / 1_000_000 * p_cw +
        cache_read_tokens / 1_000_000 * p_cr
    )


# ---------------------------------------------------------------------------
# Tool definitions — what Claude can return as structured output.
# ---------------------------------------------------------------------------

TOOLS: list[dict] = [
    {
        "name": "propose_col_d",
        "description": (
            "Propose Col D content for a Comply spec row. Use one of the 6 "
            "patterns from SKILL.md. ALWAYS include rationale citing which "
            "rule fired and a calibrated confidence in [0,1]."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "pattern": {
                    "type": "string",
                    "enum": ["brand_model", "equivalent", "higher",
                             "commitment", "filename_format", "empty",
                             "model_only"],
                    "description": "Which Col D pattern from SKILL.md."
                },
                "col_d_text": {
                    "type": "string",
                    "description": "The exact Col D string to write to xlsx."
                },
                "col_c_proposed": {
                    "type": "string",
                    "description": (
                        "Optional: a refined Col C value if the rule-based "
                        "make_col_c_from_b output is wrong. Leave empty to keep rules' output."
                    )
                },
                "page_in_catalog": {
                    "type": "integer",
                    "description": "Catalog PDF page number this row references (1-indexed). 0 if N/A."
                },
                "rationale": {
                    "type": "string",
                    "description": (
                        "1-2 sentences citing the SKILL.md rule that produced this. "
                        "If unsure, say so explicitly."
                    )
                },
                "confidence": {
                    "type": "number",
                    "minimum": 0.0, "maximum": 1.0,
                    "description": (
                        "Calibrated confidence: 0.9+ when rule clearly matches, "
                        "0.5-0.8 when plausible but ambiguous, <0.5 when guessing."
                    )
                }
            },
            "required": ["pattern", "col_d_text", "rationale", "confidence"]
        }
    },
    {
        "name": "propose_brand_model",
        "description": (
            "For section_header rows whose Col D pattern is brand_model, "
            "decompose the catalog filename into brand + model fields per SKILL.md "
            "§Brand/Model Annotation Convention. Use 'dash brand' (brand='-') for "
            "fabricate items without a vendor brand."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "brand": {
                    "type": "string",
                    "description": "Vendor brand name. Use '-' for fabricate items."
                },
                "model": {
                    "type": "string",
                    "description": "Full model number/name."
                },
                "is_fabricate": {
                    "type": "boolean",
                    "description": "True for งาน fabricate (เสา, custom items)."
                },
                "rationale": {"type": "string"},
                "confidence": {"type": "number", "minimum": 0.0, "maximum": 1.0}
            },
            "required": ["brand", "model", "rationale", "confidence"]
        }
    },
    {
        "name": "escalate_to_user",
        "description": (
            "Use only when you genuinely cannot decide between 2+ valid "
            "interpretations OR when the row needs human-supplied info "
            "(e.g. catalog mismatch, TOR ambiguity). The system will block "
            "the apply step and surface your question to the user."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "question": {"type": "string"},
                "options": {
                    "type": "array", "items": {"type": "string"},
                    "description": "If choice is between fixed alternatives, list them."
                },
                "context": {"type": "string", "description": "Why you're escalating."}
            },
            "required": ["question", "context"]
        }
    },
]


# ---------------------------------------------------------------------------
# Provider
# ---------------------------------------------------------------------------

class AnthropicProvider:
    """Wraps the Anthropic SDK with our domain-specific prompt structure,
    prompt caching, budget enforcement, and audit hooks."""

    def __init__(
        self,
        *,
        model: str | None = None,
        budget_usd_per_day: float | None = None,
        max_tokens_response: int = 1024,
        skill_md_path: Path | None = None,
        kb_root: Path | None = None,
    ):
        if anthropic is None:
            raise RuntimeError("anthropic SDK not installed. pip install anthropic")
        api_key = os.environ.get("ANTHROPIC_API_KEY", "").strip()
        if not api_key:
            raise RuntimeError("ANTHROPIC_API_KEY not set in environment")
        self.client = anthropic.Anthropic(api_key=api_key)
        self.model = model or os.environ.get("COMPLY_LLM_MODEL", "claude-sonnet-4-5")
        self.budget = float(
            budget_usd_per_day if budget_usd_per_day is not None
            else os.environ.get("COMPLY_LLM_BUDGET_USD_PER_DAY", "5.00")
        )
        self.max_tokens = max_tokens_response
        self.skill_md_path = skill_md_path
        self.kb_root = kb_root
        self._cached_system_blocks: list[dict] | None = None
        self._cached_system_loaded_at: float = 0
        self._spent_today_cache: tuple[str, float] | None = None  # (date, usd)

    # --------------------------------------------------------------
    # Spend tracking
    # --------------------------------------------------------------
    def spent_today_usd(self) -> float:
        """Sum cost_usd across today's rows in learning_feedback (UTC).
        Cached for 30 sec to avoid hammering the DB."""
        today = datetime.now(timezone.utc).strftime("%Y-%m-%d")
        if self._spent_today_cache and self._spent_today_cache[0] == today:
            ts_age = time.time() - self._spent_today_cache[2] if len(self._spent_today_cache) > 2 else 999  # type: ignore
        cur_date = today
        try:
            with db.conn() as c:
                # Sum cost_usd from llm_calls table; if column doesn't exist, return 0
                row = c.execute(
                    """SELECT COALESCE(SUM(cost_usd), 0) AS total
                       FROM llm_calls
                       WHERE substr(ts, 1, 10) = ?""",
                    (cur_date,),
                ).fetchone()
                return float(row["total"] if row else 0)
        except Exception:
            return 0.0

    def budget_status(self) -> dict:
        spent = self.spent_today_usd()
        return {
            "model": self.model,
            "budget_usd_per_day": self.budget,
            "spent_today_usd": round(spent, 4),
            "remaining_usd": round(max(0.0, self.budget - spent), 4),
            "available": (anthropic is not None) and bool(os.environ.get("ANTHROPIC_API_KEY")),
        }

    def _check_budget(self) -> None:
        if self.budget <= 0:
            return  # 0 = unlimited
        spent = self.spent_today_usd()
        if spent >= self.budget:
            raise BudgetExceededError(
                f"Daily budget exceeded: spent ${spent:.2f} / ${self.budget:.2f}"
            )

    # --------------------------------------------------------------
    # System prompt assembly (cached)
    # --------------------------------------------------------------
    def _system_blocks(self) -> list[dict]:
        """Build the cached system blocks. Each chunk gets cache_control:
        ephemeral so a 5-min TTL applies."""
        # Cache the assembled blocks for 60 sec — only re-read SKILL.md if it
        # changed on disk (mtime check would be nice, skipping for simplicity)
        if self._cached_system_blocks and (time.time() - self._cached_system_loaded_at < 60):
            return self._cached_system_blocks

        blocks: list[dict] = []
        # Block 0 — role + project context (small, leading)
        blocks.append({
            "type": "text",
            "text": (
                "You are an AI assistant for the Smart Plant comply-spec verification "
                "tool. You help fill in Col D (เอกสารอ้างอิง) and other fields by "
                "reading the row's TOR text (Col B), the catalog PDF reference, and "
                "domain rules from SKILL.md. The user reviews every proposal. You "
                "produce structured output via tool calls only — never free-form prose. "
                "Always cite which SKILL.md rule fired in your rationale."
            ),
        })

        # Block 1 — SKILL.md (the canonical domain knowledge, 63 KB)
        if self.skill_md_path and self.skill_md_path.exists():
            try:
                txt = self.skill_md_path.read_text(encoding="utf-8")
                blocks.append({
                    "type": "text",
                    "text": "## SKILL.md (project domain knowledge)\n\n" + txt,
                    "cache_control": {"type": "ephemeral"},
                })
            except Exception as e:
                sys.stderr.write(f"[anthropic_provider] failed to load SKILL.md: {e}\n")

        # Block 2 — knowledge_base/{KB.md, pitfalls.md}
        if self.kb_root and self.kb_root.exists():
            kb_text_parts: list[str] = []
            for fname in ("KB.md", "pitfalls.md"):
                p = self.kb_root / fname
                if p.exists():
                    try:
                        kb_text_parts.append(f"## {fname}\n\n" + p.read_text(encoding="utf-8"))
                    except Exception as e:
                        sys.stderr.write(f"[anthropic_provider] failed to load {fname}: {e}\n")
            if kb_text_parts:
                blocks.append({
                    "type": "text",
                    "text": "\n\n---\n\n".join(kb_text_parts),
                    "cache_control": {"type": "ephemeral"},
                })

        # Block 3 — top learned patterns (compact)
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
                blocks.append({
                    "type": "text",
                    "text": "\n".join(lines),
                    "cache_control": {"type": "ephemeral"},
                })
        except Exception as e:
            sys.stderr.write(f"[anthropic_provider] learned_patterns load failed: {e}\n")

        self._cached_system_blocks = blocks
        self._cached_system_loaded_at = time.time()
        return blocks

    # --------------------------------------------------------------
    # Single-row proposal
    # --------------------------------------------------------------
    def propose(self, *, row_context: dict, few_shot: list[dict] | None = None) -> dict:
        """Ask Claude to propose Col D / brand+model / escalation for one row.

        row_context: {
            'row': int, 'section': str, 'role': str,
            'col_a': str, 'col_b': str, 'col_c_current': str,
            'col_d_current': str, 'col_e': str,
            'pdf_rel': str | None, 'pdf_filename': str | None,
            'tor_excerpt': str | None,
            'rule_proposal': dict | None,  # what the rule-based generator said
        }
        few_shot: optional list of {input_b, final_d, generator, ...} from
            past learning_feedback rows — used as in-context examples.

        Returns: {
            'ok': True,
            'tool_calls': [...],   # parsed tool inputs
            'text': str,            # any prose Claude included
            'usage': {input_tokens, output_tokens, cache_read_tokens, ...},
            'cost_usd': float,
            'model': str,
            'stop_reason': str,
        }
        """
        self._check_budget()

        # Build the user message — compact, structured
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

        # Few-shot examples from past corrections
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
            "Call exactly ONE tool:\n"
            "- propose_col_d for the typical case\n"
            "- propose_brand_model for section_header rows whose pattern is brand_model "
            "(in addition to or instead of propose_col_d when the user only needs the brand split)\n"
            "- escalate_to_user only if you cannot decide and need clarification"
        )

        user_msg = "\n".join(msg_parts)

        # API call
        try:
            t0 = time.time()
            resp = self.client.messages.create(
                model=self.model,
                max_tokens=self.max_tokens,
                system=self._system_blocks(),
                tools=TOOLS,
                messages=[{"role": "user", "content": user_msg}],
            )
            elapsed_ms = int((time.time() - t0) * 1000)
        except anthropic.RateLimitError as e:
            return {"ok": False, "error": f"rate_limited: {e}",
                    "model": self.model, "elapsed_ms": 0}
        except anthropic.APIError as e:
            return {"ok": False, "error": f"api_error: {e}",
                    "model": self.model, "elapsed_ms": 0}
        except Exception as e:
            return {"ok": False, "error": f"unexpected: {e}",
                    "model": self.model, "elapsed_ms": 0}

        # Parse usage
        u = resp.usage
        usage = {
            "input_tokens": getattr(u, "input_tokens", 0),
            "output_tokens": getattr(u, "output_tokens", 0),
            "cache_creation_input_tokens": getattr(u, "cache_creation_input_tokens", 0) or 0,
            "cache_read_input_tokens": getattr(u, "cache_read_input_tokens", 0) or 0,
        }
        cost_usd = estimate_cost_usd(
            self.model,
            usage["input_tokens"], usage["output_tokens"],
            usage["cache_creation_input_tokens"], usage["cache_read_input_tokens"],
        )

        # Parse content blocks
        tool_calls: list[dict] = []
        text_parts: list[str] = []
        for blk in resp.content:
            if blk.type == "text":
                text_parts.append(blk.text)
            elif blk.type == "tool_use":
                tool_calls.append({
                    "name": blk.name,
                    "input": blk.input,
                })

        # Persist the call to llm_calls for budget tracking + audit
        try:
            self._record_call(
                row_num=row_context.get("row"),
                model=resp.model,
                stop_reason=resp.stop_reason or "",
                usage=usage,
                cost_usd=cost_usd,
                elapsed_ms=elapsed_ms,
                tool_calls=tool_calls,
                text="".join(text_parts),
                user_msg=user_msg,
            )
        except Exception as e:
            sys.stderr.write(f"[anthropic_provider] _record_call failed: {e}\n")

        return {
            "ok": True,
            "tool_calls": tool_calls,
            "text": "".join(text_parts),
            "usage": usage,
            "cost_usd": cost_usd,
            "model": resp.model,
            "stop_reason": resp.stop_reason,
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
                    usage["input_tokens"], usage["output_tokens"],
                    usage["cache_creation_input_tokens"], usage["cache_read_input_tokens"],
                    cost_usd, elapsed_ms,
                    json.dumps(tool_calls, ensure_ascii=False),
                    (text or "")[:4000],
                    len(user_msg),
                ),
            )


class BudgetExceededError(RuntimeError):
    """Raised when the per-day budget cap is reached."""
    pass


# ---------------------------------------------------------------------------
# Module-level singleton + bootstrap
# ---------------------------------------------------------------------------

_provider: AnthropicProvider | None = None


def get_provider() -> AnthropicProvider | None:
    """Return the singleton provider, creating it lazily. Returns None when
    LLM is disabled or misconfigured."""
    global _provider
    if _provider is not None:
        return _provider
    if not os.environ.get("ANTHROPIC_API_KEY"):
        return None
    if (os.environ.get("COMPLY_LLM") or "").lower() != "anthropic":
        return None
    if anthropic is None:
        sys.stderr.write("[anthropic_provider] anthropic SDK not installed\n")
        return None
    try:
        # Resolve paths from project root (anthropic_provider lives in app/)
        root = Path(__file__).parent.parent
        _provider = AnthropicProvider(
            skill_md_path=root / "SKILL.md",
            kb_root=root / "knowledge_base",
        )
    except Exception as e:
        sys.stderr.write(f"[anthropic_provider] init failed: {e}\n")
        return None
    return _provider


def install_into_learning() -> bool:
    """Wire this provider into app.learning so auto_annotate_plan can use it.
    Idempotent — returns True if a provider is now installed."""
    p = get_provider()
    if p is None:
        return False

    def _bridge(prompt: str, context: dict) -> dict:
        # Old LLMProvider signature was (prompt:str, context:dict)→dict.
        # We adapt by treating `context` as our row_context shape.
        try:
            r = p.propose(row_context=context, few_shot=context.get("_few_shot"))
            return r
        except BudgetExceededError as e:
            return {"ok": False, "error": str(e), "budget_exceeded": True}

    learning.set_llm_provider(_bridge, name=p.model)
    return True
