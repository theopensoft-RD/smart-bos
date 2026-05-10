"""
comply_learning.py — HITL (Human-in-the-Loop) learning core.

The verify GUI is built around two contracts:

  1. **Core proposes, human verifies visually.**  For every row that needs
     Col C/D + PDF annotations, the core produces a *suggestion* with a
     confidence score and a provenance trail (which rule / which past
     correction fed it).  The user only does visual proof-checking.

  2. **Every correction feeds back.**  Whether the user accepts the
     suggestion, edits it, or rejects it, the (input → suggestion → final)
     triple is appended to ``learning_feedback``.  A periodic
     ``retrain()`` step distills repeating corrections into
     ``learned_patterns`` that override the rule-based generator on next
     run.

This module owns:

  • the suggestion pipeline (rule + learned pattern + optional LLM)
  • the feedback recorder
  • the retrain / pattern-mining routine
  • a pluggable LLM provider hook (off by default; switch on with the
    ``COMPLY_LLM`` env var or `set_llm_provider()`)

It deliberately treats the rules in ``comply_verify_gui.py`` as a fallback
generator — the *learned patterns* take priority because they're
user-validated, while rules cover the long tail.
"""

from __future__ import annotations

import json
import os
import re
import sys
import time
from collections.abc import Callable
from datetime import datetime

from . import database as db

# ---------------------------------------------------------------------------
# LLM provider hook (pluggable, off by default)
# ---------------------------------------------------------------------------

# Signature: provider(prompt: str, context: dict) -> dict
#   context contains: row_num, section, input_b, pdf_rel, suggested_d (rule)
#   should return: {col_d?, col_c?, annotations?, confidence, rationale}
LLMProvider = Callable[[str, dict], dict]
_LLM_PROVIDER: LLMProvider | None = None
_LLM_NAME: str = "off"


def set_llm_provider(provider: LLMProvider | None, name: str = "custom") -> None:
    """Plug in an LLM (Anthropic, OpenAI, Ollama, ...).  When set, novel
    rows (low-confidence rule output, no learned pattern) are routed
    through the provider before falling back to rules."""
    global _LLM_PROVIDER, _LLM_NAME
    _LLM_PROVIDER = provider
    _LLM_NAME = name if provider is not None else "off"


def llm_status() -> dict:
    return {"available": _LLM_PROVIDER is not None, "name": _LLM_NAME}


# ---------------------------------------------------------------------------
# Pattern extractors used by retrain()
# ---------------------------------------------------------------------------

_FILENAME_BRAND_TOKEN_RE = re.compile(r"[A-Za-z][A-Za-z0-9]{2,}")


def _last_latin_token(filename_stem: str) -> str | None:
    """Naive brand candidate from the END of a filename (works for
    "...Lenovo ThinkSystem SR630 V4")."""
    tokens = _FILENAME_BRAND_TOKEN_RE.findall(filename_stem or "")
    return tokens[0] if tokens else None


def _section_root(section: str | None) -> str | None:
    if not section:
        return None
    parts = section.split(".")
    return ".".join(parts[:2]) if len(parts) >= 2 else section


# ---------------------------------------------------------------------------
# Recording feedback (called whenever the user accepts/edits/rejects)
# ---------------------------------------------------------------------------

def _classify_correction(suggested_d: str, final_d: str) -> str:
    """Coarse categorisation of what the user fixed — used by the retrain
    pass to decide which pattern types to update."""
    if not suggested_d and final_d:
        return "filled_empty"
    if suggested_d and not final_d:
        return "cleared"
    if suggested_d == final_d:
        return "no_change"
    sd = (suggested_d or "").lower()
    fd = (final_d or "").lower()
    if "ยี่ห้อ" in sd and "ยี่ห้อ" in fd:
        return "brand_model"
    if "หน้า" in sd and "หน้า" in fd:
        # page number changed?
        sm = re.search(r"หน้า\s*(\d+)", sd)
        fm = re.search(r"หน้า\s*(\d+)", fd)
        if sm and fm and sm.group(1) != fm.group(1):
            return "page"
        return "format"
    if "ยินดี" in fd:
        return "commitment"
    return "format"


def _edit_distance(a: str, b: str) -> int:
    """Cheap Levenshtein for short strings (Col D is bounded). Used purely
    as a feedback metric, not for correctness."""
    a = a or ""; b = b or ""
    la, lb = len(a), len(b)
    if la == 0: return lb
    if lb == 0: return la
    if la > 200 or lb > 200:  # bail on huge strings
        return abs(la - lb)
    prev = list(range(lb + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cost = 0 if ca == cb else 1
            cur.append(min(cur[j - 1] + 1, prev[j] + 1, prev[j - 1] + cost))
        prev = cur
    return prev[lb]


def record_feedback(*,
                    row_num: int,
                    section: str | None,
                    input_b: str,
                    input_pdf_rel: str | None,
                    input_role: str,
                    input_filename: str | None,
                    suggested_c: str,
                    suggested_d: str,
                    suggested_annots: list | None,
                    confidence: float,
                    generator: str,
                    provenance: dict | None,
                    user_action: str,
                    final_c: str,
                    final_d: str,
                    final_annots: list | None) -> int:
    """Append one row to learning_feedback."""
    correction_kind = _classify_correction(suggested_d, final_d)
    distance = _edit_distance(suggested_d or "", final_d or "")
    with db.conn() as c:
        cur = c.execute(
            """INSERT INTO learning_feedback
               (row_num, section, input_b, input_pdf_rel, input_role,
                input_filename, suggested_c, suggested_d, suggested_annots,
                confidence, generator, provenance, user_action,
                final_c, final_d, final_annots, edit_distance_d, correction_kind)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                row_num, section, input_b, input_pdf_rel, input_role,
                input_filename, suggested_c, suggested_d,
                json.dumps(suggested_annots or [], ensure_ascii=False),
                float(confidence or 0),
                generator,
                json.dumps(provenance or {}, ensure_ascii=False),
                user_action,
                final_c, final_d,
                json.dumps(final_annots or [], ensure_ascii=False),
                distance, correction_kind,
            ),
        )
        return cur.lastrowid


# ---------------------------------------------------------------------------
# Pattern mining — distil repeating corrections into learned_patterns
# ---------------------------------------------------------------------------

# Minimum agreeing samples before a pattern is promoted from "noise" to
# "rule". With too few samples we'd overfit on accidents.
PROMOTION_THRESHOLD = 2


def retrain_patterns() -> dict:
    """Walk the feedback log and update learned_patterns.

    Patterns currently mined:

    1. **filename_brand**: For each (filename Latin-token-prefix → brand
       used in final Col D), if ≥N feedbacks agree, register the mapping.
       This fixes the "QNAP filename → NAS detected as brand" class of bug
       by learning the user's actual brand pick.

    2. **section_vendor**: For each (section prefix → vendor in Col E
       used in final), register. Lets the system auto-fill Col E.

    3. **row_format_d**: For each (input_role + section_root → Col D
       template), register. Captures shape preferences ("user always uses
       full filename ref, not section-N").

    4. **annot_position**: For each (catalog folder → label position
       relative to content rect), register. (Stub — full impl when
       annot edits are tracked.)
    """
    counters: dict = {}
    with db.conn() as c:
        # 1. Filename-brand: for every "edited" row whose suggested vs final
        #    differ in brand_model
        for r in c.execute(
            """SELECT input_filename, final_d
               FROM learning_feedback
               WHERE user_action IN ('edited','accepted')
                 AND final_d LIKE 'ยี่ห้อ %'
                 AND input_filename IS NOT NULL
                 AND length(input_filename) > 6"""):
            stem = (r["input_filename"] or "").rsplit(".", 1)[0]
            tokens = _FILENAME_BRAND_TOKEN_RE.findall(stem)
            if not tokens:
                continue
            # Use the FIRST Latin token that's ≥3 chars as the trigger
            trigger = next((t for t in tokens if len(t) >= 3), tokens[0])
            m = re.match(r"ยี่ห้อ\s+(\S+)\s+รุ่น\s+(.+)$", r["final_d"] or "")
            if not m:
                continue
            brand = m.group(1)
            counters.setdefault("filename_brand", {})
            key = (trigger.lower(),)
            counters["filename_brand"].setdefault(key, []).append(brand)

        # 2. Section-vendor mapping: from row's section + final_e
        for r in c.execute(
            """SELECT row_num, section, final_d
               FROM learning_feedback
               WHERE user_action IN ('edited','accepted')
                 AND section IS NOT NULL"""):
            # We need vendor (Col E) — fetch it from rows table
            r2 = c.execute(
                "SELECT col_e FROM rows WHERE row_num = ?",
                (r["row_num"],),
            ).fetchone()
            if not r2 or not r2["col_e"]:
                continue
            sec = r["section"] or ""
            sec_root = _section_root(sec)
            if not sec_root:
                continue
            counters.setdefault("section_vendor", {})
            counters["section_vendor"].setdefault((sec_root,), []).append(r2["col_e"])

        # 3. Format pattern — what shape of Col D does the user prefer for
        #    each (role, section_root)?
        for r in c.execute(
            """SELECT input_role, section, final_d
               FROM learning_feedback
               WHERE user_action IN ('edited','accepted')
                 AND final_d IS NOT NULL"""):
            role = r["input_role"] or "unknown"
            sec = r["section"] or ""
            sec_root = _section_root(sec) or "unknown"
            d = r["final_d"]
            shape = _shape_of_col_d(d)
            counters.setdefault("row_format_d", {})
            counters["row_format_d"].setdefault((role, sec_root), []).append(shape)

    # Now write distilled patterns
    promoted = 0; updated = 0
    for ptype, group in counters.items():
        for key_tuple, values in group.items():
            from collections import Counter
            cnt = Counter(values)
            most_common, n = cnt.most_common(1)[0]
            total = sum(cnt.values())
            if n < PROMOTION_THRESHOLD:
                continue
            confidence = n / total
            trigger_key = key_tuple[0]
            trigger_extra = json.dumps(list(key_tuple[1:])) if len(key_tuple) > 1 else None
            with db.conn() as c:
                existing = c.execute(
                    """SELECT pattern_id, samples_total
                       FROM learned_patterns
                       WHERE pattern_type=? AND trigger_key=?
                         AND ifnull(trigger_extra,'')=ifnull(?, '')""",
                    (ptype, trigger_key, trigger_extra),
                ).fetchone()
                if existing:
                    c.execute(
                        """UPDATE learned_patterns
                           SET output_value=?, samples_total=?,
                               samples_correct=?, confidence=?,
                               last_used_at=CURRENT_TIMESTAMP
                           WHERE pattern_id=?""",
                        (most_common, total, n, confidence, existing["pattern_id"]),
                    )
                    updated += 1
                else:
                    c.execute(
                        """INSERT INTO learned_patterns
                           (pattern_type, trigger_key, trigger_extra, output_value,
                            samples_total, samples_correct, confidence)
                           VALUES (?, ?, ?, ?, ?, ?, ?)""",
                        (ptype, trigger_key, trigger_extra, most_common,
                         total, n, confidence),
                    )
                    promoted += 1
    return {"promoted": promoted, "updated": updated, "patterns_examined": len(counters)}


def _shape_of_col_d(d: str) -> str:
    """Coarse shape signature of a Col D value — the retrain step compares
    *shapes* not exact strings, so a per-row brand/section can vary while
    the formatting preference still gets captured."""
    if not d:
        return "(empty)"
    if d.startswith("ยินดีปฏิบัติ"):
        return "commitment"
    if d.startswith("ยี่ห้อ "):
        return "brand_model"
    if d.startswith("เทียบเท่าข้อกำหนด"):
        if "ข้อย่อย" in d:
            return "equivalent_subitem"
        return "equivalent_item"
    if d.startswith("สูงกว่าข้อกำหนด"):
        return "higher"
    if re.match(r"^\d+\.\d+(?:\.\d+)*-\d+", d):
        return "filename_format"
    if re.match(r"^\d+\.\d+", d):
        return "section_prefix"
    return "other"


# ---------------------------------------------------------------------------
# Suggestion pipeline — applies learned patterns + rules + (optional) LLM
# ---------------------------------------------------------------------------

def list_patterns(pattern_type: str | None = None) -> list[dict]:
    sql = "SELECT * FROM learned_patterns"
    params: tuple = ()
    if pattern_type:
        sql += " WHERE pattern_type = ?"
        params = (pattern_type,)
    sql += " ORDER BY confidence DESC, samples_total DESC"
    out = []
    with db.conn() as c:
        for r in c.execute(sql, params):
            out.append({
                "id": r["pattern_id"],
                "type": r["pattern_type"],
                "trigger_key": r["trigger_key"],
                "trigger_extra": r["trigger_extra"],
                "output": r["output_value"],
                "samples_total": r["samples_total"],
                "samples_correct": r["samples_correct"],
                "confidence": r["confidence"],
                "enabled": bool(r["enabled"]),
                "last_used_at": r["last_used_at"],
                "note": r["note"],
            })
    return out


def apply_learned_brand(filename_stem: str) -> tuple[str, dict] | tuple[None, None]:
    """Return (brand, provenance) if a learned filename_brand pattern fires."""
    if not filename_stem:
        return None, None
    tokens = _FILENAME_BRAND_TOKEN_RE.findall(filename_stem)
    triggers = [t.lower() for t in tokens if len(t) >= 3]
    if not triggers:
        return None, None
    with db.conn() as c:
        # Match any trigger token from the filename
        placeholder = ",".join("?" for _ in triggers)
        rows = c.execute(
            f"""SELECT trigger_key, output_value, confidence, samples_total
                FROM learned_patterns
                WHERE pattern_type='filename_brand' AND enabled=1
                  AND trigger_key IN ({placeholder})
                ORDER BY confidence DESC, samples_total DESC LIMIT 1""",
            triggers,
        ).fetchall()
    if not rows:
        return None, None
    r = rows[0]
    return r["output_value"], {
        "kind": "learned",
        "pattern_type": "filename_brand",
        "trigger": r["trigger_key"],
        "confidence": r["confidence"],
        "samples": r["samples_total"],
    }


def apply_learned_vendor(section: str | None) -> tuple[str, dict] | tuple[None, None]:
    sec_root = _section_root(section)
    if not sec_root:
        return None, None
    with db.conn() as c:
        r = c.execute(
            """SELECT output_value, confidence, samples_total
               FROM learned_patterns
               WHERE pattern_type='section_vendor' AND enabled=1
                 AND trigger_key=? LIMIT 1""",
            (sec_root,),
        ).fetchone()
    if not r:
        return None, None
    return r["output_value"], {
        "kind": "learned", "pattern_type": "section_vendor",
        "trigger": sec_root, "confidence": r["confidence"],
        "samples": r["samples_total"],
    }


def confidence_score(*, generator: str, provenance: dict | None,
                      role: str, has_match: bool, warnings: int = 0) -> float:
    """Heuristic confidence in [0, 1]."""
    base = 0.55
    if generator.startswith("learned"):
        base = 0.92
    elif generator.startswith("rules+pattern"):
        base = 0.82
    elif generator == "rules":
        base = 0.65
    elif generator.startswith("llm"):
        base = 0.78
    if role == "section_header":
        base += 0.05  # filename parsing is more reliable
    if has_match:
        base += 0.05
    base -= 0.05 * max(0, warnings)
    return max(0.05, min(0.99, base))


# ---------------------------------------------------------------------------
# Stats for the Learning UI
# ---------------------------------------------------------------------------

def feedback_stats(window_days: int = 30) -> dict:
    """Aggregate metrics over the last `window_days` days."""
    cutoff = (datetime.now().timestamp() - window_days * 86400)
    iso_cutoff = datetime.fromtimestamp(cutoff).isoformat(timespec="seconds")
    with db.conn() as c:
        total = c.execute(
            "SELECT COUNT(*) c FROM learning_feedback WHERE ts >= ?",
            (iso_cutoff,),
        ).fetchone()["c"]
        accepted = c.execute(
            "SELECT COUNT(*) c FROM learning_feedback WHERE ts >= ? AND user_action='accepted'",
            (iso_cutoff,),
        ).fetchone()["c"]
        edited = c.execute(
            "SELECT COUNT(*) c FROM learning_feedback WHERE ts >= ? AND user_action='edited'",
            (iso_cutoff,),
        ).fetchone()["c"]
        rejected = c.execute(
            "SELECT COUNT(*) c FROM learning_feedback WHERE ts >= ? AND user_action='rejected'",
            (iso_cutoff,),
        ).fetchone()["c"]
        kinds = {}
        for r in c.execute(
            """SELECT correction_kind, COUNT(*) c FROM learning_feedback
               WHERE ts >= ? GROUP BY correction_kind""",
            (iso_cutoff,),
        ):
            kinds[r["correction_kind"] or "(none)"] = r["c"]
        n_patterns = c.execute("SELECT COUNT(*) c FROM learned_patterns").fetchone()["c"]
        n_enabled = c.execute(
            "SELECT COUNT(*) c FROM learned_patterns WHERE enabled=1"
        ).fetchone()["c"]
    accuracy = (accepted / total) if total else 0.0
    return {
        "window_days": window_days,
        "total_feedbacks": total,
        "accepted": accepted,
        "edited": edited,
        "rejected": rejected,
        "accuracy": accuracy,
        "correction_kinds": kinds,
        "patterns_total": n_patterns,
        "patterns_enabled": n_enabled,
        "llm": llm_status(),
    }
