"""
Smoke tests — fast read-only checks that catch the most common
regressions before they hit the user.

Run with:
    .venv/bin/pytest -v
"""

from __future__ import annotations


def test_boot_registers_all_critical_routes(client, gui):
    """All public API routes the frontend depends on must exist."""
    rules = {str(r) for r in gui.app.url_map.iter_rules()}
    must_have = {
        "/",
        "/api/index",
        "/api/status",
        "/api/pdf_meta",
        "/api/pdf_page",
        "/api/auto_annotate/preview",
        "/api/auto_annotate/apply",
        "/api/learn/llm_status",
        "/api/learn/feedback",
        "/api/learn/retrain",
        "/api/learn/patterns",
        "/api/db/stats",
        "/api/versions/sync",
        "/api/row/col_d",
        "/api/row/col_d/suggest",   # Phase B3
        "/api/claude/stream",       # Phase 1 (Claude Code)
        "/api/manual_annotate/context",
        "/api/manual_annotate/save",
        "/api/reannotate/context",
        "/api/reannotate/save",
        "/api/settings/api_key",
    }
    missing = must_have - rules
    assert not missing, f"missing routes: {missing}"


def test_index_returns_rows_and_sections(client, gui):
    """The main index endpoint should load rows + section tree."""
    r = client.get("/api/index")
    assert r.status_code == 200
    j = r.get_json()
    assert isinstance(j.get("rows"), list)
    assert len(j["rows"]) > 0, "expected at least 1 row"
    assert isinstance(j.get("sections"), list)
    # `tree` is a single root dict with a `children` list (recursive)
    tree = j.get("tree")
    assert isinstance(tree, dict)
    assert isinstance(tree.get("children"), list)
    # Every row should have minimum keys the frontend depends on
    for r0 in j["rows"][:5]:
        for k in ("row", "B"):
            assert k in r0, f"row missing required key {k}"


def test_pdf_render_view_equals_edit_byte_exact(client, gui):
    """Phase 17 invariant: edit-mode page bytes == view-mode bytes
    (so the SVG overlay sits on the same pixels)."""
    rows = gui.ROWS or []
    pdf_rel = next((r["pdf_rel"] for r in rows if r.get("pdf_rel")), None)
    assert pdf_rel, "expected at least one row with a catalog PDF"

    view = client.get(f"/api/pdf_page?rel={pdf_rel}&page=1")
    edit = client.get(f"/api/pdf_page?rel={pdf_rel}&page=1&edit=1")
    assert view.status_code == 200 and edit.status_code == 200
    assert view.content_type.startswith("image/"), \
        f"expected image, got {view.content_type}"
    assert view.get_data() == edit.get_data(), \
        "edit-mode bytes diverged from view-mode (Phase 17 invariant broken)"


def test_col_d_suggest_returns_ranked_candidates(client, gui):
    """Phase B3: live autocomplete must return at least the AI proposal
    OR a shape template for any row with a PDF."""
    rows = gui.ROWS or []
    test_row = next((r["row"] for r in rows if r.get("pdf_rel")), None)
    assert test_row, "expected at least one row with pdf_rel"

    r = client.get(f"/api/row/col_d/suggest?row={test_row}")
    assert r.status_code == 200
    j = r.get_json()
    assert j.get("ok") is True
    assert isinstance(j.get("suggestions"), list)
    # Even if AI proposal is empty, shape templates should be present
    assert len(j["suggestions"]) >= 1
    kinds = {s.get("kind") for s in j["suggestions"]}
    assert kinds <= {"ai", "neighbor", "shape"}, \
        f"unexpected suggestion kinds: {kinds}"


def test_claude_stream_endpoint_responds_or_503(client, gui):
    """Phase 1 SSE endpoint: must either stream events (200, text/event-stream)
    or return 503 with a hint when the provider can't be initialized.
    Either way, the route must NOT 500."""
    rows = gui.ROWS or []
    test_row = next((r["row"] for r in rows if r.get("pdf_rel")), 1)
    r = client.get(f"/api/claude/stream?row={test_row}")
    assert r.status_code in (200, 503), \
        f"unexpected status {r.status_code} — body: {r.get_data(as_text=True)[:200]}"
    if r.status_code == 200:
        assert r.content_type.startswith("text/event-stream")
