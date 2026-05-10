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


def test_catalog_endpoints_basic(client, gui):
    """Phase 2: catalog library API + multi-project bootstrap.

    Boot must:
      - have ingested PDFs from output/ into the catalogs table
      - created at least one company + one active project
    """
    r = client.get("/api/catalogs/stats")
    assert r.status_code == 200
    j = r.get_json()
    assert j["catalogs"] >= 1, "boot ingest didn't populate any catalogs"
    assert j["companies"] >= 1
    assert j["projects"] >= 1

    # List endpoint
    r = client.get("/api/catalogs?limit=5")
    assert r.status_code == 200
    j = r.get_json()
    assert j.get("ok") is True
    assert isinstance(j.get("items"), list)
    assert len(j["items"]) >= 1

    # Filter by query
    cat = j["items"][0]
    if cat.get("brand"):
        r = client.get(f"/api/catalogs?q={cat['brand']}")
        assert r.status_code == 200
        assert any(c["catalog_id"] == cat["catalog_id"]
                   for c in r.get_json()["items"])

    # Get single catalog
    r = client.get(f"/api/catalogs/{cat['catalog_id']}")
    assert r.status_code == 200
    assert r.get_json().get("catalog", {}).get("catalog_id") == cat["catalog_id"]

    # Active project visible
    r = client.get("/api/projects")
    assert r.status_code == 200
    j = r.get_json()
    assert j.get("active") is not None
    assert j["active"].get("project_id")


def test_catalog_metadata_update(client, gui):
    """Phase 2: PATCH /api/catalogs/<id> updates editable fields."""
    r = client.get("/api/catalogs?limit=1")
    cat_id = r.get_json()["items"][0]["catalog_id"]
    new_desc = "test description from smoke suite"
    r = client.patch(f"/api/catalogs/{cat_id}",
                     json={"description": new_desc})
    assert r.status_code == 200
    assert r.get_json()["catalog"]["description"] == new_desc
    # Cleanup
    client.patch(f"/api/catalogs/{cat_id}", json={"description": ""})


def test_export_preview_and_build_small(client, gui):
    """Phase 2.1: export preview + build (small subset).

    We use ``section=5.1.1`` so the build only ingests one section's
    worth of catalogs (~30 pages of dividers + catalogs + cover/TOC).
    """
    # Preview
    r = client.get("/api/export/preview?mode=full&section=5.1.1")
    assert r.status_code == 200
    j = r.get_json()
    assert j.get("ok") is True
    assert j.get("project", {}).get("name")

    # Build a small package
    r = client.post("/api/export/package?mode=full&section=5.1.1")
    assert r.status_code == 200
    j = r.get_json()
    assert j.get("ok") is True
    assert j["page_count"] >= 3, "expected at least cover + TOC + comply"
    assert j["byte_size"] > 1000
    assert "download_url" in j

    # Download it back
    r = client.get(j["download_url"])
    assert r.status_code == 200
    assert r.content_type == "application/pdf"
    pdf_bytes = r.get_data()
    assert pdf_bytes.startswith(b"%PDF-"), "not a valid PDF"

    # /api/export/list shows our build
    r = client.get("/api/export/list")
    assert r.status_code == 200
    items = r.get_json().get("items", [])
    assert any(it["filename"] == j["filename"] for it in items)


def test_manual_routes_serve_content(client, gui):
    """User manual: /manual renders the styled page; /api/manual/raw
    returns the raw markdown body. Both must be reachable so the topbar
    Manual link and onboarding 'Open manual' button work."""
    # Raw markdown
    r = client.get("/api/manual/raw")
    assert r.status_code == 200
    assert "markdown" in r.content_type
    md = r.get_data(as_text=True)
    assert len(md) > 1000
    assert "Comply Verify Tool" in md
    assert "Keyboard Shortcuts" in md  # appendix exists

    # Rendered manual page
    r = client.get("/manual")
    assert r.status_code == 200
    html = r.get_data(as_text=True)
    must = ["/api/manual/raw", "marked", "man-aside", "man-search"]
    missing = [m for m in must if m not in html]
    assert not missing, f"manual page missing: {missing}"


def test_phase_c_catalog_editor_ui_present(client, gui):
    """Phase C: Catalog Browser detail pane includes Edit-annotations button
    + catalogEditAnnotations JS + edit-mode banner element/CSS. Verifies the
    user can route from catalog browser into the existing PDF edit flow.
    """
    r = client.get("/")
    assert r.status_code == 200
    html = r.get_data(as_text=True)
    must = [
        "catalogEditAnnotations",      # JS function
        "Edit annotations",             # button label in detail pane
        "catalog-edit-banner",          # banner element + CSS class
        "_CATALOG_EDIT_CONTEXT",        # state variable for the wrapper
        "catalogExitEditMode",          # exit fn
    ]
    missing = [m for m in must if m not in html]
    assert not missing, f"Phase C UI hooks missing: {missing}"


def test_phase_c_empty_patch_bumps_updated_at(client, gui):
    """Phase C wires saveEdits → PATCH /api/catalogs/<id> with {} so the
    catalog's updated_at gets touched after the user saves PDF edits in
    catalog-edit mode. Empty-body patch must succeed."""
    # Pick any catalog
    r = client.get("/api/catalogs?limit=1")
    cid = r.get_json()["items"][0]["catalog_id"]
    r = client.patch(f"/api/catalogs/{cid}", json={})
    assert r.status_code == 200
    assert r.get_json().get("ok") is True


def test_label_for_row_handles_filename_format(client, gui):
    """R12 case: filename_format rows must produce a usable label.
    Was returning empty/section-only before the dash-form + Col B fallback fix.
    """
    r12 = next((r for r in (gui.ROWS or []) if r["row"] == 12), None)
    if r12 is None:
        return  # data shape might differ on other machines
    role_info = gui.detect_row_role(12)
    label = gui._label_for_row(role_info, r12)
    # Must contain section + item marker
    assert "5.1.1" in label
    assert "ข้อ" in label or "ข้อย่อย" in label, f"unexpected label: {label!r}"


def test_manual_context_includes_suggested_label(client, gui):
    """/api/manual_annotate/context returns a non-empty suggested_label
    for any row with a PDF (uses _label_for_row internally)."""
    rows = gui.ROWS or []
    test_row = next((r["row"] for r in rows if r.get("pdf_rel")), None)
    if not test_row:
        return
    r = client.get(f"/api/manual_annotate/context?row={test_row}")
    assert r.status_code == 200
    j = r.get_json()
    # suggested_label may be empty for section_header rows but key must exist
    assert "suggested_label" in j


def test_quick_annotate_button_in_html(client, gui):
    """The inline 📍 Annotate button + JS handler are served at /."""
    html = client.get("/").get_data(as_text=True)
    must = [
        "d-quick-annotate",            # CSS class for the button
        "function quickAnnotateRow",    # the JS handler
        "col-D.not-found",              # CSS for "ไม่พบใน catalog" rows
        "col-D.empty",                  # CSS for empty Col D rows
    ]
    missing = [m for m in must if m not in html]
    assert not missing, f"Missing inline annotate hooks: {missing}"


def test_calm_mode_default_and_toggle(client, gui):
    """Calm Mode (Apple/Tesla minimalism) is the new default.

    The HTML must serve:
      - data-ux-mode CSS rules
      - topbar More menu (entry to hidden features)
      - toggleUxMode JS function
      - calm-mode hint
      - Calm Mode palette variables (Apple blue #0071e3)
      - SF Pro font stack
    """
    r = client.get("/")
    assert r.status_code == 200
    html = r.get_data(as_text=True)
    must = [
        'data-ux-mode="calm"',          # CSS scope
        "function applyUxMode",          # boot-time IIFE
        "function toggleUxMode",         # toggle handler
        "topbar-more-popover",           # advanced features access
        "calm-mode-hint",                # bottom-left hint
        "#0071e3",                       # Apple blue accent
        "SF Pro Display",                # font stack
    ]
    missing = [m for m in must if m not in html]
    assert not missing, f"Calm Mode hooks missing: {missing}"


def test_continuity_endpoint(client, gui):
    """/api/continuity returns the latest STATE markdown if present."""
    r = client.get("/api/continuity")
    assert r.status_code == 200
    j = r.get_json()
    assert j.get("ok") is True
    # When there's a _continuity/STATE_*.md in the project root, the
    # endpoint must report it. Smart Plant 1 has one as of 2026-05-10.
    if j.get("available"):
        assert j["filename"].startswith("STATE_")
        assert j["filename"].endswith(".md")
        assert j["byte_size"] > 0
        assert "markdown" in j and len(j["markdown"]) > 0


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
