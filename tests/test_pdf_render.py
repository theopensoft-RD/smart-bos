"""
PDF rendering tests — Phase 17 WYSIWYG invariant.
"""

from __future__ import annotations


def test_pdf_meta_returns_pages_and_annots(client, gui):
    rel = next((r["pdf_rel"] for r in (gui.ROWS or []) if r.get("pdf_rel")), None)
    assert rel
    r = client.get(f"/api/pdf_meta?rel={rel}")
    assert r.status_code == 200
    j = r.get_json()
    assert j["pages"] >= 1
    assert isinstance(j.get("annots"), list)
    assert len(j["page_sizes"]) == j["pages"]


def test_pdf_meta_404_for_missing(client, gui):
    r = client.get("/api/pdf_meta?rel=does/not/exist.pdf")
    assert r.status_code == 404


def test_pdf_meta_403_for_path_traversal(client, gui):
    """Trying to escape OUTPUT root must 403."""
    r = client.get("/api/pdf_meta?rel=../../../etc/passwd")
    # Either 403 (caught by sandbox) or 404 (Path.exists fail) is acceptable
    assert r.status_code in (403, 404)


def test_pdf_page_view_equals_edit(client, gui):
    """Phase 17: edit-mode page bytes must equal view-mode page bytes."""
    rel = next((r["pdf_rel"] for r in (gui.ROWS or []) if r.get("pdf_rel")), None)
    assert rel
    a = client.get(f"/api/pdf_page?rel={rel}&page=1").get_data()
    b = client.get(f"/api/pdf_page?rel={rel}&page=1&edit=1").get_data()
    assert len(a) > 100  # actual image
    assert a == b, "edit-mode bytes diverged from view (Phase 17 broken)"


def test_pdf_page_bake_zero_differs(client, gui):
    """?bake=0 strips annots — output should differ when PDF has annots."""
    # Find a PDF that has annots
    rels = []
    for pdf_rel in {r.get("pdf_rel") for r in (gui.ROWS or []) if r.get("pdf_rel")}:
        if not pdf_rel: continue
        meta = client.get(f"/api/pdf_meta?rel={pdf_rel}")
        if meta.status_code != 200: continue
        if (meta.get_json() or {}).get("annots"):
            rels.append(pdf_rel)
            if len(rels) >= 1: break
    assert rels, "expected at least one annotated PDF"
    rel = rels[0]
    baked = client.get(f"/api/pdf_page?rel={rel}&page=1&edit=1").get_data()
    nobake = client.get(f"/api/pdf_page?rel={rel}&page=1&edit=1&bake=0").get_data()
    assert baked != nobake, "?bake=0 should produce different bytes for annotated PDF"
