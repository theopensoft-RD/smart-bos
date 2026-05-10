"""
Export endpoint tests — covers Phase 2.1.
"""

from __future__ import annotations


def test_export_preview_modes(client, gui):
    """preview supports full / comply_only / catalogs_only modes."""
    for mode in ("full", "comply_only", "catalogs_only"):
        r = client.get(f"/api/export/preview?mode={mode}")
        assert r.status_code == 200, f"mode={mode}"
        j = r.get_json()
        assert j["ok"] is True


def test_export_preview_section_filter(client, gui):
    """section= filter narrows the catalog count."""
    r1 = client.get("/api/export/preview?mode=full")
    full = r1.get_json()["catalog_count"]
    r2 = client.get("/api/export/preview?mode=full&section=5.1.1")
    sec = r2.get_json()["catalog_count"]
    assert sec <= full
    assert sec > 0


def test_export_preview_bound_only(client, gui):
    """bound_only=1 only includes catalogs that have row_catalog_links."""
    r = client.get("/api/export/preview?mode=full&bound_only=1")
    assert r.status_code == 200
    j = r.get_json()
    # Could be 0 if no catalogs are bound yet — the route should still respond
    assert j["ok"] is True
    assert "catalog_count" in j


def test_export_download_404_for_missing_file(client, gui):
    r = client.get("/api/export/download?file=does-not-exist.pdf")
    assert r.status_code == 404


def test_export_download_path_traversal_blocked(client, gui):
    """File names must not contain / \\ or .."""
    for bad in ("../etc/passwd", "/etc/passwd", "..\\config",
                "good/../bad.pdf"):
        r = client.get(f"/api/export/download?file={bad}")
        assert r.status_code == 400, f"should 400 on {bad}"


def test_export_list_returns_array(client, gui):
    r = client.get("/api/export/list")
    assert r.status_code == 200
    j = r.get_json()
    assert isinstance(j.get("items"), list)
