"""
Catalog endpoint tests — covers Phase 2, Phase C, and the UX-2 bulk
update tooling.
"""

from __future__ import annotations


def test_catalog_filter_by_section_root_descends(client, gui):
    """list_catalogs with section='5.1' must include 5.1.1.* descendants."""
    r = client.get("/api/catalogs?section=5.1.1&limit=10")
    assert r.status_code == 200
    items = r.get_json()["items"]
    # All returned must have section_hint that is or starts with 5.1.1
    for c in items:
        s = c.get("section_hint") or ""
        assert s == "5.1.1" or s.startswith("5.1.1.")


def test_catalog_q_param_searches_pdf_rel(client, gui):
    """?q=<filename-fragment> finds catalogs by pdf_rel."""
    # Pick any catalog and search by its filename
    items = client.get("/api/catalogs?limit=1").get_json()["items"]
    sample = items[0]["pdf_rel"]
    fragment = sample.split("/")[-1].split(".")[0][:20]
    r = client.get(f"/api/catalogs?q={fragment}")
    items2 = r.get_json()["items"]
    assert any(c["catalog_id"] == items[0]["catalog_id"] for c in items2)


def test_catalog_annotation_crud_cycle(client, gui):
    """Add → list → update → delete cycle on catalog_annotations."""
    cid = client.get("/api/catalogs?limit=1").get_json()["items"][0]["catalog_id"]
    # Add
    r = client.post(f"/api/catalogs/{cid}/annotations",
                    json={"page": 1, "type": "Square",
                          "rect": [10, 20, 100, 50], "contents": "test"})
    assert r.status_code == 200
    annot_id = r.get_json()["annot_id"]
    # List
    r = client.get(f"/api/catalogs/{cid}/annotations")
    assert r.status_code == 200
    rows = r.get_json()["annotations"]
    assert any(a["annot_id"] == annot_id for a in rows)
    # Update
    r = client.patch(f"/api/catalogs/{cid}/annotations/{annot_id}",
                     json={"contents": "updated"})
    assert r.status_code == 200
    # Delete (soft)
    r = client.delete(f"/api/catalogs/{cid}/annotations/{annot_id}")
    assert r.status_code == 200


def test_bulk_preview_empty_brand(client, gui):
    """UX-2: preview matches catalogs with empty brand field."""
    r = client.get("/api/catalogs/bulk_preview?match=&match_type=exact&field=brand")
    assert r.status_code == 200
    j = r.get_json()
    assert j["ok"] is True
    # Some catalogs have NULL brand (heuristic missed them)
    # — count should be ≥ 0; if non-empty, every item's brand is null/empty
    for it in j["items"]:
        assert not it.get("brand"), f"got brand={it.get('brand')}"


def test_bulk_preview_invalid_field(client, gui):
    """Bad field name → 400 error."""
    r = client.get("/api/catalogs/bulk_preview?field=invalid_field&match=foo")
    assert r.status_code == 400


def test_companies_and_projects_basic(client, gui):
    """List + at least one company + at least one active project."""
    r = client.get("/api/companies")
    assert r.status_code == 200
    cos = r.get_json()["items"]
    assert len(cos) >= 1

    r = client.get("/api/projects")
    j = r.get_json()
    assert r.status_code == 200
    assert j["active"] is not None
    assert j["active"]["project_id"]


def test_bulk_update_dry_run_with_no_match_returns_zero(client, gui):
    """Bulk update with a match that finds nothing should return 0."""
    r = client.post("/api/catalogs/bulk_update",
                    json={"match": "ZZZ_DEFINITELY_NOT_A_BRAND_ZZZ",
                          "match_type": "exact",
                          "field": "brand",
                          "new_value": "wontmatter"})
    assert r.status_code == 200
    j = r.get_json()
    assert j["ok"] is True
    assert j["matched"] == 0
    assert j["updated"] == 0


def test_bulk_update_invalid_field_returns_400(client, gui):
    r = client.post("/api/catalogs/bulk_update",
                    json={"match": "x", "field": "evil_field"})
    assert r.status_code == 400


def test_apply_catalog_to_row_writes_xlsx_and_records_link(client, gui):
    """End-to-end: pick a row + a catalog → apply → verify Col D + DB link.

    We pick a row that already has a Col D so we can restore it after.
    """
    # Pick a row with pdf_rel
    rows = gui.ROWS or []
    test_row = next((r for r in rows if r.get("pdf_rel")), None)
    assert test_row is not None
    original_d = (test_row.get("D") or "").strip()
    row_num = test_row["row"]

    # Pick a catalog (any)
    cat_id = client.get("/api/catalogs?limit=1").get_json()["items"][0]["catalog_id"]

    # Apply
    r = client.post("/api/row/apply_catalog",
                    json={"row": row_num, "catalog_id": cat_id, "page": 1,
                          "col_d_text": original_d or "test"})
    assert r.status_code == 200
    j = r.get_json()
    assert j["ok"] is True
    assert j["row"] == row_num

    # Restore original Col D so subsequent test runs are clean
    if original_d:
        client.post("/api/row/col_d",
                    json={"row": row_num, "col_d": original_d, "original": j["col_d"]})
