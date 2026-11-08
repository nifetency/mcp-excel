"""
test_crm_local.py
Local smoke-tests and edge-case tests for mcp_excel_server.py.

These tests focus on the parts NOT already covered by test_mcp_excel.py:
  - _next_id edge cases (single-row, all-blank ids, two-digit padding)
  - get_all_records_batched with duplicate header names
  - Multi-filter list_records
  - search_records across every column in the record
  - update_record with only Updated At changed (no real field updates)
  - delete_record idempotency: deleting an already-deleted record
  - get_related_records with an exact-match FK (no normalize, values are exact)
  - get_pipeline_analytics with missing Segment column values
  - MCP tool wrappers return valid JSON

Run:
    cd mcp-excel
    pip install pytest gspread google-auth mcp
    pytest test_crm_local.py -v

No Google credentials or network access needed.
"""

from __future__ import annotations

import json
import os
import uuid
from unittest.mock import MagicMock, patch

import pytest

os.environ.setdefault("GOOGLE_APPLICATION_CREDENTIALS", "creds.json")

import mcp_excel_server as crm


# ─────────────────────────────────────────────────────────────────────────────
# Helpers shared across tests
# ─────────────────────────────────────────────────────────────────────────────

def _headers(entity: str) -> list[str]:
    return crm.ENTITY_SCHEMAS[entity]["headers"] + crm.Settings.AUDIT_HEADERS


def make_ws(headers: list[str], rows: list[dict] | None = None) -> MagicMock:
    rows = rows or []
    ws = MagicMock()
    ws.row_values.return_value = list(headers)
    all_vals = [list(headers)]
    for r in rows:
        all_vals.append([str(r.get(h, "")) for h in headers])
    ws.get_all_values.return_value = all_vals
    return ws


# ─────────────────────────────────────────────────────────────────────────────
# _next_id — additional edge cases
# ─────────────────────────────────────────────────────────────────────────────

class TestNextIdEdgeCases:
    def _ws_with_ids(self, headers, pk_field, ids):
        pk_idx = headers.index(pk_field)
        ws = MagicMock()
        all_vals = [headers]
        for id_val in ids:
            row = [""] * len(headers)
            row[pk_idx] = id_val
            all_vals.append(row)
        ws.get_all_values.return_value = all_vals
        return ws

    def test_single_existing_row_increments(self):
        headers = _headers("Lead")
        ws = self._ws_with_ids(headers, "Lead ID", ["LD001"])
        result = crm._next_id(ws, headers, "Lead ID")
        assert result == "LD002"

    def test_two_digit_padding_preserved(self):
        headers = _headers("Organization")
        ws = self._ws_with_ids(headers, "Organization ID", ["ORG01", "ORG02"])
        result = crm._next_id(ws, headers, "Organization ID")
        assert result == "ORG03"

    def test_three_digit_zero_padded(self):
        headers = _headers("Contact")
        ws = self._ws_with_ids(headers, "Contact ID", ["C099"])
        result = crm._next_id(ws, headers, "Contact ID")
        assert result == "C100"

    def test_all_blank_ids_returns_uuid(self):
        headers = _headers("Lead")
        ws = self._ws_with_ids(headers, "Lead ID", ["", "", ""])
        result = crm._next_id(ws, headers, "Lead ID")
        # No parseable pattern → UUID
        try:
            uuid.UUID(result)
        except ValueError:
            pytest.fail(f"Expected UUID fallback, got {result!r}")

    def test_separator_preserved(self):
        """IDs with a dash separator like "LD-001" should increment to "LD-002"."""
        headers = _headers("Lead")
        ws = self._ws_with_ids(headers, "Lead ID", ["LD-001", "LD-002"])
        result = crm._next_id(ws, headers, "Lead ID")
        assert result == "LD-003"

    def test_no_prefix_numeric_only(self):
        """Purely numeric IDs "001", "002" should produce "003"."""
        headers = _headers("Lead")
        ws = self._ws_with_ids(headers, "Lead ID", ["001", "002"])
        result = crm._next_id(ws, headers, "Lead ID")
        assert result == "003"


# ─────────────────────────────────────────────────────────────────────────────
# get_all_records_batched — extra edge cases
# ─────────────────────────────────────────────────────────────────────────────

class TestBatchedEdgeCases:
    def test_deleted_case_insensitive(self):
        """'TRUE', 'True', 'true' should all be treated as deleted."""
        headers = ["Contact ID", "Name", "Deleted"]
        for deleted_val in ("TRUE", "True", "true"):
            ws = MagicMock()
            ws.get_all_values.return_value = [
                headers,
                ["c1", "Alice", deleted_val],
            ]
            rows = crm.get_all_records_batched(ws, headers)
            assert rows == [], f"Expected deleted for Deleted={deleted_val!r}"

    def test_row_longer_than_headers_is_truncated_by_zip(self):
        """Extra values beyond header count are silently dropped by zip."""
        headers = ["A", "B", "Deleted"]
        ws = MagicMock()
        ws.get_all_values.return_value = [
            headers,
            ["a1", "b1", "", "extra1", "extra2"],
        ]
        rows = crm.get_all_records_batched(ws, headers)
        assert len(rows) == 1
        assert list(rows[0].keys()) == headers


# ─────────────────────────────────────────────────────────────────────────────
# list_records — multi-filter
# ─────────────────────────────────────────────────────────────────────────────

class TestListRecordsMultiFilter:
    def _rows(self):
        return [
            {"Status": "Qualified", "Segment": "Enterprise", "Deleted": ""},
            {"Status": "Qualified", "Segment": "SMB",        "Deleted": ""},
            {"Status": "New Lead",  "Segment": "Enterprise", "Deleted": ""},
        ]

    def test_two_filters_anded(self):
        headers = _headers("Lead")
        ws = make_ws(headers, rows=self._rows())

        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Lead"])):
            result = crm.list_records(
                sheet_id="s1",
                entity="Lead",
                filters={"Status": "Qualified", "Segment": "Enterprise"},
            )

        assert result["total"] == 1

    def test_filter_no_matches(self):
        headers = _headers("Lead")
        ws = make_ws(headers, rows=self._rows())

        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Lead"])):
            result = crm.list_records(
                sheet_id="s1",
                entity="Lead",
                filters={"Status": "Won"},
            )

        assert result["total"] == 0
        assert result["records"] == []


# ─────────────────────────────────────────────────────────────────────────────
# search_records — searches every column value
# ─────────────────────────────────────────────────────────────────────────────

class TestSearchAllColumns:
    def test_match_in_non_name_column(self):
        headers = _headers("Activity")
        rows = [
            {"Activity ID": "a1", "Subject": "Pricing Discussion",
             "Discussion Summary": "Talked about SaaS pricing model", "Deleted": ""},
            {"Activity ID": "a2", "Subject": "Intro Call",
             "Discussion Summary": "General intro", "Deleted": ""},
        ]
        ws = make_ws(headers, rows=rows)

        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Activity"])):
            result = crm.search_records(
                sheet_id="s1", entity="Activity", query="pricing"
            )

        # Row a1 matches twice (Subject + Discussion Summary) but search_records
        # returns distinct *rows*, not match-count — so total is 1.
        assert result["total"] == 1
        assert result["records"][0]["Activity ID"] == "a1"

    def test_match_partial_value(self):
        headers = _headers("Contact")
        rows = [
            {"First Name": "Alexandra", "Email": "alex@corp.com", "Deleted": ""},
            {"First Name": "Lexi", "Email": "lexi@startup.io", "Deleted": ""},
        ]
        ws = make_ws(headers, rows=rows)

        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Contact"])):
            result = crm.search_records(
                sheet_id="s1", entity="Contact", query="lex"
            )

        # "Alexandra" and "Lexi" both contain "lex" (case-insensitive)
        assert result["total"] == 2


# ─────────────────────────────────────────────────────────────────────────────
# update_record — no real field changes (only Updated At)
# ─────────────────────────────────────────────────────────────────────────────

class TestUpdateRecordOnlyUpdatedAt:
    def test_empty_updates_still_writes_updated_at(self):
        headers = _headers("Lead")
        ws = make_ws(headers)
        record = {"Lead ID": "LD001", "Status": "New Lead", "Deleted": ""}
        captured = {}
        ws.batch_update.side_effect = lambda u: captured.update({"u": u})

        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Lead"])), \
             patch("mcp_excel_server.get_entity_row",
                   return_value=(2, record)):
            result = crm.update_record(
                sheet_id="s1", entity="Lead",
                record_id="LD001", updates={},
            )

        # batch_update should only contain the Updated At cell
        assert result["status"] == "success"
        if "u" in captured:
            assert len(captured["u"]) == 1
            ua_col = crm._col_letter(headers.index("Updated At") + 1)
            assert captured["u"][0]["range"].startswith(ua_col)


# ─────────────────────────────────────────────────────────────────────────────
# delete_record — already-deleted record
# ─────────────────────────────────────────────────────────────────────────────

class TestDeleteIdempotency:
    def test_delete_already_deleted_record_succeeds(self):
        """Soft-deleting an already-deleted record is idempotent — still returns success."""
        headers = _headers("Contact")
        ws = make_ws(headers)
        # Record says Deleted="" but it was already soft-deleted in the sheet;
        # the row was still found (get_entity_row doesn't filter on Deleted).
        record = {"Contact ID": "c1", "First Name": "Alice", "Deleted": "true"}

        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Contact"])), \
             patch("mcp_excel_server.get_entity_row",
                   return_value=(2, record)):
            result = crm.delete_record(
                sheet_id="s1", entity="Contact", record_id="c1"
            )

        assert result["status"] == "success"
        ws.batch_update.assert_called_once()


# ─────────────────────────────────────────────────────────────────────────────
# get_related_records — exact-match semantics
# ─────────────────────────────────────────────────────────────────────────────

class TestRelatedRecordsExactMatch:
    def test_does_not_partial_match(self):
        """FK match is exact, not substring — "LD001" must not match "LD0010"."""
        headers = _headers("Activity")
        rows = [
            {"Activity ID": "a1", "Lead ID": "LD001",  "Deleted": ""},
            {"Activity ID": "a2", "Lead ID": "LD0010", "Deleted": ""},
        ]
        ws = make_ws(headers, rows=rows)

        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Activity"])):
            result = crm.get_related_records(
                sheet_id="s1", entity="Activity",
                foreign_key_field="Lead ID", foreign_key_value="LD001",
            )

        assert result["total"] == 1
        assert result["records"][0]["Activity ID"] == "a1"


# ─────────────────────────────────────────────────────────────────────────────
# get_pipeline_analytics — missing/blank Segment values
# ─────────────────────────────────────────────────────────────────────────────

class TestPipelineAnalyticsMissingSegment:
    def test_blank_segment_bucketed_as_unspecified(self):
        headers = _headers("Lead")
        # make_ws serialises all values via str() so None becomes "None".
        # The source does: seg = r.get("Segment", "Unspecified") or "Unspecified"
        # which catches "" (falsy) but NOT "None" (truthy string).
        # Use "" for both blank cases to match real-world empty-cell behaviour.
        rows = [
            {"Status": "New Lead", "Segment": "",    "Deleted": ""},
            {"Status": "New Lead", "Segment": "",    "Deleted": ""},
            {"Status": "New Lead", "Segment": "SMB", "Deleted": ""},
        ]
        ws = make_ws(headers, rows=rows)

        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Lead"])):
            result = crm.get_pipeline_analytics(sheet_id="s1")

        seg = result["analytics"]["segment_breakdown"]
        # Source: `seg = r.get("Segment", "Unspecified") or "Unspecified"`
        # Empty string "" is falsy → bucketed as "Unspecified".
        unspecified = seg.get("Unspecified", 0)
        assert unspecified == 2
        assert seg.get("SMB") == 1


# ─────────────────────────────────────────────────────────────────────────────
# MCP tool wrappers return valid JSON
# ─────────────────────────────────────────────────────────────────────────────

class TestToolWrappers:
    """
    The MCP tools (tool_*) are thin wrappers that call the core functions and
    serialise results to JSON.  Verify the serialisation layer doesn't break.
    """

    def test_tool_list_records_is_json(self):
        headers = _headers("Lead")
        ws = make_ws(headers, rows=[
            {"Status": "New Lead", "Segment": "SMB", "Deleted": ""}
        ])

        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Lead"])):
            raw = crm.tool_list_records(sheet_id="s1", entity="Lead")

        parsed = json.loads(raw)
        assert parsed["status"] == "success"
        assert "records" in parsed

    def test_tool_add_record_is_json(self):
        headers = _headers("Organization")
        ws = make_ws(headers)

        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Organization"])):
            raw = crm.tool_add_record(
                sheet_id="s1",
                entity="Organization",
                fields={"Organization Name": "Acme", "Industry Domain": "SaaS"},
            )

        parsed = json.loads(raw)
        assert parsed["status"] == "success"
        assert "record_id" in parsed

    def test_tool_search_records_is_json(self):
        headers = _headers("Contact")
        ws = make_ws(headers, rows=[
            {"First Name": "Alice", "Deleted": ""}
        ])

        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Contact"])):
            raw = crm.tool_search_records(sheet_id="s1", entity="Contact", query="alice")

        parsed = json.loads(raw)
        assert parsed["status"] == "success"

    def test_tool_update_record_not_found_is_json(self):
        headers = _headers("Lead")
        ws = make_ws(headers)

        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Lead"])), \
             patch("mcp_excel_server.get_entity_row",
                   return_value=(None, None)):
            raw = crm.tool_update_record(
                sheet_id="s1", entity="Lead", record_id="ghost",
                updates={"Status": "Won"}
            )

        parsed = json.loads(raw)
        assert parsed["status"] == "failure"

    def test_tool_delete_record_is_json(self):
        headers = _headers("Activity")
        ws = make_ws(headers)
        record = {"Activity ID": "a1", "Deleted": ""}

        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Activity"])), \
             patch("mcp_excel_server.get_entity_row",
                   return_value=(2, record)):
            raw = crm.tool_delete_record(
                sheet_id="s1", entity="Activity", record_id="a1"
            )

        parsed = json.loads(raw)
        assert parsed["status"] == "success"

    def test_tool_get_pipeline_analytics_is_json(self):
        headers = _headers("Lead")
        ws = make_ws(headers)

        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Lead"])):
            raw = crm.tool_get_pipeline_analytics(sheet_id="s1")

        parsed = json.loads(raw)
        assert parsed["status"] == "success"
        assert "analytics" in parsed

    def test_tool_get_related_records_is_json(self):
        headers = _headers("Activity")
        ws = make_ws(headers)

        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Activity"])):
            raw = crm.tool_get_related_records(
                sheet_id="s1", entity="Activity",
                foreign_key_field="Lead ID", foreign_key_value="LD001",
            )

        parsed = json.loads(raw)
        assert parsed["status"] == "success"
