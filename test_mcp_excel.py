"""
test_mcp_excel.py
Comprehensive unit tests for mcp_excel_server.py

Covers:
  - Utilities: normalize, sanitize_cell, extract_sheet_id, _col_letter,
    _resolve_entity, _next_id
  - Entity schema invariants
  - _ensure_entity_tab: create, blank-bootstrap, existing-partial, full-match
  - get_all_records_batched + get_entity_row
  - add_record: success, ignored fields, system-field protection, formula injection
  - list_records: all, filter, pagination, unknown-column filter
  - search_records: match, no-match, case-insensitive, pagination
  - update_record: success, ignored fields, pk protection, not-found
  - delete_record: success, not-found, missing-Deleted-column
  - get_related_records: success, pagination, unknown FK field
  - get_pipeline_analytics: empty, populated breakdowns
  - get_sheet_info: permission denied, invalid id, success path

Run:
    pip install pytest pytest-mock gspread google-auth mcp
    pytest test_mcp_excel.py -v
"""

from __future__ import annotations

import json
import os
import re
from unittest.mock import MagicMock, call, patch

import pytest

os.environ.setdefault("GOOGLE_APPLICATION_CREDENTIALS", "/fake/creds.json")

import mcp_excel_server as crm


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def _headers(entity: str) -> list[str]:
    return crm.ENTITY_SCHEMAS[entity]["headers"] + crm.Settings.AUDIT_HEADERS


def make_ws(headers: list[str], rows: list[dict] | None = None) -> MagicMock:
    """Return a mock gspread.Worksheet pre-loaded with header + data rows."""
    rows = rows or []
    ws = MagicMock()
    ws.row_values.return_value = list(headers)
    all_vals = [list(headers)]
    for r in rows:
        all_vals.append([str(r.get(h, "")) for h in headers])
    ws.get_all_values.return_value = all_vals
    return ws


def make_ss(entity_ws_map: dict[str, MagicMock]) -> MagicMock:
    """Return a mock gspread.Spreadsheet that routes worksheet() calls."""
    ss = MagicMock()
    ss.title = "Test CRM"

    def _worksheet(name):
        if name in entity_ws_map:
            return entity_ws_map[name]
        raise crm.gspread.WorksheetNotFound()

    ss.worksheet.side_effect = _worksheet
    return ss


# ─────────────────────────────────────────────────────────────────────────────
# Utilities
# ─────────────────────────────────────────────────────────────────────────────

class TestNormalize:
    def test_strips_and_lowercases(self):
        assert crm.normalize("  Hello World  ") == "hello world"

    def test_empty_string(self):
        assert crm.normalize("") == ""

    def test_none_returns_empty(self):
        assert crm.normalize(None) == ""


class TestSanitizeCell:
    def test_none_returns_empty(self):
        assert crm.sanitize_cell(None) == ""

    def test_int_coerced_to_str(self):
        assert crm.sanitize_cell(42) == "42"

    def test_safe_string_unchanged(self):
        assert crm.sanitize_cell("Alice Smith") == "Alice Smith"

    @pytest.mark.parametrize("ch", ["=", "+", "-", "@"])
    def test_formula_injection_prefixed(self, ch):
        result = crm.sanitize_cell(f"{ch}MALICIOUS()")
        assert result.startswith("'"), f"Expected ' prefix for opener {ch!r}"

    def test_normal_email_unchanged(self):
        # email starts with a letter, not a formula char
        assert crm.sanitize_cell("alice@example.com") == "alice@example.com"


class TestExtractSheetId:
    def test_full_url(self):
        url = "https://docs.google.com/spreadsheets/d/abc123XYZ/edit?usp=sharing"
        assert crm.extract_sheet_id(url) == "abc123XYZ"

    def test_bare_id(self):
        sid = "abcdefghijklmnopqrstuvwx"   # 24 chars -- passes the regex
        assert crm.extract_sheet_id(sid) == sid

    def test_empty_raises(self):
        with pytest.raises(ValueError, match="required"):
            crm.extract_sheet_id("")

    def test_whitespace_only_raises(self):
        with pytest.raises(ValueError, match="required"):
            crm.extract_sheet_id("   ")

    def test_non_sheet_url_raises(self):
        with pytest.raises(ValueError):
            crm.extract_sheet_id("https://example.com/not-a-sheet")

    def test_url_with_fragment(self):
        url = "https://docs.google.com/spreadsheets/d/myId999/edit#gid=0"
        assert crm.extract_sheet_id(url) == "myId999"


class TestColLetter:
    def test_a(self):
        assert crm._col_letter(1) == "A"

    def test_z(self):
        assert crm._col_letter(26) == "Z"

    def test_aa(self):
        assert crm._col_letter(27) == "AA"

    def test_az(self):
        assert crm._col_letter(52) == "AZ"

    def test_ba(self):
        assert crm._col_letter(53) == "BA"


class TestResolveEntity:
    def test_exact_match(self):
        assert crm._resolve_entity("Lead")["tab"] == "Lead"

    def test_lowercase(self):
        assert crm._resolve_entity("lead")["tab"] == "Lead"

    def test_mixed_case(self):
        assert crm._resolve_entity("ORGANIZATION")["tab"] == "Organization"

    def test_contact(self):
        assert crm._resolve_entity("contact")["tab"] == "Contact"

    def test_activity(self):
        assert crm._resolve_entity("Activity")["tab"] == "Activity"

    def test_unknown_raises(self):
        with pytest.raises(ValueError, match="Unknown entity"):
            crm._resolve_entity("Widget")

    def test_partial_name_raises(self):
        with pytest.raises(ValueError):
            crm._resolve_entity("Lead Extra")


class TestNextId:
    """_next_id should continue the existing sequential pattern or fall back to uuid."""

    def _ws(self, headers, id_values):
        ws = MagicMock()
        all_vals = [headers]
        pk_idx = headers.index("Lead ID")
        for v in id_values:
            row = [""] * len(headers)
            row[pk_idx] = v
            all_vals.append(row)
        ws.get_all_values.return_value = all_vals
        return ws

    def test_continues_numeric_pattern(self):
        headers = _headers("Lead")
        ws = self._ws(headers, ["LD001", "LD002", "LD003"])
        result = crm._next_id(ws, headers, "Lead ID")
        assert result == "LD004"

    def test_pads_correctly(self):
        headers = _headers("Lead")
        ws = self._ws(headers, ["ACT009", "ACT010"])
        result = crm._next_id(ws, headers, "Lead ID")
        assert result == "ACT011"

    def test_empty_sheet_returns_uuid(self):
        headers = _headers("Lead")
        ws = MagicMock()
        ws.get_all_values.return_value = [headers]
        result = crm._next_id(ws, headers, "Lead ID")
        assert len(result) > 10

    def test_mixed_prefix_returns_uuid(self):
        headers = _headers("Lead")
        ws = self._ws(headers, ["LD001", "CX002"])
        result = crm._next_id(ws, headers, "Lead ID")
        import uuid as _uuid
        try:
            _uuid.UUID(result)
            is_uuid = True
        except ValueError:
            is_uuid = False
        assert is_uuid

    def test_pk_not_in_headers_returns_uuid(self):
        headers = ["Some Column", "Other Column"]
        ws = MagicMock()
        ws.get_all_values.return_value = [headers]
        result = crm._next_id(ws, headers, "Missing PK")
        assert len(result) > 5


# ─────────────────────────────────────────────────────────────────────────────
# Entity schema invariants
# ─────────────────────────────────────────────────────────────────────────────

class TestEntitySchemas:
    def test_all_four_entities_present(self):
        assert set(crm.ENTITY_SCHEMAS.keys()) == {"Lead", "Organization", "Contact", "Activity"}

    @pytest.mark.parametrize("entity", list(crm.ENTITY_SCHEMAS.keys()))
    def test_primary_key_in_headers(self, entity):
        schema = crm.ENTITY_SCHEMAS[entity]
        assert schema["primary_key"] in schema["headers"]

    @pytest.mark.parametrize("entity", list(crm.ENTITY_SCHEMAS.keys()))
    def test_has_tab_key(self, entity):
        assert "tab" in crm.ENTITY_SCHEMAS[entity]

    def test_audit_headers_defined(self):
        assert set(crm.Settings.AUDIT_HEADERS) == {"Created At", "Updated At", "Deleted"}


# ─────────────────────────────────────────────────────────────────────────────
# _ensure_entity_tab
# ─────────────────────────────────────────────────────────────────────────────

class TestEnsureEntityTab:
    def test_creates_missing_tab(self):
        ss = MagicMock()
        ss.worksheet.side_effect = crm.gspread.WorksheetNotFound()
        new_ws = MagicMock()
        ss.add_worksheet.return_value = new_ws

        ws, headers = crm._ensure_entity_tab(ss, "Lead")

        ss.add_worksheet.assert_called_once()
        new_ws.append_row.assert_called_once_with(_headers("Lead"))
        assert headers == _headers("Lead")

    def test_bootstraps_blank_tab(self):
        ss = MagicMock()
        blank_ws = MagicMock()
        blank_ws.row_values.return_value = []
        ss.worksheet.return_value = blank_ws

        ws, headers = crm._ensure_entity_tab(ss, "Organization")

        blank_ws.append_row.assert_called_once_with(_headers("Organization"))
        assert headers == _headers("Organization")

    def test_leaves_complete_headers_untouched(self):
        ss = MagicMock()
        complete = _headers("Contact")
        ws = MagicMock()
        ws.row_values.return_value = list(complete)
        ss.worksheet.return_value = ws

        result_ws, headers = crm._ensure_entity_tab(ss, "Contact")

        ws.batch_update.assert_not_called()
        assert headers == complete

    def test_appends_missing_columns_only(self):
        """A tab with a partial header set gets exactly the missing columns appended."""
        ss = MagicMock()
        present = ["Lead ID", "Status"]
        ws = MagicMock()
        updated = list(present) + [h for h in _headers("Lead") if h not in present]
        ws.row_values.side_effect = [present, updated]
        ss.worksheet.return_value = ws

        result_ws, headers = crm._ensure_entity_tab(ss, "Lead")

        ws.batch_update.assert_called_once()
        written = [u["values"][0][0] for u in ws.batch_update.call_args[0][0]]
        expected_missing = [h for h in _headers("Lead") if h not in present]
        assert written == expected_missing

    def test_preserves_extra_custom_columns(self):
        """Columns added by the user beyond the schema must not be removed."""
        ss = MagicMock()
        custom = _headers("Contact") + ["Referral Source", "Budget"]
        ws = MagicMock()
        ws.row_values.return_value = list(custom)
        ss.worksheet.return_value = ws

        result_ws, headers = crm._ensure_entity_tab(ss, "Contact")

        ws.batch_update.assert_not_called()
        assert "Referral Source" in headers
        assert "Budget" in headers

    def test_unknown_entity_raises(self):
        ss = MagicMock()
        with pytest.raises(ValueError):
            crm._ensure_entity_tab(ss, "NotARealEntity")


# ─────────────────────────────────────────────────────────────────────────────
# get_all_records_batched / get_entity_row
# ─────────────────────────────────────────────────────────────────────────────

class TestBatchedRecords:
    def test_skips_deleted_true(self):
        headers = ["Contact ID", "First Name", "Deleted"]
        ws = MagicMock()
        ws.get_all_values.return_value = [
            headers,
            ["c1", "Alice", ""],
            ["c2", "Bob",   "true"],
            ["c3", "Carol", "false"],
        ]
        rows = crm.get_all_records_batched(ws, headers)
        names = [r["First Name"] for r in rows]
        assert "Bob" not in names
        assert "Alice" in names
        assert "Carol" in names

    def test_empty_sheet_returns_empty(self):
        ws = MagicMock()
        ws.get_all_values.return_value = []
        assert crm.get_all_records_batched(ws, ["A", "B"]) == []

    def test_header_only_returns_empty(self):
        headers = ["Lead ID", "Status"]
        ws = MagicMock()
        ws.get_all_values.return_value = [headers]
        assert crm.get_all_records_batched(ws, headers) == []

    def test_pads_short_rows(self):
        headers = ["A", "B", "C", "Deleted"]
        ws = MagicMock()
        ws.get_all_values.return_value = [headers, ["val1"]]
        rows = crm.get_all_records_batched(ws, headers)
        assert rows[0]["B"] == ""
        assert rows[0]["C"] == ""

    def test_returns_all_columns_as_dict_keys(self):
        headers = _headers("Lead")
        ws = make_ws(headers, rows=[{"Lead ID": "LD001", "Status": "New Lead", "Deleted": ""}])
        rows = crm.get_all_records_batched(ws, headers)
        assert set(rows[0].keys()) == set(headers)

    def test_get_entity_row_found(self):
        headers = ["Lead ID", "Status", "Deleted"]
        ws = MagicMock()
        ws.get_all_values.return_value = [
            headers,
            ["LD001", "New Lead", ""],
            ["LD002", "Won",      ""],
        ]
        row_num, record = crm.get_entity_row(ws, headers, "Lead ID", "LD002")
        assert row_num == 3
        assert record["Status"] == "Won"

    def test_get_entity_row_not_found(self):
        headers = ["Lead ID", "Status"]
        ws = MagicMock()
        ws.get_all_values.return_value = [headers, ["LD001", "New Lead"]]
        row_num, record = crm.get_entity_row(ws, headers, "Lead ID", "ghost")
        assert row_num is None
        assert record is None

    def test_get_entity_row_empty_sheet(self):
        ws = MagicMock()
        ws.get_all_values.return_value = []
        row_num, record = crm.get_entity_row(ws, ["A"], "A", "x")
        assert row_num is None


# ─────────────────────────────────────────────────────────────────────────────
# add_record
# ─────────────────────────────────────────────────────────────────────────────

class TestAddRecord:
    def test_success_fills_known_fields(self):
        headers = _headers("Contact")
        ws = make_ws(headers)

        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Contact"])):
            result = crm.add_record(
                sheet_id="s1",
                entity="Contact",
                fields={
                    "First Name":  "Alice",
                    "Last Name":   "Smith",
                    "Email":       "alice@x.com",
                    "Designation": "VP Sales",
                },
            )

        assert result["status"] == "success"
        assert "record_id" in result
        assert "ignored_fields" not in result

        appended = ws.append_row.call_args[0][0]
        assert appended[headers.index("First Name")] == "Alice"
        assert appended[headers.index("Email")] == "alice@x.com"
        assert appended[headers.index("Contact ID")] != ""

    def test_reports_ignored_unknown_fields(self):
        headers = _headers("Lead")
        ws = make_ws(headers)

        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Lead"])):
            result = crm.add_record(
                sheet_id="s1",
                entity="Lead",
                fields={"Status": "New Lead", "MadeUpColumn": "xyz"},
            )

        assert result["status"] == "success"
        assert "MadeUpColumn" in result["ignored_fields"]

    def test_cannot_override_primary_key(self):
        headers = _headers("Lead")
        ws = make_ws(headers)

        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Lead"])):
            result = crm.add_record(
                sheet_id="s1",
                entity="Lead",
                fields={"Lead ID": "HACKED", "Status": "Qualified"},
            )

        appended = ws.append_row.call_args[0][0]
        assert appended[headers.index("Lead ID")] != "HACKED"
        assert appended[headers.index("Lead ID")] == result["record_id"]

    def test_cannot_override_created_at(self):
        headers = _headers("Lead")
        ws = make_ws(headers)

        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Lead"])):
            crm.add_record(
                sheet_id="s1",
                entity="Lead",
                fields={"Created At": "1970-01-01", "Status": "Qualified"},
            )

        appended = ws.append_row.call_args[0][0]
        created_val = appended[headers.index("Created At")]
        assert "1970" not in created_val

    def test_formula_injection_in_field_value_is_sanitized(self):
        headers = _headers("Contact")
        ws = make_ws(headers)

        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Contact"])):
            crm.add_record(
                sheet_id="s1",
                entity="Contact",
                fields={"First Name": "=EVIL()", "Last Name": "Smith"},
            )

        appended = ws.append_row.call_args[0][0]
        assert appended[headers.index("First Name")].startswith("'")

    def test_sets_created_at_and_updated_at(self):
        headers = _headers("Activity")
        ws = make_ws(headers)

        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Activity"])):
            crm.add_record(
                sheet_id="s1",
                entity="Activity",
                fields={"Activity Type": "Call", "Subject": "Demo"},
            )

        appended = ws.append_row.call_args[0][0]
        assert appended[headers.index("Created At")] != ""
        assert appended[headers.index("Updated At")] != ""

    def test_invalid_entity_returns_failure(self):
        result = crm.add_record(sheet_id="s1", entity="Ghost", fields={})
        assert result["status"] == "failure"

    def test_no_ignored_key_when_all_fields_valid(self):
        headers = _headers("Organization")
        ws = make_ws(headers)

        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Organization"])):
            result = crm.add_record(
                sheet_id="s1",
                entity="Organization",
                fields={"Organization Name": "Acme", "Industry Domain": "SaaS"},
            )

        assert "ignored_fields" not in result


# ─────────────────────────────────────────────────────────────────────────────
# list_records
# ─────────────────────────────────────────────────────────────────────────────

class TestListRecords:
    def _rows(self):
        return [
            {"Status": "Qualified",   "Segment": "Enterprise", "Deleted": ""},
            {"Status": "New Lead",    "Segment": "SMB",        "Deleted": ""},
            {"Status": "Qualified",   "Segment": "SMB",        "Deleted": ""},
            {"Status": "New Lead",    "Segment": "Enterprise", "Deleted": "true"},
        ]

    def test_returns_all_non_deleted(self):
        headers = _headers("Lead")
        ws = make_ws(headers, rows=self._rows())
        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Lead"])):
            result = crm.list_records(sheet_id="s1", entity="Lead")
        assert result["status"] == "success"
        assert result["total"] == 3

    def test_filter_by_status(self):
        headers = _headers("Lead")
        ws = make_ws(headers, rows=self._rows())
        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Lead"])):
            result = crm.list_records(sheet_id="s1", entity="Lead",
                                       filters={"Status": "Qualified"})
        assert result["total"] == 2

    def test_filter_case_insensitive(self):
        headers = _headers("Lead")
        ws = make_ws(headers, rows=self._rows())
        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Lead"])):
            result = crm.list_records(sheet_id="s1", entity="Lead",
                                       filters={"Status": "QUALIFIED"})
        assert result["total"] == 2

    def test_filter_unknown_column_is_silently_skipped(self):
        headers = _headers("Lead")
        ws = make_ws(headers, rows=self._rows())
        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Lead"])):
            result = crm.list_records(sheet_id="s1", entity="Lead",
                                       filters={"NoSuchColumn": "x"})
        assert result["total"] == 3

    def test_pagination_page_1(self):
        headers = _headers("Lead")
        rows = [{"Status": "New Lead", "Deleted": ""} for _ in range(25)]
        ws = make_ws(headers, rows=rows)
        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Lead"])):
            result = crm.list_records(sheet_id="s1", entity="Lead", page=1, page_size=10)
        assert result["count"] == 10
        assert result["total"] == 25
        assert result["page"] == 1

    def test_pagination_page_2(self):
        headers = _headers("Lead")
        rows = [{"Status": "New Lead", "Deleted": ""} for _ in range(25)]
        ws = make_ws(headers, rows=rows)
        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Lead"])):
            result = crm.list_records(sheet_id="s1", entity="Lead", page=2, page_size=10)
        assert result["count"] == 10

    def test_pagination_last_page_partial(self):
        headers = _headers("Lead")
        rows = [{"Status": "New Lead", "Deleted": ""} for _ in range(23)]
        ws = make_ws(headers, rows=rows)
        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Lead"])):
            result = crm.list_records(sheet_id="s1", entity="Lead", page=3, page_size=10)
        assert result["count"] == 3

    def test_empty_sheet_returns_zero(self):
        headers = _headers("Lead")
        ws = make_ws(headers)
        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Lead"])):
            result = crm.list_records(sheet_id="s1", entity="Lead")
        assert result["total"] == 0
        assert result["records"] == []

    def test_invalid_entity_returns_failure(self):
        result = crm.list_records(sheet_id="s1", entity="Ghost")
        assert result["status"] == "failure"


# ─────────────────────────────────────────────────────────────────────────────
# search_records
# ─────────────────────────────────────────────────────────────────────────────

class TestSearchRecords:
    def _rows(self):
        return [
            {"First Name": "Alice",  "Email": "alice@widget.co",  "Deleted": ""},
            {"First Name": "Bob",    "Email": "bob@gadget.com",   "Deleted": ""},
            {"First Name": "Carlos", "Email": "carlos@widget.co", "Deleted": ""},
        ]

    def test_match_by_first_name(self):
        headers = _headers("Contact")
        ws = make_ws(headers, rows=self._rows())
        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Contact"])):
            result = crm.search_records(sheet_id="s1", entity="Contact", query="alice")
        assert result["total"] == 1

    def test_match_by_email_domain(self):
        headers = _headers("Contact")
        ws = make_ws(headers, rows=self._rows())
        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Contact"])):
            result = crm.search_records(sheet_id="s1", entity="Contact", query="widget")
        assert result["total"] == 2

    def test_no_match(self):
        headers = _headers("Contact")
        ws = make_ws(headers, rows=self._rows())
        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Contact"])):
            result = crm.search_records(sheet_id="s1", entity="Contact", query="zzznomatch")
        assert result["total"] == 0

    def test_case_insensitive_query(self):
        headers = _headers("Contact")
        ws = make_ws(headers, rows=self._rows())
        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Contact"])):
            result = crm.search_records(sheet_id="s1", entity="Contact", query="ALICE")
        assert result["total"] == 1

    def test_pagination(self):
        headers = _headers("Contact")
        rows = [{"First Name": f"User{i}", "Email": f"u{i}@x.com", "Deleted": ""}
                for i in range(15)]
        ws = make_ws(headers, rows=rows)
        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Contact"])):
            result = crm.search_records(sheet_id="s1", entity="Contact",
                                         query="user", page=2, page_size=5)
        assert result["total"] == 15
        assert result["count"] == 5

    def test_deleted_records_excluded(self):
        headers = _headers("Contact")
        rows = self._rows() + [{"First Name": "Alice Ghost", "Email": "g@g.com",
                                 "Deleted": "true"}]
        ws = make_ws(headers, rows=rows)
        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Contact"])):
            result = crm.search_records(sheet_id="s1", entity="Contact", query="alice")
        assert result["total"] == 1

    def test_invalid_entity_returns_failure(self):
        result = crm.search_records(sheet_id="s1", entity="Ghost", query="x")
        assert result["status"] == "failure"


# ─────────────────────────────────────────────────────────────────────────────
# update_record
# ─────────────────────────────────────────────────────────────────────────────

class TestUpdateRecord:
    def _lead_record(self):
        return {"Lead ID": "LD001", "Status": "New Lead", "Deleted": ""}

    def test_success_writes_batch(self):
        headers = _headers("Lead")
        ws = make_ws(headers)

        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Lead"])), \
             patch("mcp_excel_server.get_entity_row",
                   return_value=(2, self._lead_record())):
            result = crm.update_record(sheet_id="s1", entity="Lead",
                                        record_id="LD001",
                                        updates={"Status": "Qualified"})

        assert result["status"] == "success"
        ws.batch_update.assert_called_once()

    def test_updated_at_is_set(self):
        headers = _headers("Lead")
        ws = make_ws(headers)
        captured = {}
        ws.batch_update.side_effect = lambda u: captured.update({"u": u})

        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Lead"])), \
             patch("mcp_excel_server.get_entity_row",
                   return_value=(2, self._lead_record())):
            crm.update_record(sheet_id="s1", entity="Lead",
                               record_id="LD001", updates={"Status": "Won"})

        ua_col = crm._col_letter(headers.index("Updated At") + 1)
        ranges = [u["range"] for u in captured["u"]]
        assert any(r.startswith(ua_col) for r in ranges)

    def test_unknown_fields_reported_as_ignored(self):
        headers = _headers("Lead")
        ws = make_ws(headers)

        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Lead"])), \
             patch("mcp_excel_server.get_entity_row",
                   return_value=(2, self._lead_record())):
            result = crm.update_record(sheet_id="s1", entity="Lead",
                                        record_id="LD001",
                                        updates={"NotAColumn": "x"})

        assert result["ignored_fields"] == ["NotAColumn"]

    def test_primary_key_cannot_be_overwritten(self):
        headers = _headers("Lead")
        ws = make_ws(headers)
        captured = {}
        ws.batch_update.side_effect = lambda u: captured.update({"u": u})

        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Lead"])), \
             patch("mcp_excel_server.get_entity_row",
                   return_value=(2, self._lead_record())):
            crm.update_record(sheet_id="s1", entity="Lead", record_id="LD001",
                               updates={"Lead ID": "HACKED", "Status": "Won"})

        pk_col = crm._col_letter(headers.index("Lead ID") + 1)
        ranges = [u["range"] for u in captured.get("u", [])]
        assert not any(r.startswith(pk_col + "2") for r in ranges)

    def test_created_at_cannot_be_overwritten(self):
        headers = _headers("Lead")
        ws = make_ws(headers)
        captured = {}
        ws.batch_update.side_effect = lambda u: captured.update({"u": u})

        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Lead"])), \
             patch("mcp_excel_server.get_entity_row",
                   return_value=(2, self._lead_record())):
            crm.update_record(sheet_id="s1", entity="Lead", record_id="LD001",
                               updates={"Created At": "1970-01-01", "Status": "Won"})

        ca_col = crm._col_letter(headers.index("Created At") + 1)
        ranges = [u["range"] for u in captured.get("u", [])]
        assert not any(r.startswith(ca_col + "2") for r in ranges)

    def test_record_not_found_returns_failure(self):
        headers = _headers("Lead")
        ws = make_ws(headers)

        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Lead"])), \
             patch("mcp_excel_server.get_entity_row",
                   return_value=(None, None)):
            result = crm.update_record(sheet_id="s1", entity="Lead",
                                        record_id="ghost", updates={"Status": "x"})

        assert result["status"] == "failure"
        assert "ghost" in result["message"]

    def test_invalid_entity_returns_failure(self):
        result = crm.update_record(sheet_id="s1", entity="Ghost",
                                    record_id="x", updates={})
        assert result["status"] == "failure"


# ─────────────────────────────────────────────────────────────────────────────
# delete_record
# ─────────────────────────────────────────────────────────────────────────────

class TestDeleteRecord:
    def test_soft_delete_sets_deleted_true(self):
        headers = _headers("Activity")
        ws = make_ws(headers)
        record = {"Activity ID": "ACT001", "Deleted": ""}

        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Activity"])), \
             patch("mcp_excel_server.get_entity_row",
                   return_value=(2, record)):
            result = crm.delete_record(sheet_id="s1", entity="Activity",
                                        record_id="ACT001")

        assert result["status"] == "success"
        ws.batch_update.assert_called_once()
        batch = ws.batch_update.call_args[0][0]
        assert any(u["values"] == [["true"]] for u in batch)

    def test_record_not_found_returns_failure(self):
        headers = _headers("Activity")
        ws = make_ws(headers)

        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Activity"])), \
             patch("mcp_excel_server.get_entity_row",
                   return_value=(None, None)):
            result = crm.delete_record(sheet_id="s1", entity="Activity",
                                        record_id="ghost")

        assert result["status"] == "failure"

    def test_missing_deleted_column_returns_failure(self):
        headers = ["Activity ID", "Subject"]
        ws = make_ws(headers)
        record = {"Activity ID": "ACT001", "Subject": "Call"}

        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Activity"])), \
             patch("mcp_excel_server.get_entity_row",
                   return_value=(2, record)):
            result = crm.delete_record(sheet_id="s1", entity="Activity",
                                        record_id="ACT001")

        assert result["status"] == "failure"
        assert "Deleted" in result["message"]

    def test_invalid_entity_returns_failure(self):
        result = crm.delete_record(sheet_id="s1", entity="Ghost", record_id="x")
        assert result["status"] == "failure"


# ─────────────────────────────────────────────────────────────────────────────
# get_related_records
# ─────────────────────────────────────────────────────────────────────────────

class TestGetRelatedRecords:
    def _activity_rows(self):
        return [
            {"Activity ID": "a1", "Lead ID": "LD001", "Contact ID": "c1", "Deleted": ""},
            {"Activity ID": "a2", "Lead ID": "LD002", "Contact ID": "c2", "Deleted": ""},
            {"Activity ID": "a3", "Lead ID": "LD001", "Contact ID": "c3", "Deleted": ""},
            {"Activity ID": "a4", "Lead ID": "LD001", "Contact ID": "c4", "Deleted": "true"},
        ]

    def test_returns_matching_rows(self):
        headers = _headers("Activity")
        ws = make_ws(headers, rows=self._activity_rows())

        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Activity"])):
            result = crm.get_related_records(
                sheet_id="s1", entity="Activity",
                foreign_key_field="Lead ID", foreign_key_value="LD001",
            )

        assert result["status"] == "success"
        assert result["total"] == 2

    def test_unknown_fk_field_returns_failure(self):
        headers = _headers("Activity")
        ws = make_ws(headers)

        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Activity"])):
            result = crm.get_related_records(
                sheet_id="s1", entity="Activity",
                foreign_key_field="NonExistentField", foreign_key_value="x",
            )

        assert result["status"] == "failure"
        assert "NonExistentField" in result["message"]

    def test_pagination(self):
        headers = _headers("Activity")
        rows = [{"Activity ID": f"a{i}", "Lead ID": "LD001", "Deleted": ""}
                for i in range(15)]
        ws = make_ws(headers, rows=rows)

        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Activity"])):
            result = crm.get_related_records(
                sheet_id="s1", entity="Activity",
                foreign_key_field="Lead ID", foreign_key_value="LD001",
                page=2, page_size=5,
            )

        assert result["total"] == 15
        assert len(result["records"]) == 5

    def test_no_matches_returns_empty_list(self):
        headers = _headers("Activity")
        ws = make_ws(headers, rows=self._activity_rows())

        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Activity"])):
            result = crm.get_related_records(
                sheet_id="s1", entity="Activity",
                foreign_key_field="Lead ID", foreign_key_value="LD999",
            )

        assert result["status"] == "success"
        assert result["total"] == 0
        assert result["records"] == []

    def test_invalid_entity_returns_failure(self):
        result = crm.get_related_records(
            sheet_id="s1", entity="Ghost",
            foreign_key_field="x", foreign_key_value="y",
        )
        assert result["status"] == "failure"


# ─────────────────────────────────────────────────────────────────────────────
# get_pipeline_analytics
# ─────────────────────────────────────────────────────────────────────────────

class TestPipelineAnalytics:
    def test_empty_sheet(self):
        headers = _headers("Lead")
        ws = make_ws(headers)

        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Lead"])):
            result = crm.get_pipeline_analytics(sheet_id="s1")

        assert result["status"] == "success"
        assert result["analytics"]["total_leads"] == 0
        assert result["analytics"]["status_breakdown"] == {}
        assert result["analytics"]["segment_breakdown"] == {}

    def test_status_breakdown(self):
        headers = _headers("Lead")
        rows = [
            {"Status": "New Lead",  "Segment": "Enterprise", "Deleted": ""},
            {"Status": "Qualified", "Segment": "SMB",        "Deleted": ""},
            {"Status": "Qualified", "Segment": "SMB",        "Deleted": ""},
            {"Status": "Won",       "Segment": "Enterprise", "Deleted": ""},
        ]
        ws = make_ws(headers, rows=rows)

        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Lead"])):
            result = crm.get_pipeline_analytics(sheet_id="s1")

        a = result["analytics"]
        assert a["total_leads"] == 4
        assert a["status_breakdown"]["Qualified"] == 2
        assert a["status_breakdown"]["New Lead"] == 1
        assert a["status_breakdown"]["Won"] == 1

    def test_segment_breakdown(self):
        headers = _headers("Lead")
        rows = [
            {"Status": "New Lead", "Segment": "Enterprise", "Deleted": ""},
            {"Status": "New Lead", "Segment": "SMB",        "Deleted": ""},
            {"Status": "New Lead", "Segment": "SMB",        "Deleted": ""},
        ]
        ws = make_ws(headers, rows=rows)

        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Lead"])):
            result = crm.get_pipeline_analytics(sheet_id="s1")

        assert result["analytics"]["segment_breakdown"]["SMB"] == 2

    def test_deleted_leads_excluded(self):
        headers = _headers("Lead")
        rows = [
            {"Status": "New Lead", "Segment": "SMB", "Deleted": ""},
            {"Status": "Won",      "Segment": "SMB", "Deleted": "true"},
        ]
        ws = make_ws(headers, rows=rows)

        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Lead"])):
            result = crm.get_pipeline_analytics(sheet_id="s1")

        assert result["analytics"]["total_leads"] == 1
        assert "Won" not in result["analytics"]["status_breakdown"]

    def test_missing_status_falls_back_to_unknown(self):
        headers = _headers("Lead")
        rows = [{"Status": "", "Segment": "SMB", "Deleted": ""}]
        ws = make_ws(headers, rows=rows)

        with patch("mcp_excel_server._open_entity",
                   return_value=(ws, headers, crm.ENTITY_SCHEMAS["Lead"])):
            result = crm.get_pipeline_analytics(sheet_id="s1")

        a = result["analytics"]
        assert "Unknown" in a["status_breakdown"] or "" in a["status_breakdown"]

    def test_invalid_entity_in_open_returns_failure(self):
        result = crm.get_pipeline_analytics(sheet_id="")
        assert result["status"] == "failure"


# ─────────────────────────────────────────────────────────────────────────────
# get_sheet_info
# ─────────────────────────────────────────────────────────────────────────────

class TestGetSheetInfo:
    def test_empty_sheet_id_returns_failure(self):
        result = crm.get_sheet_info("")
        assert result["status"] == "failure"

    @staticmethod
    def _api_error(status: int, message: str):
        """Build a real gspread APIError with the response object it expects."""
        resp = MagicMock()
        resp.status_code = status
        resp.text = json.dumps({"error": {"code": status, "message": message}})
        # APIError.__init__ calls response.json()["error"] -- must be a real dict
        resp.json.return_value = {"error": {"code": status, "message": message}}
        return crm.gspread.exceptions.APIError(resp)

    def test_spreadsheet_not_found(self):
        mock_client = MagicMock()
        mock_client.open_by_key.side_effect = crm.gspread.exceptions.SpreadsheetNotFound()
        with patch("mcp_excel_server._gc", return_value=mock_client), \
             patch("mcp_excel_server._service_account_email",
                   return_value="sa@project.iam.gserviceaccount.com"):
            result = crm.get_sheet_info("not-a-real-id-abcdefghijklmnop")

        assert result["status"] == "failure"
        assert (
            "service account" in result["message"].lower()
            or "not found" in result["message"].lower()
        )

    def test_permission_denied_403(self):
        mock_client = MagicMock()
        mock_client.open_by_key.side_effect = self._api_error(403, "PERMISSION_DENIED")
        with patch("mcp_excel_server._gc", return_value=mock_client), \
             patch("mcp_excel_server._service_account_email",
                   return_value="sa@project.iam.gserviceaccount.com"):
            result = crm.get_sheet_info("forbidden-sheet-id-12345678901")

        assert result["status"] == "failure"
        assert "permission" in result["message"].lower()

    def test_success_returns_all_four_tab_headers(self):
        ss = MagicMock()
        ss.title = "My CRM Sheet"

        lead_ws = make_ws(_headers("Lead"))
        lead_ws.get_all_values.return_value = [
            _headers("Lead"),
            [""] * len(_headers("Lead")),
        ]

        new_ws = MagicMock()
        new_ws.row_values.return_value = []  # blank -> bootstrapped
        new_ws.get_all_values.return_value = []  # no data rows

        def _worksheet(name):
            if name == "Lead":
                return lead_ws
            raise crm.gspread.WorksheetNotFound()

        ss.worksheet.side_effect = _worksheet
        ss.add_worksheet.return_value = new_ws

        mock_client = MagicMock()
        mock_client.open_by_key.return_value = ss
        with patch("mcp_excel_server._gc", return_value=mock_client), \
             patch("mcp_excel_server._service_account_email",
                   return_value="sa@project.iam.gserviceaccount.com"):
            result = crm.get_sheet_info("validsheetid1234567890")

        assert result["status"] == "success"
        assert result["sheet_title"] == "My CRM Sheet"
        assert result["write_access_confirmed"] is True
        assert set(result["tabs"].keys()) == {"Lead", "Organization", "Contact", "Activity"}

    def test_viewer_only_returns_failure(self):
        """A 403 during the write-probe (append_row) indicates Viewer-only access."""
        ss = MagicMock()
        ss.title = "Read-Only CRM"

        lead_ws = MagicMock()
        lead_ws.row_values.return_value = _headers("Lead")
        lead_ws.get_all_values.return_value = [_headers("Lead")]
        lead_ws.append_row.side_effect = self._api_error(403, "PERMISSION_DENIED")

        ss.worksheet.return_value = lead_ws
        ss.add_worksheet.return_value = MagicMock()

        mock_client = MagicMock()
        mock_client.open_by_key.return_value = ss
        with patch("mcp_excel_server._gc", return_value=mock_client), \
             patch("mcp_excel_server._service_account_email",
                   return_value="sa@project.iam.gserviceaccount.com"):
            result = crm.get_sheet_info("readonly-sheet-1234567890123")

        assert result["status"] == "failure"
        assert (
            "editor" in result["message"].lower()
            or "viewer" in result["message"].lower()
        )
