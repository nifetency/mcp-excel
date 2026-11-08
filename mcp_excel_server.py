"""
mcp_excel_server.py
Google Sheets-backed CRM -- MCP Server

Transport: FastMCP streamable-HTTP on port 5002 (default for container/server)
           Set MCP_TRANSPORT=stdio for Claude Desktop / local stdio use.

Schema-driven & dynamic:
  - ONE generic engine (ENTITY_SCHEMAS + generic CRUD functions).
  - Four entities: Lead, Organization, Contact, Activity.
  - Existing tabs are used as-is; missing default columns are appended.
  - New tabs are created fresh with the default schema.
  - Stateless per-call: every tool takes an explicit sheet_id.
  - New record IDs continue the existing pattern (LD001->LD002, ACT010->ACT011).
    Falls back to UUID if no consistent pattern is detectable.
"""

from __future__ import annotations

import asyncio
import datetime
import json
import os
import re
import uuid
from typing import Any, Dict, List, Optional, Tuple

import gspread
from google.oauth2.service_account import Credentials as SACredentials
from mcp.server.fastmcp import FastMCP
from mcp import types


# ---- Settings ----

class Settings:
    GOOGLE_APPLICATION_CREDENTIALS: str = os.environ.get(
        "GOOGLE_APPLICATION_CREDENTIALS", "creds.json"
    )
    SCOPES = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    AUDIT_HEADERS: List[str] = ["Created At", "Updated At", "Deleted"]
    PORT: int = int(os.environ.get("MCP_PORT", "5002"))
    TRANSPORT: str = os.environ.get("MCP_TRANSPORT", "http")


ENTITY_SCHEMAS: Dict[str, Dict[str, Any]] = {
    "Lead": {
        "tab": "Lead",
        "primary_key": "Lead ID",
        "headers": [
            "Lead ID", "Organization ID", "Contact Person ID", "Segment",
            "Opportunity Details", "Status", "Last Contact Date",
        ],
    },
    "Organization": {
        "tab": "Organization",
        "primary_key": "Organization ID",
        "headers": [
            "Organization ID", "Organization Name", "Industry Domain",
            "Sub Domain", "Company Type", "Website",
        ],
    },
    "Contact": {
        "tab": "Contact",
        "primary_key": "Contact ID",
        "headers": [
            "Contact ID", "First Name", "Last Name", "Organization ID",
            "Designation", "Mobile", "Alternate Mobile", "Email",
            "LinkedIn Profile", "Decision Maker", "Influencer",
            "Relationship Strength",
        ],
    },
    "Activity": {
        "tab": "Activity",
        "primary_key": "Activity ID",
        "headers": [
            "Activity ID", "Activity Date", "Lead ID", "Organization ID",
            "Contact ID", "Activity Type", "Activity Sub-Type", "Subject",
            "Discussion Summary", "Key Pain Points", "Action Items",
            "Follow-up Date", "Outcome",
        ],
    },
}


# ---- Utilities ----

def today() -> str:
    return datetime.date.today().isoformat()


def now_iso() -> str:
    return datetime.datetime.now(datetime.timezone.utc).strftime("%d-%m-%Y %H:%M")


def normalize(value: str) -> str:
    return (value or "").strip().lower()


def sanitize_cell(value: Any) -> str:
    if value is None:
        return ""
    s = str(value)
    if s and s[0] in ("=", "+", "-", "@"):
        return "'" + s
    return s


def extract_sheet_id(url_or_id: str) -> str:
    if not url_or_id or not url_or_id.strip():
        raise ValueError("sheet_id is required and was empty.")
    url_or_id = url_or_id.strip()
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", url_or_id)
    if match:
        return match.group(1)
    if re.match(r"^[a-zA-Z0-9_-]{20,}$", url_or_id):
        return url_or_id
    raise ValueError(f"Cannot extract sheet ID from: {url_or_id!r}")


def _normalize_key(s: str) -> str:
    """Normalize a field/header name for fuzzy matching: lowercase and strip
    everything but letters/digits. Lets callers pass slightly different
    spellings, spacing, or casing (e.g. 'contact_id' or 'Contact Id' for the
    header 'Contact ID') without the field being silently dropped."""
    return re.sub(r"[^a-z0-9]", "", (s or "").lower())


def _build_header_lookup(headers: List[str]) -> Dict[str, str]:
    """Map normalized header name -> actual header string, for fuzzy field matching."""
    return {_normalize_key(h): h for h in headers}


# Synonyms for common conceptual variants that _normalize_key alone can't
# catch (it only tolerates spacing/case/punctuation differences, not
# different wording for the same concept). Keys and values are both
# normalized (lowercase, letters/digits only). Used as a fallback when a
# caller-supplied field name doesn't match any real header directly -- e.g.
# an LLM sending "Phone Number" for a sheet whose actual column is "Mobile".
FIELD_SYNONYMS: Dict[str, str] = {
    "phone": "mobile",
    "phonenumber": "mobile",
    "phoneno": "mobile",
    "mobilenumber": "mobile",
    "mobileno": "mobile",
    "cell": "mobile",
    "cellphone": "mobile",
    "cellnumber": "mobile",
    "telephone": "mobile",
    "tel": "mobile",
    "contactnumber": "mobile",
    "primaryphone": "mobile",
    "alternatephone": "alternatemobile",
    "alternatephonenumber": "alternatemobile",
    "secondaryphone": "alternatemobile",
    "altmobile": "alternatemobile",
    "altphone": "alternatemobile",
}


def _resolve_field(header_lookup: Dict[str, str], field_name: str) -> Optional[str]:
    """Resolve a caller-supplied field name to a real header, tolerating
    both spelling variants (via normalization, already in header_lookup)
    and known conceptual synonyms (via FIELD_SYNONYMS). Returns None if
    nothing matches either way."""
    key = _normalize_key(field_name)
    matched = header_lookup.get(key)
    if matched:
        return matched
    canonical = FIELD_SYNONYMS.get(key)
    if canonical:
        return header_lookup.get(canonical)
    return None


def _resolve_header(headers: List[str], target: str) -> Optional[str]:
    """Fuzzy-resolve a canonical field name (e.g. 'Contact ID', 'Created At')
    to whatever the real header actually is on this sheet, tolerating case/
    spacing/wording differences -- the SAME tolerance add_record/update_record
    already apply to user-supplied field names. Used for the primary key and
    audit columns so an exact-spelling mismatch there doesn't silently break
    ID generation or timestamping the way `pk_field in headers` used to.
    Returns None if no column matches at all."""
    return _build_header_lookup(headers).get(_normalize_key(target))


def _col_letter(n: int) -> str:
    result = ""
    while n:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def _resolve_entity(entity: str) -> Dict[str, Any]:
    key_norm = normalize(entity).replace(" ", "")
    for key, schema in ENTITY_SCHEMAS.items():
        if normalize(key).replace(" ", "") == key_norm:
            return schema
    raise ValueError(
        f"Unknown entity {entity!r}. Valid entities: {list(ENTITY_SCHEMAS.keys())}"
    )


def _next_id(ws: gspread.Worksheet, headers: List[str], pk_field: str) -> str:
    """Continue the existing ID pattern (LD001->LD002 etc.), fallback to UUID.
    Reads only the primary-key column instead of all_values to minimise data transfer."""
    pk_header = _resolve_header(headers, pk_field)
    if not pk_header:
        return str(uuid.uuid4())

    pk_col = headers.index(pk_header) + 1  # 1-based for gspread
    col_letter = _col_letter(pk_col)
    # col_values returns a flat list including the header in row 1
    all_ids = ws.col_values(pk_col)[1:]  # skip header
    existing_ids = [v.strip() for v in all_ids if v.strip()]

    if not existing_ids:
        return str(uuid.uuid4())

    pat = re.compile(r"^([A-Za-z]*)([^A-Za-z0-9]?)([0-9]+)$")
    parsed = []
    for id_val in existing_ids:
        m = pat.match(id_val)
        if m:
            parsed.append((m.group(1), m.group(2), int(m.group(3)), len(m.group(3))))

    if not parsed:
        return str(uuid.uuid4())

    prefixes = {p[0] for p in parsed}
    seps = {p[1] for p in parsed}
    if len(prefixes) != 1 or len(seps) != 1:
        return str(uuid.uuid4())

    prefix = prefixes.pop()
    sep = seps.pop()
    max_num = max(p[2] for p in parsed)
    widths = [p[3] for p in parsed]
    pad_width = max(set(widths), key=widths.count)
    return f"{prefix}{sep}{str(max_num + 1).zfill(pad_width)}"


# ---- Google auth + connection cache ----
# Re-authenticating and re-opening the spreadsheet on every tool call
# costs 2 unnecessary API round trips. Cache the gspread Client and
# Spreadsheet objects per sheet_id for the process lifetime.
# gspread Clients hold an OAuth2 token that expires after 1 hour --
# the cache stores the expiry time and refreshes automatically.

import time as _time

_gc_cache: dict = {}          # {} -> (gspread.Client, expiry_epoch)
_ss_cache: dict = {}          # sheet_id -> gspread.Spreadsheet
_header_cache: dict = {}      # (sheet_id, entity) -> List[str]
_GC_TTL = 55 * 60            # refresh client 5 min before token expires


def _creds() -> SACredentials:
    return SACredentials.from_service_account_file(
        Settings.GOOGLE_APPLICATION_CREDENTIALS,
        scopes=Settings.SCOPES,
    )


def _service_account_email() -> str:
    with open(Settings.GOOGLE_APPLICATION_CREDENTIALS) as f:
        return json.load(f).get("client_email", "")


def _gc() -> gspread.Client:
    now = _time.monotonic()
    entry = _gc_cache.get("singleton")
    if entry and now < entry[1]:
        return entry[0]
    client = gspread.authorize(_creds())
    _gc_cache["singleton"] = (client, now + _GC_TTL)
    return client


def _open_spreadsheet(sheet_id: str) -> gspread.Spreadsheet:
    """Open a spreadsheet by ID, cached for the process lifetime.
    Cache is invalidated only if the sheet is not found (deleted/moved)."""
    if sheet_id in _ss_cache:
        return _ss_cache[sheet_id]
    ss = _gc().open_by_key(sheet_id)
    _ss_cache[sheet_id] = ss
    return ss


def _invalidate_sheet_cache(sheet_id: str) -> None:
    """Call after structural changes (new tab, new columns) so the next
    call re-reads headers rather than using a stale cached list."""
    _ss_cache.pop(sheet_id, None)
    keys_to_drop = [k for k in _header_cache if k[0] == sheet_id]
    for k in keys_to_drop:
        _header_cache.pop(k, None)


# ---- Dynamic tab + header reconciliation ----

def _ensure_entity_tab(
    spreadsheet: gspread.Spreadsheet, entity: str
) -> Tuple[gspread.Worksheet, List[str]]:
    schema = _resolve_entity(entity)
    tab_name = schema["tab"]
    default_headers = schema["headers"] + Settings.AUDIT_HEADERS

    try:
        ws = spreadsheet.worksheet(tab_name)
    except gspread.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(
            title=tab_name, rows=1000, cols=max(len(default_headers), 10)
        )
        ws.append_row(default_headers)
        return ws, default_headers

    existing_headers = ws.row_values(1)

    if not existing_headers:
        ws.append_row(default_headers)
        return ws, default_headers

    # Fuzzy-match against existing headers before deciding a schema column
    # is "missing". A plain exact-string check here would treat any
    # slightly-differently-named existing column (different case, spacing,
    # or wording -- e.g. 'Contact Id' vs 'Contact ID') as absent and create
    # a brand-new duplicate blank column for it on every single tool call,
    # silently forking future writes away from the real column.
    existing_lookup = _build_header_lookup(existing_headers)  # normalized -> actual header
    missing = [h for h in default_headers if _normalize_key(h) not in existing_lookup]
    if missing:
        start_col = len(existing_headers) + 1
        updates = [
            {"range": f"{_col_letter(start_col + i)}1", "values": [[h]]}
            for i, h in enumerate(missing)
        ]
        ws.batch_update(updates)
        existing_headers = ws.row_values(1)
        # Columns changed — drop header cache for this entity so next call re-reads
        _header_cache.pop((extract_sheet_id(spreadsheet.id), entity), None)

    return ws, existing_headers


def _open_entity(sheet_id: str, entity: str) -> Tuple[gspread.Worksheet, List[str], Dict[str, Any]]:
    sid = extract_sheet_id(sheet_id)
    cache_key = (sid, entity)

    # Check header cache first -- avoids re-reading row 1 on every call
    if cache_key in _header_cache:
        ss = _open_spreadsheet(sid)
        schema = _resolve_entity(entity)
        ws = ss.worksheet(schema["tab"])
        return ws, _header_cache[cache_key], schema

    ss = _open_spreadsheet(sid)
    ws, headers = _ensure_entity_tab(ss, entity)
    schema = _resolve_entity(entity)
    _header_cache[cache_key] = headers
    return ws, headers, schema


def get_all_records_batched(ws: gspread.Worksheet, headers: List[str]) -> List[Dict]:
    all_values = ws.get_all_values()
    if len(all_values) <= 1:
        return []
    deleted_header = _resolve_header(headers, "Deleted")
    rows = []
    for raw in all_values[1:]:
        padded = raw + [""] * (len(headers) - len(raw))
        record = dict(zip(headers, padded))
        if deleted_header and str(record.get(deleted_header, "")).lower() == "true":
            continue
        rows.append(record)
    return rows


def get_entity_row(
    ws: gspread.Worksheet, headers: List[str], pk_field: str, record_id: str
) -> Tuple[Optional[int], Optional[Dict]]:
    all_values = ws.get_all_values()
    if not all_values:
        return None, None
    pk_header = _resolve_header(headers, pk_field)
    id_col = headers.index(pk_header) if pk_header else 0
    for i, row in enumerate(all_values[1:], start=2):
        padded = row + [""] * (len(headers) - len(row))
        if padded[id_col] == record_id:
            return i, dict(zip(headers, padded))
    return None, None


# ---- Sheet validation ----

def get_sheet_info(sheet_id: str) -> Dict:
    try:
        sid = extract_sheet_id(sheet_id)
    except ValueError as e:
        return {"status": "failure", "message": str(e)}

    try:
        gc = _gc()
        ss = gc.open_by_key(sid)
        title = ss.title
    except gspread.exceptions.SpreadsheetNotFound:
        return {
            "status": "failure",
            "message": (
                f"Sheet not found. Share it with the service account "
                f"({_service_account_email()}) as Editor, then try again."
            ),
        }
    except gspread.exceptions.APIError as e:
        if "403" in str(e) or "PERMISSION_DENIED" in str(e):
            return {
                "status": "failure",
                "message": (
                    f"Permission denied on sheet. Please share the sheet with the "
                    f"service account ({_service_account_email()}) as Editor."
                ),
            }
        return {"status": "failure", "message": str(e)}
    except Exception as e:
        msg = str(e)
        if not msg:
            msg = (
                f"Sheet not found or access denied. Make sure the sheet is shared "
                f"with {_service_account_email()} as Editor."
            )
        return {"status": "failure", "message": msg}

    tabs_summary = {}
    try:
        for entity in ENTITY_SCHEMAS:
            ws, headers = _ensure_entity_tab(ss, entity)
            tabs_summary[entity] = headers
    except gspread.exceptions.APIError as e:
        if "403" in str(e) or "PERMISSION_DENIED" in str(e):
            return {
                "status": "failure",
                "message": (
                    f"Read access confirmed, but write access failed setting up tabs. "
                    f"Re-share with {_service_account_email()} as Editor."
                ),
            }
        return {"status": "failure", "message": str(e)}

    # Probe write access
    try:
        lead_ws = ss.worksheet("Lead")
        lead_headers = tabs_summary["Lead"]
        lead_ws.append_row([""] * len(lead_headers))
        all_values = lead_ws.get_all_values()
        if len(all_values) > 1:
            lead_ws.delete_rows(len(all_values))
    except gspread.exceptions.APIError as e:
        if "403" in str(e) or "PERMISSION_DENIED" in str(e):
            return {
                "status": "failure",
                "message": (
                    f"Sheet is shared as Viewer only. Re-share with "
                    f"{_service_account_email()} as Editor so records can be saved."
                ),
            }
        return {"status": "failure", "message": str(e)}
    except Exception as e:
        return {"status": "failure", "message": str(e)}

    return {
        "status": "success",
        "spreadsheet_id": sid,
        "sheet_title": title,
        "url": f"https://docs.google.com/spreadsheets/d/{sid}",
        "write_access_confirmed": True,
        "tabs": tabs_summary,
    }


# ---- Generic CRUD ----

def add_record(sheet_id: str, entity: str, fields: Dict[str, Any]) -> Dict:
    try:
        ws, headers, schema = _open_entity(sheet_id, entity)
    except (RuntimeError, ValueError) as e:
        return {"status": "failure", "message": str(e)}

    pk_field = schema["primary_key"]
    record_id = _next_id(ws, headers, pk_field)
    header_lookup = _build_header_lookup(headers)
    pk_header = header_lookup.get(_normalize_key(pk_field))
    created_header = header_lookup.get(_normalize_key("Created At"))
    updated_header = header_lookup.get(_normalize_key("Updated At"))
    skip_keys = {_normalize_key(pk_field), _normalize_key("Created At"),
                 _normalize_key("Updated At"), _normalize_key("Deleted")}

    row_data = [""] * len(headers)
    ignored = []
    for col, val in fields.items():
        if _normalize_key(col) in skip_keys:
            continue
        matched = _resolve_field(header_lookup, col)
        if matched:
            row_data[headers.index(matched)] = sanitize_cell(val)
        else:
            ignored.append(col)

    if pk_header:
        row_data[headers.index(pk_header)] = record_id
    if created_header:
        row_data[headers.index(created_header)] = now_iso()
    if updated_header:
        row_data[headers.index(updated_header)] = now_iso()

    ws.append_row(row_data)
    result = {"status": "success", "entity": schema["tab"], "record_id": record_id}
    if ignored:
        result["ignored_fields"] = ignored
    return result


def list_records(
    sheet_id: str,
    entity: str,
    filters: Optional[Dict[str, str]] = None,
    page: int = 1,
    page_size: int = 20,
) -> Dict:
    try:
        ws, headers, schema = _open_entity(sheet_id, entity)
    except (RuntimeError, ValueError) as e:
        return {"status": "failure", "message": str(e)}

    rows = get_all_records_batched(ws, headers)
    if filters:
        for field, value in filters.items():
            if field in headers:
                rows = [r for r in rows if normalize(r.get(field, "")) == normalize(value)]

    total = len(rows)
    start = (page - 1) * page_size
    return {
        "status": "success",
        "entity": schema["tab"],
        "total": total,
        "page": page,
        "page_size": page_size,
        "count": len(rows[start:start + page_size]),
        "records": rows[start:start + page_size],
    }


def search_records(
    sheet_id: str, entity: str, query: str, page: int = 1, page_size: int = 20
) -> Dict:
    try:
        ws, headers, schema = _open_entity(sheet_id, entity)
    except (RuntimeError, ValueError) as e:
        return {"status": "failure", "message": str(e)}

    rows = get_all_records_batched(ws, headers)
    q = normalize(query)
    matched = [r for r in rows if any(q in normalize(str(v)) for v in r.values())]

    total = len(matched)
    start = (page - 1) * page_size
    return {
        "status": "success",
        "entity": schema["tab"],
        "total": total,
        "page": page,
        "page_size": page_size,
        "count": len(matched[start:start + page_size]),
        "records": matched[start:start + page_size],
    }


def update_record(sheet_id: str, entity: str, record_id: str, updates: Dict[str, Any]) -> Dict:
    try:
        ws, headers, schema = _open_entity(sheet_id, entity)
    except (RuntimeError, ValueError) as e:
        return {"status": "failure", "message": str(e)}

    pk_field = schema["primary_key"]
    row_num, _ = get_entity_row(ws, headers, pk_field, record_id)
    if row_num is None:
        return {"status": "failure", "message": f"{schema['tab']} record {record_id!r} not found."}

    header_lookup = _build_header_lookup(headers)
    skip_keys = {_normalize_key(pk_field), _normalize_key("Created At")}
    batch = []
    ignored = []
    for field, value in updates.items():
        if _normalize_key(field) in skip_keys:
            continue
        matched = _resolve_field(header_lookup, field)
        if not matched:
            ignored.append(field)
            continue
        col = headers.index(matched) + 1
        batch.append({"range": f"{_col_letter(col)}{row_num}", "values": [[sanitize_cell(value)]]})

    updated_header = header_lookup.get(_normalize_key("Updated At"))
    if updated_header:
        ua_col = headers.index(updated_header) + 1
        batch.append({"range": f"{_col_letter(ua_col)}{row_num}", "values": [[now_iso()]]})

    if batch:
        ws.batch_update(batch)

    result = {"status": "success", "entity": schema["tab"], "record_id": record_id}
    if ignored:
        result["ignored_fields"] = ignored
    return result


def delete_record(sheet_id: str, entity: str, record_id: str) -> Dict:
    try:
        ws, headers, schema = _open_entity(sheet_id, entity)
    except (RuntimeError, ValueError) as e:
        return {"status": "failure", "message": str(e)}

    pk_field = schema["primary_key"]
    row_num, _ = get_entity_row(ws, headers, pk_field, record_id)
    if row_num is None:
        return {"status": "failure", "message": f"{schema['tab']} record {record_id!r} not found."}

    deleted_header = _resolve_header(headers, "Deleted")
    if not deleted_header:
        return {"status": "failure", "message": "Sheet has no 'Deleted' column for soft-delete."}

    del_col = headers.index(deleted_header) + 1
    ws.batch_update([{"range": f"{_col_letter(del_col)}{row_num}", "values": [["true"]]}])
    return {"status": "success", "entity": schema["tab"], "record_id": record_id}


def get_related_records(
    sheet_id: str,
    entity: str,
    foreign_key_field: str,
    foreign_key_value: str,
    page: int = 1,
    page_size: int = 20,
) -> Dict:
    try:
        ws, headers, schema = _open_entity(sheet_id, entity)
    except (RuntimeError, ValueError) as e:
        return {"status": "failure", "message": str(e)}

    if foreign_key_field not in headers:
        return {
            "status": "failure",
            "message": f"{foreign_key_field!r} is not a column on {schema['tab']}. "
                       f"Available columns: {headers}",
        }

    rows = get_all_records_batched(ws, headers)
    rows = [r for r in rows if r.get(foreign_key_field) == foreign_key_value]
    total = len(rows)
    start = (page - 1) * page_size
    return {"status": "success", "entity": schema["tab"], "total": total, "records": rows[start:start + page_size]}


def get_pipeline_analytics(sheet_id: str) -> Dict:
    try:
        ws, headers, schema = _open_entity(sheet_id, "Lead")
    except (RuntimeError, ValueError) as e:
        return {"status": "failure", "message": str(e)}

    rows = get_all_records_batched(ws, headers)
    status_breakdown: Dict[str, int] = {}
    segment_breakdown: Dict[str, int] = {}
    for r in rows:
        s = r.get("Status", "Unknown") or "Unknown"
        status_breakdown[s] = status_breakdown.get(s, 0) + 1
        seg = r.get("Segment", "Unspecified") or "Unspecified"
        segment_breakdown[seg] = segment_breakdown.get(seg, 0) + 1

    return {
        "status": "success",
        "analytics": {
            "total_leads": len(rows),
            "status_breakdown": status_breakdown,
            "segment_breakdown": segment_breakdown,
        },
    }


# ---- Duplicate-column cleanup ----
# One-time repair for sheets that already accumulated duplicate near-identical
# columns from before _ensure_entity_tab did fuzzy header matching (e.g. two
# columns that both mean "Contact ID" but were spelled slightly differently).
# Safe by default: dry_run=True only reports what it found. Pass
# dry_run=False to actually merge each group into its leftmost column and
# delete the rest.

def _group_duplicate_columns(headers: List[str]) -> Dict[str, List[int]]:
    """Return {normalized_key: [0-based col index, ...]} for any normalized
    header name that maps to 2+ actual columns."""
    groups: Dict[str, List[int]] = {}
    for i, h in enumerate(headers):
        if not h:
            continue
        key = _normalize_key(h)
        groups.setdefault(key, []).append(i)
    return {k: v for k, v in groups.items() if len(v) > 1}


def dedupe_entity_headers(sheet_id: str, entity: str, dry_run: bool = True) -> Dict:
    """
    Find and (optionally) fix duplicate near-identical columns on one entity
    tab. For each group of columns that normalize to the same header name:
      - merges every row's data into the leftmost column of the group
        (first non-empty value across the group wins, left to right),
      - deletes the other columns in that group (only when dry_run=False).
    Returns a report describing what was found / done, so callers (and the
    LLM) can show the user exactly what happened before or after applying.
    """
    try:
        sid = extract_sheet_id(sheet_id)
        gc = _gc()
        ss = gc.open_by_key(sid)
        schema = _resolve_entity(entity)
        ws = ss.worksheet(schema["tab"])
    except (RuntimeError, ValueError) as e:
        return {"status": "failure", "message": str(e)}
    except gspread.WorksheetNotFound:
        return {"status": "failure", "message": f"Tab {schema['tab']!r} not found."}

    all_values = ws.get_all_values()
    if not all_values:
        return {"status": "success", "entity": schema["tab"], "duplicate_groups": [], "message": "Tab is empty."}

    headers = all_values[0]
    data_rows = all_values[1:]
    dup_groups = _group_duplicate_columns(headers)

    if not dup_groups:
        return {
            "status": "success",
            "entity": schema["tab"],
            "duplicate_groups": [],
            "message": "No duplicate columns found.",
        }

    report = []
    for key, col_indices in dup_groups.items():
        cols_info = []
        for i in col_indices:
            non_empty = sum(1 for row in data_rows if i < len(row) and row[i].strip())
            cols_info.append({
                "column_letter": _col_letter(i + 1),
                "header": headers[i],
                "non_empty_cells": non_empty,
            })
        report.append({
            "normalized_key": key,
            "columns": cols_info,
            "kept_column": cols_info[0]["column_letter"],
            "deleted_columns": [c["column_letter"] for c in cols_info[1:]],
        })

    if dry_run:
        return {
            "status": "success",
            "entity": schema["tab"],
            "dry_run": True,
            "duplicate_groups": report,
            "message": (
                f"Found {len(dup_groups)} duplicate column group(s) on {schema['tab']}. "
                "Nothing changed -- call again with dry_run=false to merge and delete."
            ),
        }

    # ---- Apply: merge into leftmost column of each group, then delete the rest ----
    cols_to_delete = set()
    for key, col_indices in dup_groups.items():
        leftmost = col_indices[0]
        others = col_indices[1:]
        cols_to_delete.update(others)

        merged_col = []
        for row in data_rows:
            val = ""
            for i in col_indices:  # leftmost first -- first non-empty wins
                if i < len(row) and row[i].strip():
                    val = row[i]
                    break
            merged_col.append(val)

        col_letter = _col_letter(leftmost + 1)
        rng = f"{col_letter}2:{col_letter}{len(merged_col) + 1}"
        if merged_col:
            ws.update(rng, [[v] for v in merged_col])

    for col_idx in sorted(cols_to_delete, reverse=True):
        ws.delete_columns(col_idx + 1)  # gspread delete_columns is 1-based

    return {
        "status": "success",
        "entity": schema["tab"],
        "dry_run": False,
        "duplicate_groups": report,
        "message": (
            f"Merged and removed {len(cols_to_delete)} duplicate column(s) across "
            f"{len(dup_groups)} group(s) on {schema['tab']}."
        ),
    }


# ---- FastMCP server (HTTP by default, stdio optional) ----

mcp = FastMCP("mcp-excel", host="0.0.0.0", port=Settings.PORT, stateless_http=True, json_response=True)

_SHEET_ID_DESC = (
    "The org's connected Google Sheet -- full URL or bare spreadsheet ID. "
    "Always pass the sheet_id given to you in the system prompt; never omit it."
)
_ENTITY_DESC = "Which tab/entity to operate on: Lead, Organization, Contact, or Activity."


@mcp.tool()
def tool_get_sheet_info(sheet_id: str) -> str:
    """Validate access to a Google Sheet and ensure all 4 entity tabs exist."""
    return json.dumps(get_sheet_info(sheet_id))


@mcp.tool()
def tool_add_record(sheet_id: str, entity: str, fields: dict) -> str:
    """Add a new record. Primary key is auto-generated continuing the existing ID pattern."""
    return json.dumps(add_record(sheet_id, entity, fields))


@mcp.tool()
def tool_list_records(sheet_id: str, entity: str, filters: dict = None, page: int = 1, page_size: int = 20) -> str:
    """List records from an entity tab with optional column-based filters."""
    return json.dumps(list_records(sheet_id, entity, filters, page, page_size))


@mcp.tool()
def tool_search_records(sheet_id: str, entity: str, query: str, page: int = 1, page_size: int = 20) -> str:
    """Search an entity tab by keyword across all columns."""
    return json.dumps(search_records(sheet_id, entity, query, page, page_size))


@mcp.tool()
def tool_update_record(sheet_id: str, entity: str, record_id: str, updates: dict) -> str:
    """Update fields on an existing record by its primary-key ID."""
    return json.dumps(update_record(sheet_id, entity, record_id, updates))


@mcp.tool()
def tool_delete_record(sheet_id: str, entity: str, record_id: str) -> str:
    """Soft-delete a record (sets its Deleted column to true)."""
    return json.dumps(delete_record(sheet_id, entity, record_id))


@mcp.tool()
def tool_get_related_records(
    sheet_id: str, entity: str, foreign_key_field: str, foreign_key_value: str,
    page: int = 1, page_size: int = 20
) -> str:
    """Look up records on one entity related to another via a foreign key."""
    return json.dumps(get_related_records(sheet_id, entity, foreign_key_field, foreign_key_value, page, page_size))


@mcp.tool()
def tool_get_pipeline_analytics(sheet_id: str) -> str:
    """Get Lead pipeline analytics: total leads, status breakdown, segment breakdown."""
    return json.dumps(get_pipeline_analytics(sheet_id))


@mcp.tool()
def tool_dedupe_headers(sheet_id: str, entity: str, dry_run: bool = True) -> str:
    """
    One-time repair for duplicate near-identical columns on an entity tab
    (e.g. two columns that both mean 'Contact ID' but were spelled slightly
    differently, left over from before header matching was fuzzy). Always
    call with dry_run=true FIRST and show the user the report -- this is a
    destructive operation once applied (it deletes columns) and must never
    be run with dry_run=false without the user explicitly confirming after
    seeing the dry-run report.
    """
    return json.dumps(dedupe_entity_headers(sheet_id, entity, dry_run))


# ---- OCR: Business card / image extraction ----

def extract_business_card(image_base64: str, media_type: str = "image/jpeg") -> Dict:
    """
    Extract structured contact fields from a business card image (or any
    image containing contact details) using Claude Vision.

    Requires ANTHROPIC_API_KEY in the environment. Returns a dict with
    fields ready to pass directly to add_record(entity='Contact', fields=...).
    All fields are optional -- only those visibly present in the image are
    returned.
    """
    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        return {
            "status": "failure",
            "message": "ANTHROPIC_API_KEY not set -- cannot call Claude Vision for OCR.",
        }

    try:
        import anthropic as _anthropic
    except ImportError:
        return {
            "status": "failure",
            "message": "anthropic package not installed. Run: pip install anthropic",
        }

    client = _anthropic.Anthropic(api_key=api_key)

    extraction_prompt = (
        "Extract every contact detail visible in this image. "
        "Return ONLY a JSON object with these keys (omit any key whose value is not clearly visible):\n"
        "{\n"
        '  "First Name": string,\n'
        '  "Last Name": string,\n'
        '  "Designation": string,       // job title\n'
        '  "Organization Name": string, // company name\n'
        '  "Email": string,\n'
        '  "Mobile": string,            // primary phone\n'
        '  "Alternate Mobile": string,  // secondary phone if present\n'
        '  "LinkedIn Profile": string,  // full LinkedIn URL or handle\n'
        '  "Website": string,           // company or personal website\n'
        '  "Address": string            // full address if present\n'
        "}\n\n"
        "Return ONLY the JSON object, no explanation, no markdown fences."
    )

    try:
        response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=512,
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": media_type,
                                "data": image_base64,
                            },
                        },
                        {"type": "text", "text": extraction_prompt},
                    ],
                }
            ],
        )
    except Exception as e:
        return {"status": "failure", "message": f"Claude Vision API error: {e}"}

    raw_text = response.content[0].text.strip() if response.content else ""

    # Strip markdown fences if the model wrapped it anyway
    raw_text = re.sub(r"^```[a-z]*\n?", "", raw_text)
    raw_text = re.sub(r"\n?```$", "", raw_text)

    try:
        fields = json.loads(raw_text)
    except json.JSONDecodeError:
        return {
            "status": "failure",
            "message": "Claude returned non-JSON output.",
            "raw": raw_text,
        }

    # Sanitize every value
    fields = {k: sanitize_cell(v) for k, v in fields.items() if v}

    return {
        "status": "success",
        "fields": fields,
        "note": (
            "These fields are extracted but NOT yet saved. "
            "Search for an existing Contact/Lead first, then call add_record or "
            "update_record with these fields after the user confirms."
        ),
    }


@mcp.tool()
def tool_extract_business_card(image_base64: str, media_type: str = "image/jpeg") -> str:
    """
    Extract structured contact fields from a business card or contact-detail
    image using Claude Vision. Returns JSON with 'fields' dict ready to pass
    to tool_add_record(entity='Contact'). Does NOT save anything -- the LLM
    must search for an existing record first, confirm with the user, then
    call tool_add_record or tool_update_record.

    media_type: MIME type of the image (image/jpeg, image/png, image/webp, image/gif).
    image_base64: Raw base64-encoded image bytes (no data URI prefix).
    """
    return json.dumps(extract_business_card(image_base64, media_type))


if __name__ == "__main__":
    if Settings.TRANSPORT == "stdio":
        mcp.run(transport="stdio")
    else:
        mcp.run(transport="streamable-http")
