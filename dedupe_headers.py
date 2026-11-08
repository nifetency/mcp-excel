"""
dedupe_headers.py
One-time cleanup for duplicate near-identical columns created by the
_ensure_entity_tab exact-match bug (fixed in mcp_excel_server.py).

For each entity tab (Lead, Organization, Contact, Activity):
  1. Reads the real header row.
  2. Groups columns whose normalized name (lowercase, letters/digits only)
     matches -- these are duplicates of the same logical column.
  3. For each group with more than one column, merges every row's data into
     the LEFTMOST column (first non-empty value across the group wins),
     then deletes the other columns in that group.

SAFE BY DEFAULT: run with no arguments (or --dry-run) first. It will only
PRINT a report of what it found -- nothing is written or deleted. Once
you've reviewed the report and made a backup copy of the spreadsheet
(File > Make a copy in Google Sheets), re-run with --apply to actually
merge and delete.

Usage:
    python dedupe_headers.py <sheet_id_or_url>              # dry run
    python dedupe_headers.py <sheet_id_or_url> --apply       # apply changes
    python dedupe_headers.py <sheet_id_or_url> --apply --entity Contact
        (limit to one entity tab instead of all four)
"""

import sys
import argparse

import gspread
from google.oauth2.service_account import Credentials as SACredentials

from mcp_excel_server import (
    Settings,
    ENTITY_SCHEMAS,
    extract_sheet_id,
    _normalize_key,
    _col_letter,
)


def _creds():
    return SACredentials.from_service_account_file(
        Settings.GOOGLE_APPLICATION_CREDENTIALS,
        scopes=Settings.SCOPES,
    )


def _gc():
    return gspread.authorize(_creds())


def group_duplicate_columns(headers):
    """Return {normalized_key: [col_index, ...]} for keys with 2+ columns.
    col_index is 0-based, matching position in `headers`."""
    groups = {}
    for i, h in enumerate(headers):
        if not h:
            continue
        key = _normalize_key(h)
        groups.setdefault(key, []).append(i)
    return {k: v for k, v in groups.items() if len(v) > 1}


def process_tab(ss, entity, apply_changes):
    schema = ENTITY_SCHEMAS[entity]
    tab_name = schema["tab"]
    try:
        ws = ss.worksheet(tab_name)
    except gspread.WorksheetNotFound:
        print(f"[{entity}] tab '{tab_name}' not found -- skipping.")
        return

    all_values = ws.get_all_values()
    if not all_values:
        print(f"[{entity}] tab is empty -- skipping.")
        return

    headers = all_values[0]
    data_rows = all_values[1:]
    dup_groups = group_duplicate_columns(headers)

    print(f"\n[{entity}] tab: {tab_name}  ({len(headers)} columns, {len(data_rows)} data rows)")
    if not dup_groups:
        print("  No duplicate columns found. Nothing to do.")
        return

    for key, col_indices in dup_groups.items():
        col_names = [headers[i] for i in col_indices]
        col_letters = [_col_letter(i + 1) for i in col_indices]
        non_empty_counts = []
        for i in col_indices:
            count = sum(1 for row in data_rows if i < len(row) and row[i].strip())
            non_empty_counts.append(count)
        print(f"  DUPLICATE GROUP '{key}':")
        for name, letter, count in zip(col_names, col_letters, non_empty_counts):
            print(f"    - column {letter}: header={name!r}  non_empty_cells={count}")
        print(f"    -> would merge into leftmost column {col_letters[0]} "
              f"({col_names[0]!r}) and delete the other {len(col_indices) - 1} column(s).")

    if not apply_changes:
        print("\n  (dry run -- no changes made. Re-run with --apply to merge & delete.)")
        return

    # ---- Apply: merge data into leftmost column of each group ----
    print("\n  Applying merge...")
    merged_values = {}  # 0-based col index (leftmost of group) -> list of merged cell values
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
        merged_values[leftmost] = merged_col

    # Write merged values into the leftmost column of each group
    for col_idx, values in merged_values.items():
        col_letter = _col_letter(col_idx + 1)
        rng = f"{col_letter}2:{col_letter}{len(values) + 1}"
        ws.update(rng, [[v] for v in values])
        print(f"    Wrote merged data into column {col_letter}.")

    # Delete duplicate columns in descending order so indices don't shift
    for col_idx in sorted(cols_to_delete, reverse=True):
        sheet_col_num = col_idx + 1  # gspread delete_columns is 1-based
        ws.delete_columns(sheet_col_num)
        print(f"    Deleted column {_col_letter(sheet_col_num)}.")

    print(f"  [{entity}] done.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("sheet_id", help="Sheet URL or bare spreadsheet ID")
    parser.add_argument("--apply", action="store_true",
                         help="Actually merge and delete columns (default: dry run / report only)")
    parser.add_argument("--entity", choices=list(ENTITY_SCHEMAS.keys()),
                         help="Limit to a single entity tab (default: all four)")
    args = parser.parse_args()

    sid = extract_sheet_id(args.sheet_id)
    gc = _gc()
    ss = gc.open_by_key(sid)

    print(f"Spreadsheet: {ss.title}")
    if args.apply:
        print("MODE: APPLY -- this WILL modify your sheet. Make sure you have a backup copy.")
    else:
        print("MODE: DRY RUN -- report only, nothing will be changed.")

    entities = [args.entity] if args.entity else list(ENTITY_SCHEMAS.keys())
    for entity in entities:
        process_tab(ss, entity, apply_changes=args.apply)


if __name__ == "__main__":
    main()
