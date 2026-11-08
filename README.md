# MCP-Excel — Google Sheets CRM MCP Server

A stateless MCP server that turns any Google Sheet into a 4-entity CRM
(Lead, Organization, Contact, Activity). Runs as a standalone HTTP service
(FastMCP streamable-HTTP, default port 5002) or over stdio for Claude
Desktop and other local MCP hosts — no gateway, no ngrok, no subprocess
spawning required.

---

## Architecture

```
Any MCP host (Claude Desktop, your backend, an agent framework)
  │  HTTP POST /mcp  (JSON-RPC) — or stdio
  ▼
MCP-Excel :5002  (FastMCP streamable-HTTP)
  │  gspread
  ▼
Google Sheets API
```

Every tool call includes an explicit `sheet_id` — there is no session state,
no "active sheet" variable. Multiple teams can use the same running instance
safely because nothing is shared between requests.

---

## Sheet schema (auto-managed)

Four tabs, created on first connection if they don't exist. Existing tabs are
used as-is; missing columns are appended without disturbing existing data.

| Tab | Primary Key | Key columns |
|---|---|---|
| **Lead** | Lead ID | Organization ID, Status, Segment, Opportunity Details, Last Contact Date |
| **Organization** | Organization ID | Organization Name, Industry Domain, Sub Domain, Company Type, Website |
| **Contact** | Contact ID | First Name, Last Name, Organization ID, Designation, Mobile, Email, LinkedIn Profile |
| **Activity** | Activity ID | Lead ID, Organization ID, Contact ID, Activity Type, Activity Sub-Type, Subject, Discussion Summary, Outcome |

IDs auto-increment continuing the existing pattern (`LD001 → LD002`,
`ACT010 → ACT011`). Falls back to UUID if no consistent pattern is found.

Each tab also gets `Created At`, `Updated At`, `Deleted` audit columns
appended if they aren't already present.

Headers are matched fuzzily (synonyms and near-duplicates are resolved), so
an existing sheet with slightly different column names still works. Run
`tool_dedupe_headers` (dry-run first) to merge genuine duplicate columns.

---

## Setup

### 1. Google service account

You need a `creds.json` service account key with **Google Sheets API** and
**Google Drive API** enabled.

Place it in this directory (the default path) or set `GOOGLE_APPLICATION_CREDENTIALS`
to its full path.

> **Never commit `creds.json` to git.** It is excluded in `.gitignore`;
> keep it local or mount it as a secret.

### 2. Share the sheet

For each Google Sheet you want to connect: open the sheet → Share → add
the service account email as **Editor** → Done.

The service account email is printed at startup and returned by `get_sheet_info`
on permission errors. It is also readable from `creds.json`:
```bash
python3 -c "import json; print(json.load(open('creds.json'))['client_email'])"
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

Optional — business-card OCR (Claude Vision) support:
```bash
pip install anthropic
export ANTHROPIC_API_KEY=...
```

---

## Running

### Local dev (HTTP on :5002)

```bash
GOOGLE_APPLICATION_CREDENTIALS=creds.json python mcp_excel_server.py
```

Verify it's healthy:
```bash
curl -s -X POST http://localhost:5002/mcp \
  -H "Content-Type: application/json" \
  -d '{"jsonrpc":"2.0","id":1,"method":"tools/list","params":{}}' \
  | python3 -m json.tool
```

### Docker (staging / production)

```bash
docker build -t mcp-excel .
docker run -p 5002:5002 \
  -v /path/to/creds.json:/secrets/creds.json \
  -e GOOGLE_APPLICATION_CREDENTIALS=/secrets/creds.json \
  mcp-excel
```

### Claude Desktop (stdio transport)

Set `MCP_TRANSPORT=stdio` to use stdio instead of HTTP:

```json
{
  "mcpServers": {
    "crm": {
      "command": "python",
      "args": ["/path/to/mcp-excel/mcp_excel_server.py"],
      "env": {
        "GOOGLE_APPLICATION_CREDENTIALS": "/path/to/mcp-excel/creds.json",
        "MCP_TRANSPORT": "stdio"
      }
    }
  }
}
```

---

## Environment variables

| Variable | Default | Notes |
|---|---|---|
| `GOOGLE_APPLICATION_CREDENTIALS` | `creds.json` | Path to service account JSON |
| `MCP_TRANSPORT` | `http` | `http` for server mode, `stdio` for Claude Desktop |
| `MCP_PORT` | `5002` | HTTP listen port |
| `ANTHROPIC_API_KEY` | — | Only for `tool_extract_business_card` |

---

## Tools

All tools are stateless — every call requires `sheet_id` (full Google Sheets
URL or bare spreadsheet ID).

| Tool | Description |
|---|---|
| `tool_get_sheet_info` | Validate access, ensure all 4 tabs exist, confirm write access |
| `tool_add_record` | Add a record to Lead / Organization / Contact / Activity |
| `tool_list_records` | List records with optional column-based filters (paged) |
| `tool_search_records` | Keyword search across all columns |
| `tool_update_record` | Update fields on an existing record by primary-key ID |
| `tool_delete_record` | Soft-delete a record (sets `Deleted` to `true`) |
| `tool_get_related_records` | Look up records via foreign key (e.g. all Activities for a Lead) |
| `tool_get_pipeline_analytics` | Lead pipeline analytics: total, status breakdown, segment breakdown |
| `tool_dedupe_headers` | Find and merge duplicate columns (dry-run report first) |
| `tool_extract_business_card` | Extract contact fields from a business-card image via Claude Vision (needs `ANTHROPIC_API_KEY`) |

---

## Development

```bash
pip install pytest
pytest test_mcp_excel.py test_crm_local.py -v
```

Tests run fully offline (Google calls are mocked); no credentials needed.

---

## Troubleshooting

| Problem | Fix |
|---|---|
| `status: failure, message: ""` | Sheet not found — share it with the service account as Editor |
| `Permission denied on sheet` | Re-share with the service account as **Editor** (not Viewer) |
| `Cannot connect to MCP-Excel service` | MCP-Excel not running — start it on `:5002` |
| `sheet_id is required and was empty` | Always pass `sheet_id` — there is no implicit active sheet |
| Tab not found / wrong columns | `get_sheet_info` auto-creates missing tabs and appends missing columns |

---

## License

MIT — see [LICENSE](LICENSE). Built by [Nife](https://nife.io).
