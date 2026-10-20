# OpenText ALM Octane Test Exporter

Read-only Python exporter for OpenText ALM Octane. It retrieves manual and Gherkin tests, fetches each test's stored script, and writes one Excel workbook.

## Workbook tabs

- **Test Inventory** — one row per test, with available metadata and script retrieval status.
- **Manual Steps** — best-effort parsing of manual script lines into action and validation columns.
- **Gherkin Lines** — line-level breakdown of feature/scenario headers, tags, steps, examples tables, and other lines.
- **Raw Scripts** — the complete unmodified script text for every test, plus any script retrieval error.
- **Export Log** — run summary and script retrieval failures.

The raw script tab is the source of truth. Manual script formats and metadata fields can vary between Octane configurations, so the parsed step-level view is intentionally best-effort.

## Requirements

- Python 3.10 or later recommended.
- Network access to the Octane server.
- An Octane API access key (client ID and client secret).
- Permissions to view the workspace's tests and scripts.

## Setup on Windows (PowerShell)

```powershell
mkdir octane-test-exporter
cd octane-test-exporter
# Copy the files from this package into this directory.

py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
Copy-Item .env.example .env
notepad .env
```

Edit `.env` with your Octane server root URL, shared space ID, workspace ID, client ID and client secret.

If PowerShell blocks virtual environment activation, you can run the environment's Python directly instead:

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe .\octane_exporter.py
```

## Setup on Bash

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
pip install -r requirements.txt
cp .env.example .env
# Edit .env using your preferred editor.
python octane_exporter.py
```

## Configuration

| Variable | Required | Default | Meaning |
|---|---:|---|---|
| `OCTANE_BASE_URL` | Yes | — | Server root, e.g. `https://octane.example.com`; do not append the API workspace path |
| `OCTANE_SHARED_SPACE_ID` | Yes | — | Octane shared space ID |
| `OCTANE_WORKSPACE_ID` | Yes | — | Octane workspace ID |
| `OCTANE_CLIENT_ID` | Yes | — | API access key client ID |
| `OCTANE_CLIENT_SECRET` | Yes | — | API access key secret |
| `OCTANE_OUTPUT_FILE` | No | `octane_tests_export.xlsx` | Output workbook path |
| `OCTANE_PAGE_SIZE` | No | `200` | Collection page size, accepted range 1–2000 |
| `OCTANE_REQUEST_TIMEOUT` | No | `45` | HTTP timeout in seconds |
| `OCTANE_VERIFY_TLS` | No | `true` | Verify HTTPS certificates; keep enabled in normal use |
| `OCTANE_TEST_QUERY` | No | blank | Optional fallback Octane query expression for both test collections |
| `OCTANE_MANUAL_TEST_QUERY` | No | inherits `OCTANE_TEST_QUERY` | Optional query applied only to `/manual_tests` |
| `OCTANE_GHERKIN_TEST_QUERY` | No | inherits `OCTANE_TEST_QUERY` | Optional query applied only to `/gherkin_tests` |

Protect `.env`; never commit API secrets to Git. `.gitignore` excludes it.

## Find IDs and validate your API

The API URL convention is:

```text
https://<server>/api/shared_spaces/<shared_space_id>/workspaces/<workspace_id>
```

The exporter signs in at:

```text
POST https://<server>/authentication/sign_in
```

with JSON body:

```json
{
  "client_id": "<client-id>",
  "client_secret": "<client-secret>"
}
```

The exporter wraps each query expression in the double quotes required by Octane's query syntax (unless already quoted); `requests` URL-encodes the value.

Example `.env` filters:

```dotenv
# Export manual tests assigned to owner 54010
OCTANE_MANUAL_TEST_QUERY=owner EQ {id EQ 54010}

# Export Gherkin tests whose name contains Rule Engine
OCTANE_GHERKIN_TEST_QUERY=name CONTAINS ^Rule Engine^

# Optional common fallback if per-type filters are not set
OCTANE_TEST_QUERY=
```

A per-type variable takes precedence over `OCTANE_TEST_QUERY`. If a per-type variable is omitted, the common query is used as its fallback. To export all tests, leave all three query variables blank.

It then reads these collections:

```text
GET .../manual_tests?limit=200&offset=0
GET .../gherkin_tests?limit=200&offset=0
```

and retrieves each script using:

```text
GET .../tests/<test-id>/script
```

If your server uses a context path (for example, a path segment before `/api`), include that context path in `OCTANE_BASE_URL`.

Open the Octane REST API documentation/Swagger for your actual deployment to confirm resource names, available fields, and access permissions. The exporter requests common metadata fields; if your Octane version rejects one of those fields, edit `DEFAULT_TEST_FIELDS` in `octane_exporter.py` to remove or adjust it. You can also set `DEFAULT_TEST_FIELDS = ""` to let Octane return its default fields.

## Run

```powershell
python .\octane_exporter.py
```

Exit codes:
- `0` — completed without script retrieval errors.
- `1` — workbook created, but one or more scripts could not be retrieved.
- `2` — configuration problem.
- `3` — fatal Octane API/authentication/collection retrieval error.
- `4` — local file or export error.

## Error handling and safety

- Read-only: this tool uses GET requests for test retrieval; sign-in/sign-out are the only POST requests.
- Retries transient network errors, HTTP 429, and HTTP 5xx responses.
- Uses request timeouts and paginated collection retrieval.
- Reports missing scripts and per-test script retrieval errors in the workbook.
- Avoids disabling TLS verification by default.
- Escapes Excel formula-leading text from API data.
- Does not log the client secret.

## Important notes

1. The API key must have permission to read the selected shared space/workspace and test scripts.
2. The exporter uses `limit` and `offset` pagination and handles the common Octane collection response (`data`, `total_count`). Verify response details against your deployed version if your instance is customized.
3. Metadata relations may be returned as objects or arrays; the workbook stores a readable name/ID representation.
4. The **Manual Steps** sheet is best-effort parsing. The complete script for both manual and Gherkin tests is retained in **Raw Scripts**.
5. Gherkin is retained in full in **Raw Scripts** and represented line-by-line in **Gherkin Lines**. Scenario outlines and Examples tables are preserved.
6. If your Octane instance has a very large number of tests, consider running outside business hours and using `OCTANE_TEST_QUERY` to narrow the export. Confirm the exact query expression syntax for your Octane version before relying on a filter.

## Troubleshooting

- **401 at sign-in**: check the client ID/secret, server root URL, and API access key status.
- **403 on a collection or script**: check workspace access and read permissions.
- **404**: check the server root URL, shared space/workspace IDs, and API route.
- **400 mentioning `fields`**: remove unavailable fields from `DEFAULT_TEST_FIELDS` or request default fields.
- **TLS/certificate error**: install/configure your organization's trusted CA certificate. Do not turn off TLS verification as a routine fix.
- **Empty collection**: confirm the IDs and that the API key can see the relevant tests.
- **Scripts fail but tests list successfully**: verify script-read permissions and inspect the per-test error in the workbook.

## Suggested first run

Use a test workspace or a narrow query if you have one. Verify a few manual tests and Gherkin tests in Octane against the exported **Raw Scripts**, **Manual Steps**, and **Gherkin Lines** before relying on the workbook for audit or migration purposes.
