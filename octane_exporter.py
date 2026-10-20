#!/usr/bin/env python3
"""
Read-only OpenText ALM Octane exporter.

Exports manual and Gherkin tests, their raw scripts, and a best-effort
step/scenario breakdown to one Excel workbook.

Configuration is read from environment variables or a local .env file.
"""

from __future__ import annotations

import json
import logging
import os
import re
import sys
import time
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import quote, urljoin

import requests
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from requests import Response, Session
from requests.exceptions import RequestException


LOGGER = logging.getLogger("octane_exporter")

DEFAULT_TEST_FIELDS = (
    "id,name,description,subtype,phase,owner,application_modules,"
    "product_areas,tags,creation_time,last_modified"
)


class ConfigurationError(Exception):
    """Raised for missing or invalid configuration."""


class OctaneAPIError(Exception):
    """Raised when an Octane API request fails."""


@dataclass(frozen=True)
class Config:
    base_url: str
    shared_space_id: str
    workspace_id: str
    client_id: str
    client_secret: str
    output_file: Path
    page_size: int = 200
    timeout: int = 45
    verify_tls: bool = True
    test_query: str = ""
    manual_test_query: str = ""
    gherkin_test_query: str = ""

    @property
    def workspace_url(self) -> str:
        return (
            f"{self.base_url}/api/shared_spaces/"
            f"{quote(self.shared_space_id, safe='')}/workspaces/"
            f"{quote(self.workspace_id, safe='')}"
        )


def env_bool(name: str, default: bool) -> bool:
    raw = os.getenv(name)
    if raw is None or not raw.strip():
        return default
    value = raw.strip().lower()
    if value in {"1", "true", "yes", "y", "on"}:
        return True
    if value in {"0", "false", "no", "n", "off"}:
        return False
    raise ConfigurationError(f"{name} must be true or false.")


def load_config() -> Config:
    load_dotenv()
    required = [
        "OCTANE_BASE_URL",
        "OCTANE_SHARED_SPACE_ID",
        "OCTANE_WORKSPACE_ID",
        "OCTANE_CLIENT_ID",
        "OCTANE_CLIENT_SECRET",
    ]
    missing = [name for name in required if not os.getenv(name, "").strip()]
    if missing:
        raise ConfigurationError(
            "Missing required environment variable(s): "
            + ", ".join(missing)
            + ". Copy .env.example to .env and configure it."
        )

    base_url = os.environ["OCTANE_BASE_URL"].strip().rstrip("/")
    if not base_url.lower().startswith(("https://", "http://")):
        raise ConfigurationError("OCTANE_BASE_URL must start with https:// or http://")

    try:
        page_size = int(os.getenv("OCTANE_PAGE_SIZE", "200"))
        timeout = int(os.getenv("OCTANE_REQUEST_TIMEOUT", "45"))
    except ValueError as exc:
        raise ConfigurationError(
            "OCTANE_PAGE_SIZE and OCTANE_REQUEST_TIMEOUT must be integers."
        ) from exc

    if not 1 <= page_size <= 2000:
        raise ConfigurationError("OCTANE_PAGE_SIZE must be between 1 and 2000.")
    if timeout < 1:
        raise ConfigurationError("OCTANE_REQUEST_TIMEOUT must be at least 1 second.")

    output_file = Path(
        os.getenv("OCTANE_OUTPUT_FILE", "octane_tests_export.xlsx")
    ).expanduser()

    return Config(
        base_url=base_url,
        shared_space_id=os.environ["OCTANE_SHARED_SPACE_ID"].strip(),
        workspace_id=os.environ["OCTANE_WORKSPACE_ID"].strip(),
        client_id=os.environ["OCTANE_CLIENT_ID"].strip(),
        client_secret=os.environ["OCTANE_CLIENT_SECRET"],
        output_file=output_file,
        page_size=page_size,
        timeout=timeout,
        verify_tls=env_bool("OCTANE_VERIFY_TLS", True),
        test_query=os.getenv("OCTANE_TEST_QUERY", "").strip(),
        manual_test_query=os.getenv(
            "OCTANE_MANUAL_TEST_QUERY",
            os.getenv("OCTANE_TEST_QUERY", ""),
        ).strip(),
        gherkin_test_query=os.getenv(
            "OCTANE_GHERKIN_TEST_QUERY",
            os.getenv("OCTANE_TEST_QUERY", ""),
        ).strip(),
    )


def json_safe(value: Any) -> str:
    """Render nested Octane metadata as readable JSON for Excel."""
    if value is None:
        return ""
    if isinstance(value, (dict, list)):
        return json.dumps(value, ensure_ascii=False, sort_keys=True)
    return str(value)


def unwrap_collection(payload: Any) -> tuple[list[dict[str, Any]], int | None]:
    """Handle the common Octane collection response shape."""
    if isinstance(payload, list):
        return [x for x in payload if isinstance(x, dict)], None
    if not isinstance(payload, dict):
        raise OctaneAPIError(
            f"Unexpected collection response type: {type(payload).__name__}"
        )

    rows = payload.get("data", [])
    if rows is None:
        rows = []
    if not isinstance(rows, list):
        raise OctaneAPIError("Collection response 'data' was not a list.")

    total = payload.get("total_count")
    if total is None:
        total = payload.get("totalCount")
    try:
        total_count = int(total) if total is not None else None
    except (TypeError, ValueError):
        total_count = None

    return [x for x in rows if isinstance(x, dict)], total_count


class OctaneClient:
    def __init__(self, config: Config):
        self.config = config
        self.session: Session = requests.Session()
        self.session.headers.update(
            {
                "Accept": "application/json",
                "Content-Type": "application/json",
            }
        )
        self._authenticated = False

    def close(self) -> None:
        if self._authenticated:
            try:
                self.session.post(
                    f"{self.config.base_url}/authentication/sign_out",
                    timeout=self.config.timeout,
                    verify=self.config.verify_tls,
                )
            except RequestException:
                LOGGER.debug("Sign-out failed; closing session anyway.", exc_info=True)
        self.session.close()

    def authenticate(self) -> None:
        url = f"{self.config.base_url}/authentication/sign_in"
        payload = {
            "client_id": self.config.client_id,
            "client_secret": self.config.client_secret,
        }
        try:
            response = self.session.post(
                url,
                json=payload,
                timeout=self.config.timeout,
                verify=self.config.verify_tls,
            )
        except RequestException as exc:
            raise OctaneAPIError(
                f"Could not connect to Octane sign-in endpoint: {exc}"
            ) from exc

        if response.status_code != 200:
            raise OctaneAPIError(
                self._format_http_error("Octane sign-in failed", response)
            )

        self._authenticated = True
        LOGGER.info("Authenticated with Octane API.")

    def _format_http_error(self, context: str, response: Response) -> str:
        detail = response.text.strip().replace("\n", " ")
        # Avoid accidentally logging secrets if an upstream proxy echoes request data.
        if len(detail) > 1200:
            detail = detail[:1200] + "…"
        return f"{context}: HTTP {response.status_code} {response.reason}. {detail}"

    def request_json(
        self,
        method: str,
        url: str,
        *,
        params: dict[str, Any] | None = None,
        retry_auth: bool = True,
    ) -> Any:
        for attempt in range(4):
            try:
                response = self.session.request(
                    method,
                    url,
                    params=params,
                    timeout=self.config.timeout,
                    verify=self.config.verify_tls,
                )
            except RequestException as exc:
                if attempt < 3:
                    delay = min(2 ** attempt, 8)
                    LOGGER.warning(
                        "Request network error (attempt %s/4); retrying in %ss: %s",
                        attempt + 1, delay, exc,
                    )
                    time.sleep(delay)
                    continue
                raise OctaneAPIError(f"Request failed for {url}: {exc}") from exc

            if response.status_code == 429 or 500 <= response.status_code <= 599:
                if attempt < 3:
                    retry_after = response.headers.get("Retry-After", "")
                    try:
                        delay = min(max(int(retry_after), 1), 30)
                    except ValueError:
                        delay = min(2 ** attempt, 8)
                    LOGGER.warning(
                        "Octane returned HTTP %s; retrying in %ss.",
                        response.status_code, delay,
                    )
                    time.sleep(delay)
                    continue

            if response.status_code in (401, 403):
                raise OctaneAPIError(
                    self._format_http_error(
                        "Authentication/permission error; check API key and workspace access",
                        response,
                    )
                )
            if not response.ok:
                raise OctaneAPIError(self._format_http_error("Octane API request failed", response))

            if not response.content:
                return {}
            try:
                return response.json()
            except ValueError as exc:
                raise OctaneAPIError(
                    f"Octane returned non-JSON content for {url} "
                    f"(HTTP {response.status_code}). Check server URL and proxy/login page."
                ) from exc

        raise OctaneAPIError(f"Request failed after retries: {url}")

    @staticmethod
    def normalize_query(query: str) -> str:
        """Wrap an Octane query statement in double quotes as required by Octane."""
        query = query.strip()
        if not query:
            return ""
        if len(query) >= 2 and query.startswith('"') and query.endswith('"'):
            return query
        return f'"{query}"'

    def get_collection(
        self,
        resource: str,
        *,
        fields: str | None = None,
        query: str = "",
    ) -> list[dict[str, Any]]:
        """Read a paginated Octane collection using limit/offset."""
        url = f"{self.config.workspace_url}/{resource}"
        collected: list[dict[str, Any]] = []
        offset = 0
        seen_ids: set[str] = set()

        request_fields = fields
        while True:
            params: dict[str, Any] = {
                "limit": self.config.page_size,
                "offset": offset,
            }
            if request_fields:
                params["fields"] = request_fields
            if query:
                params["query"] = self.normalize_query(query)

            try:
                payload = self.request_json("GET", url, params=params)
            except OctaneAPIError as exc:
                # Metadata fields can differ between Octane versions/customizations.
                # Retry this collection without an explicit field list if it was rejected.
                if request_fields and ("field" in str(exc).lower() or "400" in str(exc)):
                    LOGGER.warning(
                        "Explicit field list rejected for %s; retrying with server defaults.",
                        resource,
                    )
                    request_fields = None
                    params.pop("fields", None)
                    payload = self.request_json("GET", url, params=params)
                else:
                    raise
            batch, total_count = unwrap_collection(payload)

            if not batch:
                break

            # Defensive deduplication if the server ignores offsets or returns overlap.
            fresh: list[dict[str, Any]] = []
            for item in batch:
                item_id = item.get("id")
                key = str(item_id) if item_id is not None else ""
                if key and key in seen_ids:
                    continue
                if key:
                    seen_ids.add(key)
                fresh.append(item)

            collected.extend(fresh)
            LOGGER.info(
                "%s: retrieved %s records (unique total %s%s)",
                resource,
                len(batch),
                len(collected),
                f" / {total_count}" if total_count is not None else "",
            )

            # If all rows were duplicates, pagination may be ignored or unstable.
            if not fresh:
                LOGGER.warning(
                    "%s returned no new records at offset %s; stopping to avoid a loop.",
                    resource, offset,
                )
                break

            offset += len(batch)
            if total_count is not None and offset >= total_count:
                break
            if len(batch) < self.config.page_size:
                break

        return collected

    def get_test_details(self, test_id: Any) -> dict[str, Any]:
        """Fetch a full test record when a collection row only contains id/type."""
        url = (
            f"{self.config.workspace_url}/tests/"
            f"{quote(str(test_id), safe='')}"
        )
        payload = self.request_json(
            "GET",
            url,
            params={"fields": "id,name,subtype,description,phase,owner"},
        )
        if isinstance(payload, dict):
            if isinstance(payload.get("data"), list) and payload["data"]:
                first = payload["data"][0]
                if isinstance(first, dict):
                    return first
            return payload
        return {}

    def get_script(self, test_id: Any) -> tuple[str, str]:
        """Return (script text, status/error message)."""
        url = (
            f"{self.config.workspace_url}/tests/"
            f"{quote(str(test_id), safe='')}/script"
        )
        try:
            payload = self.request_json("GET", url)
            if isinstance(payload, dict):
                script = payload.get("script", "")
                if script is None:
                    script = ""
                if not isinstance(script, str):
                    script = json_safe(script)
                return script, ""
            return json_safe(payload), "Script response was not a JSON object."
        except OctaneAPIError as exc:
            LOGGER.warning("Could not retrieve script for test %s: %s", test_id, exc)
            return "", str(exc)


def extract_test_name(test: dict[str, Any]) -> str:
    """Read a test's display name from common Octane/API response shapes."""
    for key in ("name", "test_name", "testName", "title"):
        value = test.get(key)
        if isinstance(value, str) and value.strip():
            return value.strip()
        if isinstance(value, dict):
            nested = value.get("value") or value.get("name")
            if isinstance(nested, str) and nested.strip():
                return nested.strip()
    return ""


def relation_names(value: Any) -> str:
    """Format Octane entity references without assuming a single schema."""
    if value is None:
        return ""
    if isinstance(value, list):
        return "; ".join(
            str(item.get("name", item.get("id", json_safe(item))))
            if isinstance(item, dict) else str(item)
            for item in value
        )
    if isinstance(value, dict):
        return str(value.get("name", value.get("id", json_safe(value))))
    return str(value)


def test_inventory_row(test: dict[str, Any], test_type: str) -> dict[str, Any]:
    return {
        "Test ID": test.get("id", ""),
        "Test Name": extract_test_name(test),
        "Test Type": test_type,
        "Subtype": test.get("subtype", ""),
        "Description": test.get("description", ""),
        "Status / Phase": relation_names(test.get("phase")),
        "Owner": relation_names(test.get("owner")),
        "Application Modules": relation_names(test.get("application_modules")),
        "Product Areas": relation_names(test.get("product_areas")),
        "Tags": relation_names(test.get("tags")),
        "Created": test.get("creation_time", ""),
        "Last Modified": test.get("last_modified", ""),
        "Script Retrieved": "",
        "Script Error": "",
    }


def parse_manual_script(script: str) -> list[dict[str, str]]:
    """
    Best-effort parser for the text-based manual script representation.

    Octane manual scripts commonly use lines prefixed with '- ' for steps
    and '? ' for validation/expected-result lines. Preserve the full raw
    script separately because exact conventions can vary by version/config.
    """
    rows: list[dict[str, str]] = []
    current_step = 0
    current_action: list[str] = []
    current_expected: list[str] = []

    def flush() -> None:
        nonlocal current_step, current_action, current_expected
        if current_action or current_expected:
            current_step += 1
            rows.append({
                "Step Number": str(current_step),
                "Step Type": "Manual step",
                "Action / Instruction": "\n".join(current_action).strip(),
                "Expected Result / Validation": "\n".join(current_expected).strip(),
            })
            current_action = []
            current_expected = []

    for raw_line in script.splitlines():
        line = raw_line.rstrip()
        stripped = line.lstrip()

        if stripped.startswith("- "):
            flush()
            current_action.append(stripped[2:])
        elif stripped.startswith("?"):
            # Octane's script representation uses '?' for validation lines.
            value = stripped[1:]
            if value.startswith(" "):
                value = value[1:]
            current_expected.append(value)
        elif not stripped:
            if current_action or current_expected:
                # Preserve blank line within an ongoing step only when content follows.
                current_action.append("")
        else:
            # Continuation line: attach to expected text if validation has begun,
            # otherwise to action/instruction text.
            if current_expected and not current_action:
                current_expected.append(stripped)
            elif current_expected:
                current_expected.append(stripped)
            else:
                current_action.append(stripped)

    flush()

    # If script is non-empty but no known prefixes were detected, preserve it.
    if script.strip() and not rows:
        rows.append({
            "Step Number": "1",
            "Step Type": "Unparsed script",
            "Action / Instruction": script.strip(),
            "Expected Result / Validation": "",
        })

    return rows


def parse_gherkin_script(script: str) -> list[dict[str, str]]:
    """Extract a readable line-level view while preserving full script elsewhere."""
    rows: list[dict[str, str]] = []
    feature = ""
    scenario = ""
    step_number = 0

    for line_number, raw_line in enumerate(script.splitlines(), start=1):
        stripped = raw_line.strip()
        if not stripped or stripped.startswith("#Auto generated Octane revision tag"):
            continue

        if re.match(r"^(Feature|Rule)\s*:", stripped):
            feature = stripped
            scenario = ""
            rows.append({
                "Line Number": str(line_number),
                "Feature": feature,
                "Scenario": "",
                "Line Type": "Feature",
                "Keyword": stripped.split(":", 1)[0],
                "Text": stripped,
            })
        elif re.match(r"^(Background|Scenario Outline|Scenario|Examples)\s*:", stripped):
            line_type = stripped.split(":", 1)[0]
            scenario = stripped
            rows.append({
                "Line Number": str(line_number),
                "Feature": feature,
                "Scenario": scenario,
                "Line Type": line_type,
                "Keyword": line_type,
                "Text": stripped,
            })
        elif re.match(r"^@\S+", stripped):
            rows.append({
                "Line Number": str(line_number),
                "Feature": feature,
                "Scenario": scenario,
                "Line Type": "Tag",
                "Keyword": "@tag",
                "Text": stripped,
            })
        elif re.match(r"^(Given|When|Then|And|But|\*)\b", stripped):
            step_number += 1
            keyword = re.match(r"^(Given|When|Then|And|But|\*)\b", stripped).group(1)
            rows.append({
                "Line Number": str(line_number),
                "Feature": feature,
                "Scenario": scenario,
                "Line Type": "Step",
                "Keyword": keyword,
                "Text": stripped,
            })
        elif stripped.startswith("|"):
            rows.append({
                "Line Number": str(line_number),
                "Feature": feature,
                "Scenario": scenario,
                "Line Type": "Examples/Table",
                "Keyword": "",
                "Text": stripped,
            })
        else:
            rows.append({
                "Line Number": str(line_number),
                "Feature": feature,
                "Scenario": scenario,
                "Line Type": "Other",
                "Keyword": "",
                "Text": stripped,
            })
    return rows


def safe_cell(value: Any) -> Any:
    """Normalize values; formula-like strings are forced to text when written."""
    if value is None:
        return ""
    if isinstance(value, (dict, list)):
        value = json_safe(value)
    return value


def write_sheet(ws, headers: list[str], rows: list[dict[str, Any]]) -> None:
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(vertical="top", wrap_text=True)

    for row in rows:
        ws.append([safe_cell(row.get(header, "")) for header in headers])
        # Force potentially formula-leading API text to string cells without
        # adding apostrophes that would corrupt the raw script value.
        for cell in ws[ws.max_row]:
            if isinstance(cell.value, str) and cell.value.startswith(("=", "+", "-", "@")):
                cell.data_type = "s"

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for col_idx, header in enumerate(headers, start=1):
        values = [str(header)]
        # Limit width sampling for large exports.
        for row_idx in range(2, min(ws.max_row, 202) + 1):
            value = ws.cell(row_idx, col_idx).value
            if value is not None:
                values.append((str(value).splitlines() or [""])[0][:100])
        width = min(max(max(map(len, values), default=10) + 2, 12), 48)
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)


def create_workbook(
    config: Config,
    inventory: list[dict[str, Any]],
    manual_steps: list[dict[str, Any]],
    gherkin_lines: list[dict[str, Any]],
    gherkin_scripts: list[dict[str, Any]],
    log_rows: list[dict[str, Any]],
) -> Path:
    wb = Workbook()
    default = wb.active
    wb.remove(default)

    inventory_headers = [
        "Test ID", "Test Name", "Test Type", "Subtype", "Description",
        "Status / Phase", "Owner", "Application Modules", "Product Areas",
        "Tags", "Created", "Last Modified", "Script Retrieved", "Script Error",
    ]
    manual_headers = [
        "Test ID", "Test Name", "Step Number", "Step Type",
        "Action / Instruction", "Expected Result / Validation",
    ]
    gherkin_headers = [
        "Test ID", "Test Name", "Line Number", "Feature", "Scenario",
        "Line Type", "Keyword", "Text",
    ]
    script_headers = ["Test ID", "Test Name", "Test Type", "Script", "Script Error"]
    log_headers = ["Timestamp UTC", "Level", "Event", "Details"]

    ws = wb.create_sheet("Test Inventory")
    write_sheet(ws, inventory_headers, inventory)

    ws = wb.create_sheet("Manual Steps")
    write_sheet(ws, manual_headers, manual_steps)

    ws = wb.create_sheet("Gherkin Lines")
    write_sheet(ws, gherkin_headers, gherkin_lines)

    ws = wb.create_sheet("Raw Scripts")
    write_sheet(ws, script_headers, gherkin_scripts)

    ws = wb.create_sheet("Export Log")
    write_sheet(ws, log_headers, log_rows)

    output = config.output_file
    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)
    return output


def main() -> int:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
    )
    started = datetime.now(timezone.utc)
    log_rows: list[dict[str, Any]] = []

    def log_event(level: str, event: str, details: str = "") -> None:
        log_rows.append({
            "Timestamp UTC": datetime.now(timezone.utc).isoformat(timespec="seconds"),
            "Level": level,
            "Event": event,
            "Details": details,
        })

    try:
        config = load_config()
    except ConfigurationError as exc:
        LOGGER.error("%s", exc)
        return 2

    client = OctaneClient(config)
    inventory: list[dict[str, Any]] = []
    manual_steps: list[dict[str, Any]] = []
    gherkin_lines: list[dict[str, Any]] = []
    raw_scripts: list[dict[str, Any]] = []
    failed_scripts = 0

    try:
        client.authenticate()
        manual_query = config.manual_test_query
        gherkin_query = config.gherkin_test_query
        manual_tests = client.get_collection(
            "manual_tests", fields=DEFAULT_TEST_FIELDS, query=manual_query
        )
        gherkin_tests = client.get_collection(
            "gherkin_tests", fields=DEFAULT_TEST_FIELDS, query=gherkin_query
        )
        log_event("INFO", "Tests retrieved",
                  f"Manual: {len(manual_tests)}; Gherkin: {len(gherkin_tests)}")

        for test_type, tests in (
            ("Manual", manual_tests),
            ("Gherkin", gherkin_tests),
        ):
            for test in tests:
                test_id = test.get("id", "")
                name = extract_test_name(test)

                if test_id not in ("", None) and not name:
                    try:
                        details = client.get_test_details(test_id)
                        test.update({k: v for k, v in details.items() if v is not None})
                        name = extract_test_name(test)
                        if name:
                            LOGGER.info(
                                "Hydrated missing name for %s test %s via /tests/%s",
                                test_type, test_id, test_id,
                            )
                    except OctaneAPIError as exc:
                        LOGGER.warning(
                            "Could not hydrate %s test %s details: %s",
                            test_type, test_id, exc,
                        )

                inventory_row = test_inventory_row(test, test_type)
                if not name:
                    keys = ", ".join(sorted(str(key) for key in test.keys()))
                    LOGGER.warning(
                        "Test %s from %s still has no usable name after detail lookup. "
                        "Response fields: %s",
                        test_id or "<missing ID>", test_type, keys or "<none>",
                    )
                    log_event(
                        "WARNING",
                        "Test name missing after detail lookup",
                        f"{test_type} test ID {test_id or '<missing ID>'}; "
                        f"response fields: {keys or '<none>'}",
                    )

                if test_id in ("", None):
                    inventory_row["Script Retrieved"] = "No"
                    inventory_row["Script Error"] = "Test has no ID in API response."
                    inventory.append(inventory_row)
                    failed_scripts += 1
                    log_event("ERROR", "Test missing ID", str(name))
                    continue

                script, script_error = client.get_script(test_id)
                inventory_row["Script Retrieved"] = "Yes" if not script_error else "No"
                inventory_row["Script Error"] = script_error
                inventory.append(inventory_row)

                raw_scripts.append({
                    "Test ID": test_id,
                    "Test Name": name,
                    "Test Type": test_type,
                    "Script": script,
                    "Script Error": script_error,
                })

                if script_error:
                    failed_scripts += 1
                    log_event("ERROR", "Script retrieval failed",
                              f"{test_type} test {test_id} ({name}): {script_error}")
                    continue

                if test_type == "Manual":
                    for step in parse_manual_script(script):
                        manual_steps.append({
                            "Test ID": test_id,
                            "Test Name": name,
                            **step,
                        })
                else:
                    for line in parse_gherkin_script(script):
                        gherkin_lines.append({
                            "Test ID": test_id,
                            "Test Name": name,
                            **line,
                        })

        log_event(
            "INFO" if failed_scripts == 0 else "WARNING",
            "Export completed",
            f"Manual tests: {len(manual_tests)}; Gherkin tests: {len(gherkin_tests)}; "
            f"script failures: {failed_scripts}",
        )
        output = create_workbook(
            config, inventory, manual_steps, gherkin_lines, raw_scripts, log_rows
        )
        elapsed = (datetime.now(timezone.utc) - started).total_seconds()
        LOGGER.info("Workbook saved: %s", output.resolve())
        LOGGER.info(
            "Summary: %s manual tests, %s Gherkin tests, %s script retrieval failures "
            "(%.1f seconds).",
            len(manual_tests), len(gherkin_tests), failed_scripts, elapsed,
        )
        return 0 if failed_scripts == 0 else 1

    except OctaneAPIError as exc:
        LOGGER.error("%s", exc)
        log_event("ERROR", "Fatal API error", str(exc))
        # A workbook is only written after collection retrieval and processing;
        # on fatal collection/auth errors, don't create a misleading partial report.
        return 3
    except (OSError, ValueError) as exc:
        LOGGER.exception("Export failed: %s", exc)
        return 4
    finally:
        client.close()


if __name__ == "__main__":
    sys.exit(main())
