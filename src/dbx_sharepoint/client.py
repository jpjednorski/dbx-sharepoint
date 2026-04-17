from __future__ import annotations

import time
from typing import Any, Optional, Union

import pandas as pd
import requests

from dbx_sharepoint._url import detect_environment, parse_sharepoint_url
from dbx_sharepoint.auth import build_credential_from_databricks_secrets
from dbx_sharepoint.excel import (
    Template,
    dataframe_from_excel_bytes,
    dataframe_to_excel_bytes,
)
from dbx_sharepoint.exceptions import (
    SharePointAuthError,
    SharePointError,
    SharePointFileNotFoundError,
    SharePointPermissionError,
    SharePointThrottledError,
)


class SharePointClient:
    """Simplified client for SharePoint file operations via Microsoft Graph API.

    Supports Azure Gov and Commercial environments, auto-detected from the site URL.

    Args:
        credential: Any azure-identity TokenCredential (e.g., ClientSecretCredential).
        site_url: SharePoint site URL (e.g., "https://myorg.sharepoint.us/sites/Team").
        graph_endpoint: Optional override for the Graph API endpoint.
    """

    _MAX_RETRIES = 3

    def __init__(
        self,
        credential: Any,
        site_url: str,
        graph_endpoint: Optional[str] = None,
    ):
        self._credential = credential
        self._site_url = site_url

        parsed = parse_sharepoint_url(site_url)
        self._hostname = parsed.hostname
        self._site_name = parsed.site_name

        env = detect_environment(site_url)
        self._graph_endpoint = graph_endpoint or env.graph_endpoint
        self._graph_scope = env.graph_scope

        self._site_id: Optional[str] = None

    @classmethod
    def from_databricks_secrets(
        cls,
        dbutils: Any,
        scope: str,
        prefix: str = "",
        site_url: Optional[str] = None,
        graph_endpoint: Optional[str] = None,
    ) -> "SharePointClient":
        """Create a client using credentials from a Databricks secret scope.

        Expects these keys in the scope (with optional prefix):
            {prefix}-tenant-id, {prefix}-client-id, {prefix}-client-secret
            {prefix}-site-url (optional, overridden by site_url param)
        """
        key_prefix = f"{prefix}-" if prefix else ""
        resolved_site_url = site_url
        if not resolved_site_url:
            try:
                resolved_site_url = dbutils.secrets.get(
                    scope=scope, key=f"{key_prefix}site-url"
                )
            except Exception:
                pass
        if not resolved_site_url:
            raise ValueError(
                "site_url must be provided either as a parameter or as a "
                f"'{key_prefix}site-url' secret in scope '{scope}'."
            )

        env = detect_environment(resolved_site_url)
        credential = build_credential_from_databricks_secrets(
            dbutils=dbutils, scope=scope, prefix=prefix,
            authority=env.login_authority,
        )
        return cls(
            credential=credential,
            site_url=resolved_site_url,
            graph_endpoint=graph_endpoint,
        )

    def _get_token(self) -> str:
        """Acquire a Bearer token from the credential."""
        token = self._credential.get_token(self._graph_scope)
        return token.token

    def _headers(self) -> dict:
        return {"Authorization": f"Bearer {self._get_token()}"}

    def _get_site_id(self) -> str:
        """Resolve and cache the Graph site ID for this SharePoint site."""
        if self._site_id is not None:
            return self._site_id
        url = f"{self._graph_endpoint}/v1.0/sites/{self._hostname}:/sites/{self._site_name}"
        resp = self._request("GET", url)
        self._site_id = resp.json()["id"]
        return self._site_id

    def _request(
        self,
        method: str,
        url: str,
        **kwargs: Any,
    ) -> requests.Response:
        """Make an HTTP request with error handling and retry on 429."""
        kwargs.setdefault("timeout", 60)
        kwargs.setdefault("headers", self._headers())

        for attempt in range(self._MAX_RETRIES):
            resp = requests.request(method, url, **kwargs)

            if resp.status_code in (429, 503, 504):
                try:
                    retry_after = int(resp.headers.get("Retry-After", 2))
                except (ValueError, TypeError):
                    retry_after = 2
                if attempt < self._MAX_RETRIES - 1:
                    time.sleep(retry_after)
                    continue
                if resp.status_code == 429:
                    raise SharePointThrottledError(
                        f"Rate limited after {self._MAX_RETRIES} retries",
                        retry_after=retry_after,
                    )

            if resp.status_code == 401:
                raise SharePointAuthError(
                    f"Authentication failed: {resp.text}"
                )
            if resp.status_code == 403:
                raise SharePointPermissionError(
                    f"Permission denied: {resp.text}"
                )
            if resp.status_code == 404:
                raise SharePointFileNotFoundError(
                    f"Not found: {url}"
                )

            try:
                resp.raise_for_status()
            except requests.HTTPError as exc:
                raise SharePointError(
                    f"HTTP {resp.status_code}: {resp.text}"
                ) from exc
            return resp

        raise SharePointThrottledError("Max retries exceeded")

    def _resolve_path(self, url_or_path: str) -> str:
        """Convert a full SharePoint URL or relative path to a Graph drive path."""
        if url_or_path.startswith(("http://", "https://")):
            parsed = parse_sharepoint_url(url_or_path)
            path = parsed.file_path
        else:
            path = url_or_path
        return path if path == "/" else path.rstrip("/")

    def list_files(self, url_or_path: str) -> pd.DataFrame:
        """List files and folders at the given SharePoint location.

        Args:
            url_or_path: Full SharePoint URL or path relative to the site
                (e.g., "/Shared Documents/reports/").

        Returns:
            DataFrame with columns: name, path, size_bytes, modified_at,
            modified_by, is_folder.
        """
        file_path = self._resolve_path(url_or_path)
        site_id = self._get_site_id()
        url: Optional[str] = (
            f"{self._graph_endpoint}/v1.0/sites/{site_id}/drive/root:{file_path}:/children"
        )

        items: list = []
        while url:
            resp = self._request("GET", url)
            body = resp.json()
            items.extend(body.get("value", []))
            url = body.get("@odata.nextLink")

        rows = []
        for item in items:
            modified_by = ""
            if "lastModifiedBy" in item and "user" in item["lastModifiedBy"]:
                modified_by = item["lastModifiedBy"]["user"].get("displayName", "")

            rows.append({
                "name": item["name"],
                "path": item.get("webUrl", ""),
                "size_bytes": item.get("size", 0),
                "modified_at": item.get("lastModifiedDateTime", ""),
                "modified_by": modified_by,
                "is_folder": "folder" in item,
            })

        return pd.DataFrame(rows, columns=["name", "path", "size_bytes", "modified_at", "modified_by", "is_folder"])

    def download(self, url_or_path: str) -> bytes:
        """Download a file from SharePoint and return its contents as bytes.

        Args:
            url_or_path: Full SharePoint URL or path relative to the site.

        Returns:
            File contents as bytes.
        """
        file_path = self._resolve_path(url_or_path)
        site_id = self._get_site_id()
        url = f"{self._graph_endpoint}/v1.0/sites/{site_id}/drive/root:{file_path}:/content"

        resp = self._request("GET", url)
        return resp.content

    def upload(self, content: bytes, url_or_path: str) -> None:
        """Upload bytes to a file on SharePoint.

        Args:
            content: File contents as bytes.
            url_or_path: Full SharePoint URL or path relative to the site.
        """
        file_path = self._resolve_path(url_or_path)
        site_id = self._get_site_id()
        url = f"{self._graph_endpoint}/v1.0/sites/{site_id}/drive/root:{file_path}:/content"

        headers = self._headers()
        headers["Content-Type"] = "application/octet-stream"
        self._request("PUT", url, headers=headers, data=content)

    def read_excel(
        self,
        url_or_path: str,
        sheet_name: Optional[Union[str, int]] = None,
    ) -> pd.DataFrame:
        """Download an Excel file from SharePoint and return as a DataFrame.

        Args:
            url_or_path: Full SharePoint URL or path relative to the site.
            sheet_name: Sheet to read. Defaults to first sheet.

        Returns:
            DataFrame with the sheet data.
        """
        content = self.download(url_or_path)
        return dataframe_from_excel_bytes(content, sheet_name=sheet_name)

    def write_excel(
        self,
        df: pd.DataFrame,
        url_or_path: str,
        sheet_name: str = "Sheet1",
    ) -> None:
        """Write a DataFrame as a new .xlsx file to SharePoint.

        Args:
            df: The DataFrame to write.
            url_or_path: Full SharePoint URL or path for the destination file.
            sheet_name: Name of the sheet in the workbook.
        """
        xlsx_bytes = dataframe_to_excel_bytes(df, sheet_name=sheet_name)
        self.upload(xlsx_bytes, url_or_path)

    def open_template(self, url_or_path: str) -> Template:
        """Download an Excel template from SharePoint for editing.

        Args:
            url_or_path: Full SharePoint URL or path to the template file.

        Returns:
            A Template object that can be populated with data.
        """
        content = self.download(url_or_path)
        return Template(content)

    def save(self, template: Template, url_or_path: str) -> None:
        """Save a populated Template to SharePoint.

        Args:
            template: The Template object to save.
            url_or_path: Full SharePoint URL or path for the destination file.
        """
        self.upload(template.to_bytes(), url_or_path)
