"""dbx-sharepoint: Simplified SharePoint file and Excel interfaces for Azure Gov Databricks."""
from __future__ import annotations

from dbx_sharepoint.client import SharePointClient
from dbx_sharepoint.excel import (
    SensitivityLabel,
    Template,
    apply_sensitivity_label,
    dataframe_to_excel_bytes_from_template,
    extract_sensitivity_label,
)
from dbx_sharepoint.exceptions import (
    SharePointAuthError,
    SharePointError,
    SharePointFileNotFoundError,
    SharePointPermissionError,
    SharePointThrottledError,
)
from dbx_sharepoint.shared_link import read_excel_from_shared_link

__all__ = [
    "SharePointClient",
    "Template",
    "SensitivityLabel",
    "apply_sensitivity_label",
    "extract_sensitivity_label",
    "dataframe_to_excel_bytes_from_template",
    "SharePointError",
    "SharePointAuthError",
    "SharePointFileNotFoundError",
    "SharePointPermissionError",
    "SharePointThrottledError",
    "read_excel_from_shared_link",
]
