from __future__ import annotations

from typing import Optional


class SharePointError(Exception):
    """Base exception for all dbx-sharepoint errors."""


class SharePointAuthError(SharePointError):
    """Authentication or credential error."""


class SharePointFileNotFoundError(SharePointError):
    """The requested file or folder was not found on SharePoint."""


class SharePointPermissionError(SharePointError):
    """Insufficient permissions to access the resource."""


class SharePointThrottledError(SharePointError):
    """SharePoint API rate limit hit (HTTP 429)."""

    def __init__(self, message: str, retry_after: Optional[int] = None):
        super().__init__(message)
        self.retry_after = retry_after
