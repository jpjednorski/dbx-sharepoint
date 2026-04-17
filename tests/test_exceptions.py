from __future__ import annotations

from dbx_sharepoint.exceptions import (
    SharePointError,
    SharePointAuthError,
    SharePointFileNotFoundError,
    SharePointPermissionError,
    SharePointThrottledError,
)


class TestExceptionHierarchy:
    def test_all_exceptions_inherit_from_base(self):
        for exc_class in [
            SharePointAuthError,
            SharePointFileNotFoundError,
            SharePointPermissionError,
            SharePointThrottledError,
        ]:
            assert issubclass(exc_class, SharePointError)
            assert issubclass(exc_class, Exception)

    def test_auth_error_message(self):
        err = SharePointAuthError("Token acquisition failed")
        assert str(err) == "Token acquisition failed"

    def test_file_not_found_with_path(self):
        err = SharePointFileNotFoundError("/Shared Documents/missing.xlsx")
        assert "/Shared Documents/missing.xlsx" in str(err)

    def test_throttled_error_with_retry_after(self):
        err = SharePointThrottledError("Rate limited", retry_after=30)
        assert err.retry_after == 30

    def test_throttled_error_default_retry_after(self):
        err = SharePointThrottledError("Rate limited")
        assert err.retry_after is None
