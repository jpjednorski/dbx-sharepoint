from __future__ import annotations


class TestPublicApi:
    def test_import_client(self):
        from dbx_sharepoint import SharePointClient
        assert SharePointClient is not None

    def test_import_template(self):
        from dbx_sharepoint import Template
        assert Template is not None

    def test_import_exceptions(self):
        from dbx_sharepoint import (
            SharePointError,
            SharePointAuthError,
            SharePointFileNotFoundError,
            SharePointPermissionError,
            SharePointThrottledError,
        )
        assert issubclass(SharePointAuthError, SharePointError)

    def test_import_shared_link(self):
        from dbx_sharepoint import read_excel_from_shared_link
        assert callable(read_excel_from_shared_link)

    def test_import_sensitivity_label(self):
        from dbx_sharepoint import (
            SensitivityLabel,
            apply_sensitivity_label,
            extract_sensitivity_label,
        )
        assert SensitivityLabel is not None
        assert callable(apply_sensitivity_label)
        assert callable(extract_sensitivity_label)
