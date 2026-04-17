from __future__ import annotations

import pytest
from dbx_sharepoint._url import parse_sharepoint_url, detect_environment, Environment


class TestDetectEnvironment:
    def test_gov_from_sharepoint_us(self):
        env = detect_environment("https://myorg.sharepoint.us/sites/Team")
        assert env == Environment.GOV

    def test_commercial_from_sharepoint_com(self):
        env = detect_environment("https://myorg.sharepoint.com/sites/Team")
        assert env == Environment.COMMERCIAL

    def test_unknown_domain_raises(self):
        with pytest.raises(ValueError, match="Cannot detect environment"):
            detect_environment("https://example.com/sites/Team")


class TestParseSharePointUrl:
    def test_parse_full_file_url_gov(self):
        url = "https://myorg.sharepoint.us/sites/TeamSite/Shared Documents/reports/q1.xlsx"
        result = parse_sharepoint_url(url)
        assert result.hostname == "myorg.sharepoint.us"
        assert result.site_name == "TeamSite"
        assert result.file_path == "/Shared Documents/reports/q1.xlsx"

    def test_parse_full_file_url_commercial(self):
        url = "https://myorg.sharepoint.com/sites/TeamSite/Shared Documents/data.xlsx"
        result = parse_sharepoint_url(url)
        assert result.hostname == "myorg.sharepoint.com"
        assert result.site_name == "TeamSite"
        assert result.file_path == "/Shared Documents/data.xlsx"

    def test_parse_folder_url(self):
        url = "https://myorg.sharepoint.us/sites/TeamSite/Shared Documents/reports/"
        result = parse_sharepoint_url(url)
        assert result.site_name == "TeamSite"
        assert result.file_path == "/Shared Documents/reports/"

    def test_parse_url_with_encoded_spaces(self):
        url = "https://myorg.sharepoint.us/sites/TeamSite/Shared%20Documents/My%20Folder/file.xlsx"
        result = parse_sharepoint_url(url)
        assert result.file_path == "/Shared Documents/My Folder/file.xlsx"

    def test_parse_url_no_sites_raises(self):
        with pytest.raises(ValueError, match="Cannot parse"):
            parse_sharepoint_url("https://myorg.sharepoint.us/something/else")


class TestEnvironmentEndpoints:
    def test_gov_graph_endpoint(self):
        env = detect_environment("https://myorg.sharepoint.us/sites/Team")
        assert env.graph_endpoint == "https://graph.microsoft.us"
        assert env.login_authority == "https://login.microsoftonline.us"

    def test_commercial_graph_endpoint(self):
        env = detect_environment("https://myorg.sharepoint.com/sites/Team")
        assert env.graph_endpoint == "https://graph.microsoft.com"
        assert env.login_authority == "https://login.microsoftonline.com"
