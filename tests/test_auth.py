from __future__ import annotations

from unittest.mock import MagicMock, patch

import pytest
from dbx_sharepoint.auth import build_credential_from_databricks_secrets


class TestBuildCredentialFromDatabricksSecrets:
    def test_with_prefix(self):
        mock_dbutils = MagicMock()
        mock_dbutils.secrets.get.side_effect = lambda scope, key: {
            "prod-tenant-id": "t123",
            "prod-client-id": "c456",
            "prod-client-secret": "s789",
        }[key]

        with patch("dbx_sharepoint.auth.ClientSecretCredential") as mock_cred:
            build_credential_from_databricks_secrets(
                dbutils=mock_dbutils,
                scope="sharepoint",
                prefix="prod",
            )
            mock_cred.assert_called_once_with(
                tenant_id="t123",
                client_id="c456",
                client_secret="s789",
            )

    def test_without_prefix(self):
        mock_dbutils = MagicMock()
        mock_dbutils.secrets.get.side_effect = lambda scope, key: {
            "tenant-id": "t123",
            "client-id": "c456",
            "client-secret": "s789",
        }[key]

        with patch("dbx_sharepoint.auth.ClientSecretCredential") as mock_cred:
            build_credential_from_databricks_secrets(
                dbutils=mock_dbutils,
                scope="sharepoint",
            )
            mock_cred.assert_called_once_with(
                tenant_id="t123",
                client_id="c456",
                client_secret="s789",
            )

    def test_authority_passed_when_provided(self):
        mock_dbutils = MagicMock()
        mock_dbutils.secrets.get.side_effect = lambda scope, key: {
            "tenant-id": "t123",
            "client-id": "c456",
            "client-secret": "s789",
        }[key]

        with patch("dbx_sharepoint.auth.ClientSecretCredential") as mock_cred:
            build_credential_from_databricks_secrets(
                dbutils=mock_dbutils,
                scope="sharepoint",
                authority="https://login.microsoftonline.us",
            )
            mock_cred.assert_called_once_with(
                tenant_id="t123",
                client_id="c456",
                client_secret="s789",
                authority="https://login.microsoftonline.us",
            )

    def test_missing_required_secret_raises(self):
        mock_dbutils = MagicMock()
        mock_dbutils.secrets.get.side_effect = Exception("Secret not found")

        with pytest.raises(Exception, match="Secret not found"):
            build_credential_from_databricks_secrets(
                dbutils=mock_dbutils,
                scope="sharepoint",
            )
