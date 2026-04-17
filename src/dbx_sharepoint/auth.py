from __future__ import annotations

from typing import Any, Dict, Optional

from azure.identity import ClientSecretCredential


def build_credential_from_databricks_secrets(
    dbutils: Any,
    scope: str,
    prefix: str = "",
    authority: Optional[str] = None,
) -> ClientSecretCredential:
    """Build an azure-identity credential from Databricks secret scope.

    Args:
        dbutils: The Databricks dbutils object.
        scope: Name of the Databricks secret scope.
        prefix: Optional prefix for secret key names. If "prod", looks for
            "prod-tenant-id", "prod-client-id", "prod-client-secret".
            If empty, looks for "tenant-id", "client-id", "client-secret".
        authority: Optional Azure AD authority URL. For Gov use
            "https://login.microsoftonline.us". Auto-detected from site_url
            when called via SharePointClient.from_databricks_secrets().

    Returns:
        A ClientSecretCredential suitable for Microsoft Graph.
    """
    key_prefix = f"{prefix}-" if prefix else ""

    tenant_id = dbutils.secrets.get(scope=scope, key=f"{key_prefix}tenant-id")
    client_id = dbutils.secrets.get(scope=scope, key=f"{key_prefix}client-id")
    client_secret = dbutils.secrets.get(scope=scope, key=f"{key_prefix}client-secret")

    kwargs: Dict[str, Any] = {
        "tenant_id": tenant_id,
        "client_id": client_id,
        "client_secret": client_secret,
    }
    if authority:
        kwargs["authority"] = authority

    return ClientSecretCredential(**kwargs)
