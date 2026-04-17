from __future__ import annotations

import enum
from dataclasses import dataclass
from urllib.parse import unquote, urlparse


class Environment(enum.Enum):
    GOV = "gov"
    COMMERCIAL = "commercial"

    @property
    def graph_endpoint(self) -> str:
        if self is Environment.GOV:
            return "https://graph.microsoft.us"
        return "https://graph.microsoft.com"

    @property
    def login_authority(self) -> str:
        if self is Environment.GOV:
            return "https://login.microsoftonline.us"
        return "https://login.microsoftonline.com"

    @property
    def graph_scope(self) -> str:
        return f"{self.graph_endpoint}/.default"


@dataclass(frozen=True)
class ParsedSharePointUrl:
    hostname: str
    site_name: str
    file_path: str


def detect_environment(site_url: str) -> Environment:
    """Auto-detect Azure Gov vs Commercial from a SharePoint URL."""
    parsed = urlparse(site_url)
    hostname = parsed.hostname or ""
    if hostname.endswith(".sharepoint.us"):
        return Environment.GOV
    if hostname.endswith(".sharepoint.com"):
        return Environment.COMMERCIAL
    raise ValueError(
        f"Cannot detect environment from '{hostname}'. "
        "Expected *.sharepoint.us (Gov) or *.sharepoint.com (Commercial). "
        "Pass graph_endpoint explicitly if using a non-standard domain."
    )


def parse_sharepoint_url(url: str) -> ParsedSharePointUrl:
    """Parse a full SharePoint URL into hostname, site name, and file path."""
    parsed = urlparse(url)
    hostname = parsed.hostname or ""
    path = unquote(parsed.path)

    parts = path.split("/")
    if len(parts) < 3 or parts[1] != "sites":
        raise ValueError(
            f"Cannot parse SharePoint URL: '{url}'. "
            "Expected format: https://{{host}}/sites/{{site_name}}/{{path}}"
        )

    site_name = parts[2]
    file_path = "/" + "/".join(parts[3:]) if len(parts) > 3 else "/"

    return ParsedSharePointUrl(
        hostname=hostname,
        site_name=site_name,
        file_path=file_path,
    )
