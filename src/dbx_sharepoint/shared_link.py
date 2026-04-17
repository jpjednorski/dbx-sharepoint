from __future__ import annotations

from typing import Optional, Union
from urllib.parse import parse_qsl, urlencode, urlparse, urlunparse

import pandas as pd
import requests

from dbx_sharepoint.excel import dataframe_from_excel_bytes
from dbx_sharepoint.exceptions import SharePointError

_XLSX_MAGIC = b"PK\x03\x04"


def read_excel_from_shared_link(
    url: str,
    sheet_name: Optional[Union[str, int]] = None,
) -> pd.DataFrame:
    """Read an Excel file from a SharePoint shared link (no auth required).

    Works with "anyone with the link" URLs. Appends download=1 to trigger
    a direct file download from SharePoint.

    Args:
        url: SharePoint shared link URL.
        sheet_name: Sheet to read. Defaults to first sheet.

    Returns:
        DataFrame with the sheet data.

    Raises:
        SharePointError: If the link returned HTML (e.g., expired, requires
            login) instead of an Excel file.
    """
    parsed = urlparse(url)
    query = dict(parse_qsl(parsed.query, keep_blank_values=True))
    query["download"] = "1"
    download_url = urlunparse(parsed._replace(query=urlencode(query)))

    resp = requests.get(download_url, timeout=60)
    resp.raise_for_status()

    if not resp.content.startswith(_XLSX_MAGIC):
        raise SharePointError(
            "Shared link did not return an Excel file. This usually means the "
            "link has expired, requires sign-in, or points to a different file "
            "type. Verify the link works in an incognito browser window."
        )

    return dataframe_from_excel_bytes(resp.content, sheet_name=sheet_name)
