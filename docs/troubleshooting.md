# Troubleshooting

This page maps the most common errors to their root causes and fixes.

All errors raised by the library inherit from `SharePointError`. Status-specific subclasses allow targeted handling of known cases:

| Exception | HTTP | Typical cause |
|---|---|---|
| `SharePointAuthError` | 401 | Bad tenant/client/secret, wrong authority for cloud, expired secret |
| `SharePointPermissionError` | 403 | App lacks permission for this site or role |
| `SharePointFileNotFoundError` | 404 | Path is wrong, site-relative vs. absolute mismatch, file doesn't exist |
| `SharePointThrottledError` | 429 | Rate limited after 3 retries; carries `retry_after` |
| `SharePointError` | other | Anything else — message contains the Graph error body |

```python
from dbx_sharepoint import (
    SharePointAuthError,
    SharePointPermissionError,
    SharePointFileNotFoundError,
    SharePointThrottledError,
    SharePointError,
)

try:
    df = sp.read_excel("/Shared Documents/data.xlsx")
except SharePointFileNotFoundError:
    log.warning("data.xlsx missing — skipping")
except SharePointThrottledError as e:
    log.error("throttled after retries; retry after %s seconds", e.retry_after)
```

---

## Common errors

### `SharePointAuthError: Authentication failed: ...`

The OAuth flow failed. Check the following in order:

1. Correct authority for the cloud. Gov tenants require `authority="https://login.microsoftonline.us"` on the `ClientSecretCredential`. The `from_databricks_secrets` factory handles this automatically. When constructing the credential directly, ensure the authority is passed. See [azure-gov.md](azure-gov.md).
2. Tenant ID is the directory GUID, not the tenant name or domain.
3. Client ID is the Application (client) ID from the app registration overview page.
4. Client secret is the Value column of the secret, shown once at creation time. Using the Secret ID instead is a common mistake, since the two fields look similar.
5. Secret expired. Azure AD client secrets expire. Check the app registration's Certificates & secrets page.
6. Admin consent missing. If Graph permissions were added but admin consent was not granted, the token request succeeds but no Graph permissions are attached. Expect 401 or 403 depending on the endpoint.

### `SharePointPermissionError: Permission denied: ...`

Authentication succeeded, but the app is not authorized for the requested operation.

1. With `Sites.Selected`: confirm the app has been explicitly granted access to this site. `Sites.Selected` is empty until per-site grants are made. See [authentication.md](authentication.md#option-a-sitesselected-recommended) for the grant call.
2. With `Sites.Read.All` or `Files.Read.All`: reads work, but writes fail with 403. Upgrade to the corresponding read-write permission.
3. Wrong tenant's app registration. Separate Gov and Commercial tenants have separate app registrations. A Commercial app cannot access Gov sites, and vice versa.

### `SharePointFileNotFoundError: Not found: ...`

1. Path format. Paths are accepted either as site-relative (`/Shared Documents/report.xlsx`) or as full URLs. For a site at `https://myorg.sharepoint.us/sites/TeamSite`, both of these resolve identically:
   - `/Shared Documents/report.xlsx`
   - `https://myorg.sharepoint.us/sites/TeamSite/Shared Documents/report.xlsx`
2. Library name. SharePoint's default library is displayed as "Documents" in the web UI but addressed as "Shared Documents" in Graph. Use "Shared Documents" in paths.
3. Case and spacing. Paths are case-insensitive in SharePoint, but trailing slashes and exact folder names matter.
4. Site versus subsite. The `site_url` must point at the correct site. If the file lives under `/sites/TeamSite/subsite-a/`, then `TeamSite` is the site and `/subsite-a/Shared Documents/...` is the path. Alternatively, set `site_url` to point at the subsite directly.
5. Deleted or moved file. Confirm the file exists by calling `sp.list_files` on the parent folder.

### `SharePointThrottledError: Rate limited after 3 retries`

The library already retried three times, honoring `Retry-After`, and Graph is still throttling.

1. Reduce concurrency. Serial is safest. Thread pools should be shrunk.
2. Back off longer. Catch the exception, sleep for `exc.retry_after` (or longer), and retry.
3. Do not run SharePoint I/O from Spark executors. See [usage-patterns.md](usage-patterns.md#rate-limiting).
4. App-level throttling applies when a single app registration generates heavy traffic against Graph. When multiple jobs share one registration, separate registrations for independent workloads allow them to throttle independently.

### `ValueError: Cannot detect environment from '...'`

The site URL's hostname doesn't end in `.sharepoint.com` or `.sharepoint.us`. Pass `graph_endpoint=` explicitly and build the credential with the right authority yourself. See [azure-gov.md](azure-gov.md#manual-override).

### `ValueError: Cannot parse SharePoint URL: '...'`

The library expects URLs in the form `https://{host}/sites/{site-name}/...`. Tenants using a different path structure (for example, personal sites under `/personal/`, or root-site files not under `/sites/`) are not currently supported. Open an issue describing the URL shape required.

### `ValueError: Named range '...' not found in workbook`

The template does not define a named range with that name.

1. Open the template in Excel and check Formulas → Name Manager for the exact name. Names are case-sensitive in openpyxl.
2. Workbook-scoped names live in `workbook.defined_names`. Sheet-scoped names live on the sheet. The library currently resolves only workbook-scoped names.

### `ValueError: Data (... rows x ... cols) exceeds range ...`

The call asked `fill_range` to write more data than the bounded range holds.

1. When the data is genuinely larger than expected, the template is out of date. Enlarge the range in the template, or investigate why the data grew.
2. To allow overflow when using `named_range`, pass `allow_expand=True`. This flag does not apply to `end_cell`. End-cell bounds are always strict.

### Excel file is corrupt or does not open

1. The write included `NaN` as a cell value. openpyxl writes NaN as the string "nan", which Excel may interpret inconsistently. Convert before writing: `df = df.where(df.notna(), None)`.
2. The write included non-serializable Python objects (custom classes, `numpy.datetime64` without conversion). Stick to native Python types, pandas Timestamps, or numeric types.
3. A formula string contains a character Excel rejects. Prefix formulas with `=` and confirm the body is valid Excel syntax.

### `read_excel` returns an empty DataFrame

1. Wrong sheet. When `sheet_name=None`, pandas normally returns all sheets as a dict. This library defaults to `0` (first sheet) when `None` is passed. Supply an explicit sheet name to avoid ambiguity.
2. Headers not on row 1. pandas assumes headers on the first row. When the template has a title row above the data, `skiprows` is required. This parameter is not exposed by `sp.read_excel`, so handle it manually:
   ```python
   content = sp.download(path)
   df = pd.read_excel(io.BytesIO(content), engine="openpyxl", sheet_name="Data", skiprows=3)
   ```
3. Merged cells can confuse pandas. Clean up the template or use `openpyxl` directly to read the sheet cell-by-cell.

### Driver runs out of memory on `toPandas()` or `read_excel`

Too much data is being moved to the driver. See [usage-patterns.md](usage-patterns.md#anti-pattern-collect-a-huge-spark-table-and-write-to-xlsx). The fix is almost always upstream: filter, aggregate, or sample in Spark before collecting.

### Everything looks right but 401 or 403 still occurs

Walk through this isolation checklist:

```python
from azure.identity import ClientSecretCredential

cred = ClientSecretCredential(
    tenant_id="<tenant-guid>",
    client_id="<client-id>",
    client_secret="<client-secret>",
    authority="https://login.microsoftonline.us",  # Gov. Omit for Commercial.
)

# This should succeed and print the start of a token string.
print(cred.get_token("https://graph.microsoft.us/.default").token[:20], "...")
```

If this fails, the problem is with the credential. If it succeeds but Graph calls still return 401, the credential is valid and a permission or consent is missing.

Call an endpoint that only needs a valid token, such as `/v1.0/$metadata`. If that succeeds but a specific site URL returns 403, the problem is site-level: a missing `Sites.Selected` grant, or the wrong site ID.
