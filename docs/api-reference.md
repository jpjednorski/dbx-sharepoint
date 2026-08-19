# API reference

Complete reference for every public symbol exported by `dbx_sharepoint`.

```python
from dbx_sharepoint import (
    SharePointClient,
    Template,
    SensitivityLabel,
    apply_sensitivity_label,
    extract_sensitivity_label,
    read_excel_from_shared_link,
    SharePointError,
    SharePointAuthError,
    SharePointPermissionError,
    SharePointFileNotFoundError,
    SharePointThrottledError,
)
```

---

## `SharePointClient`

Main entry point for SharePoint operations via Microsoft Graph.

### Constructor

```python
SharePointClient(
    credential: TokenCredential,
    site_url: str,
    graph_endpoint: Optional[str] = None,
    default_sensitivity_label: Optional[SensitivityLabel] = None,
)
```

**Parameters**

- `credential` — any `azure.identity.TokenCredential` (`ClientSecretCredential`, `ManagedIdentityCredential`, `DefaultAzureCredential`, etc.).
- `site_url` — the SharePoint site URL in the form `https://{host}/sites/{site-name}`. The hostname determines whether the client uses Commercial or Gov Graph endpoints.
- `graph_endpoint` *(optional)* — override the auto-detected Graph endpoint. For unusual sovereign-cloud scenarios.
- `default_sensitivity_label` *(optional)* — a [`SensitivityLabel`](#sensitivitylabel) applied by `write_excel` when no per-call label is given. Template writes preserve the template's own label, so this default does not override them.

**Raises**

- `ValueError` if the site URL doesn't match `https://{host}/sites/{site-name}` or if the hostname isn't recognizable as `.sharepoint.com`/`.sharepoint.us` and no override is given.

### `from_databricks_secrets` (classmethod)

```python
SharePointClient.from_databricks_secrets(
    dbutils: Any,
    scope: str,
    prefix: str = "",
    site_url: Optional[str] = None,
    graph_endpoint: Optional[str] = None,
    default_sensitivity_label: Optional[SensitivityLabel] = None,
) -> SharePointClient
```

Factory that reads credentials from a Databricks secret scope and builds a client with the correct authority for the detected cloud. `default_sensitivity_label` is forwarded to the constructor.

**Expected secret keys** (prepended with `{prefix}-` if `prefix` is given):

- `tenant-id` — Azure AD tenant GUID
- `client-id` — app registration client ID
- `client-secret` — app registration client secret value
- `site-url` *(optional)* — SharePoint site URL; can be passed as `site_url=` instead

**Parameters**

- `dbutils` — the Databricks `dbutils` object from a notebook.
- `scope` — name of the secret scope.
- `prefix` — optional key prefix, e.g. `"prod"` looks for `prod-tenant-id` etc.
- `site_url` — if given, overrides the scope's `site-url` secret.
- `graph_endpoint` — optional explicit Graph endpoint override.

**Raises**

- `ValueError` if `site_url` isn't provided as a parameter and isn't in the scope.

### `list_files`

```python
list_files(url_or_path: str) -> pd.DataFrame
```

List items in a SharePoint folder.

**Parameters**

- `url_or_path` — full SharePoint URL or path relative to the site (e.g. `/Shared Documents/reports/`).

**Returns**

A pandas DataFrame with columns:

| Column | Type | Description |
|---|---|---|
| `name` | str | File or folder name |
| `path` | str | `webUrl` from Graph |
| `size_bytes` | int | 0 for folders |
| `modified_at` | str | ISO 8601 timestamp |
| `modified_by` | str | Display name of last modifier |
| `is_folder` | bool | `True` for folders |

### `download`

```python
download(url_or_path: str) -> bytes
```

Download a file and return its contents as bytes.

### `upload`

```python
upload(content: bytes, url_or_path: str) -> None
```

Upload bytes to the given path. Overwrites the file when it exists and creates it when it does not. Does not create missing parent folders. The parent folder must already exist.

### `read_excel`

```python
read_excel(
    url_or_path: str,
    sheet_name: Optional[Union[str, int]] = None,
) -> pd.DataFrame
```

Download an Excel file and return the specified sheet as a DataFrame. Defaults to the first sheet when `sheet_name` is `None`.

Uses pandas' openpyxl engine internally. For advanced reads (skip rows, specific dtypes, and similar), download the bytes with `download()` and call `pd.read_excel` directly.

### `write_excel`

```python
write_excel(
    df: pd.DataFrame,
    url_or_path: str,
    sheet_name: str = "Sheet1",
    sensitivity_label: Optional[SensitivityLabel] = None,
) -> None
```

Write a DataFrame as a new `.xlsx` file. This is a single-sheet convenience helper. `sensitivity_label` stamps a MIP label on the output; it falls back to the client's `default_sensitivity_label`. See [`SensitivityLabel`](#sensitivitylabel).

Writing to a `.xlsm` path raises `ValueError`. `write_excel` builds a fresh workbook with no macros, so the bytes are a regular `.xlsx`; saving them under a `.xlsm` extension produces a file Excel rejects as corrupt. To produce a valid `.xlsm`, use [`write_excel_from_template`](#write_excel_from_template) with a macro-enabled template. Reading a `.xlsm` with `read_excel` works normally — you get the cell data; macros are not represented in a DataFrame.

For multi-sheet workbooks, combine `pd.ExcelWriter` with `sp.upload`:

```python
buf = io.BytesIO()
with pd.ExcelWriter(buf, engine="openpyxl") as writer:
    df1.to_excel(writer, sheet_name="A", index=False)
    df2.to_excel(writer, sheet_name="B", index=False)
sp.upload(buf.getvalue(), "/Shared Documents/out.xlsx")
```

### `open_template`

```python
open_template(url_or_path: str) -> Template
```

Download an Excel file from SharePoint and return a `Template` object for editing. See [`Template`](#template).

### `write_excel_from_template`

```python
write_excel_from_template(
    df: pd.DataFrame,
    template_url_or_path: str,
    dest_url_or_path: str,
    sheet_name: Optional[str] = None,
    start_cell: str = "A1",
    orientation: str = "rows",
    include_header: bool = True,
    sensitivity_label: Optional[SensitivityLabel] = None,
) -> None
```

Download a template, fill it with `df` starting at `start_cell`, and upload the result to `dest_url_or_path`. `sheet_name` defaults to the template's active sheet. If `sheet_name` is missing from a single-sheet blank template, that blank sheet is renamed; otherwise a new sheet is created. `include_header=True` writes DataFrame column names before the values, matching `write_excel`. This is the one-call path for the common "fill one sheet and save" case.

Macro-enabled templates (`.xlsm`) are detected automatically and their VBA project is preserved, so the saved file opens cleanly in Excel. This is the only write path that produces a valid `.xlsm` — `write_excel` cannot, because it builds a fresh workbook with no macros. Keep the `.xlsm` extension on `dest_url_or_path`. See [Excel templates → Macro-enabled templates](excel-templates.md#macro-enabled-templates-xlsm).

A sensitivity label on the template is preserved on the output. `sensitivity_label` sets or overrides it explicitly. See [Excel templates → Sensitivity labels](excel-templates.md#sensitivity-labels-microsoft-information-protection).

### `save`

```python
save(template: Template, url_or_path: str) -> None
```

Upload a populated `Template` to the given path.

---

## `Template`

Wrapper around an `openpyxl` workbook with helpers for filling values into cells, ranges, and named ranges. Preserves styles, formulas, and workbook structure.

### Constructor

```python
Template(data: bytes, keep_vba: Optional[bool] = None)
```

This constructor is not typically called directly. Use `SharePointClient.open_template(...)` instead. Pass raw `.xlsx` or `.xlsm` bytes when loading a template from a source other than SharePoint. `keep_vba` defaults to auto-detecting macro-enabled files; pass `True`/`False` to force it.

### `fill_range`

```python
fill_range(
    sheet: Optional[str] = None,
    start_cell: Optional[str] = None,
    end_cell: Optional[str] = None,
    named_range: Optional[str] = None,
    data: Optional[pd.DataFrame] = None,
    orientation: str = "rows",
    allow_expand: bool = False,
    include_header: bool = False,
) -> None
```

Fill a rectangular range with DataFrame values.

Targeting: provide either `sheet + start_cell` or `named_range`. Not both.

**Parameters**

- `sheet` — sheet name (required with `start_cell`).
- `start_cell` — top-left anchor (e.g. `"B3"`).
- `end_cell` *(optional)* — bottom-right bound. If the data would overflow, raises `ValueError` before writing anything.
- `named_range` — named range defined in the workbook.
- `data` — DataFrame to write. **Required.**
- `orientation` — `"rows"` (default) writes each DataFrame row as a row in Excel; `"columns"` transposes so each DataFrame row becomes an Excel column.
- `allow_expand` — when using `named_range`, allows writing beyond the range's defined bounds. Ignored when using `end_cell`.
- `include_header` — when `True`, writes DataFrame column names before values. Defaults `False` because styled templates usually already contain headers.

**Raises**

- `ValueError` if neither or both targeting modes are provided, if the named range doesn't exist, if `data` is `None`, or if data exceeds the specified bounds.

Headers are only written when `include_header=True`. Styles, fonts, fills, number formats, and conditional formatting in the template are preserved.

### `set_value`

```python
set_value(sheet: str, cell: str, value: object) -> None
```

Set a single cell's value.

### `to_bytes`

```python
to_bytes() -> bytes
```

Serialize the current workbook state to `.xlsx` bytes. In typical use, pass the `Template` to `sp.save` instead of calling this directly. Any sensitivity label the template carries (or one set via `set_sensitivity_label`) is re-applied, since openpyxl drops the label part on save.

### Sensitivity label methods

```python
Template.sensitivity_label          # property → Optional[SensitivityLabel]
Template.set_sensitivity_label(label: SensitivityLabel) -> None
Template.clear_sensitivity_label() -> None
```

`sensitivity_label` is initialized from the label the template was loaded with, so a labeled template keeps its label through the round-trip. `set_sensitivity_label` overrides it; `clear_sensitivity_label` produces an unlabeled file. See [`SensitivityLabel`](#sensitivitylabel).

---

## `SensitivityLabel`

```python
SensitivityLabel(
    label_id: str,
    site_id: str,
    method: str = "Privileged",
    content_bits: int = 0,
    enabled: bool = True,
    removed: bool = False,
    name: Optional[str] = None,
)
```

A Microsoft Information Protection (MIP) sensitivity label in its classification-only (unencrypted) form — the metadata Office stores in `docMetadata/LabelInfo.xml`. Frozen dataclass; equality ignores `name`.

**Parameters**

- `label_id` — the label GUID, with or without braces (normalized to `{...}` on write).
- `site_id` — the tenant/site GUID the label belongs to.
- `method` — `"Privileged"` (set by explicit user action, what Excel writes when a person picks a label) or `"Standard"` (automatically applied). Either satisfies a mandatory-labeling policy.
- `content_bits` — visual-marking flags (header/footer/watermark). `0` for the classification-only case.
- `enabled` / `removed` — label state. Almost always `True` / `False`.
- `name` — optional human-readable name. Informational only; not written to the file and not part of equality.

Obtain the GUIDs from the Purview portal, from `Get-Label` in Security & Compliance PowerShell, or with `extract_sensitivity_label` on a hand-labeled file. See [Excel templates → Sensitivity labels](excel-templates.md#sensitivity-labels-microsoft-information-protection).

### `apply_sensitivity_label`

```python
apply_sensitivity_label(data: bytes, label: SensitivityLabel) -> bytes
```

Stamp a classification-only label onto workbook bytes, wiring `docMetadata/LabelInfo.xml` into the package and leaving every other part (including `xl/vbaProject.bin`) untouched. Idempotent. Does not apply encryption.

### `extract_sensitivity_label`

```python
extract_sensitivity_label(data: bytes) -> Optional[SensitivityLabel]
```

Read the label from a workbook, or return `None` if it is unlabeled (or not a valid OOXML package).

---

## `read_excel_from_shared_link`

```python
read_excel_from_shared_link(
    url: str,
    sheet_name: Optional[Union[str, int]] = None,
) -> pd.DataFrame
```

Read an Excel file from a SharePoint shared link (an "anyone with the link can view" URL) without authentication.

The helper appends `download=1` to the URL and issues an HTTP GET. Useful for quick tests, one-off reads, or reading files in a tenant where no app registration is available.

**Parameters**

- `url` — the shared link URL (can include existing query parameters).
- `sheet_name` — sheet to read; defaults to the first sheet.

**Raises**

- `requests.HTTPError` if the link is revoked, requires auth, or returns a non-200 status.

This helper is intentionally minimal and does not use the Graph API. Avoid using it in production workflows. Shared links rotate and can be revoked.

---

## Exceptions

All inherit from `SharePointError`:

```text
SharePointError
├── SharePointAuthError           # 401
├── SharePointPermissionError     # 403
├── SharePointFileNotFoundError   # 404
└── SharePointThrottledError      # 429 (after retries)
      .retry_after: Optional[int]
```

### `SharePointError`

Base class for every error raised from the library. Catch this when the specific subclass is not needed.

### `SharePointAuthError`

Authentication failure: bad credentials, wrong authority, expired secret, or missing admin consent.

### `SharePointPermissionError`

App is authenticated but not authorized for the resource.

### `SharePointFileNotFoundError`

Path doesn't exist. Usually a path-formatting or library-name issue (see [troubleshooting.md](troubleshooting.md#sharepointfilenotfounderror-not-found-)).

### `SharePointThrottledError`

HTTP 429 after internal retries are exhausted. The instance exposes `retry_after: Optional[int]`: seconds to wait before trying again, parsed from the `Retry-After` header on the final response.

```python
try:
    sp.read_excel(path)
except SharePointThrottledError as e:
    time.sleep(e.retry_after or 30)
    sp.read_excel(path)
```

---

## Internal modules (not part of the public API)

The following are imported internally, and their shapes may change. Do not depend on them from outside the package:

- `dbx_sharepoint._url`: URL parsing and environment detection.
- `dbx_sharepoint.auth.build_credential_from_databricks_secrets`: used by the factory method. Callable directly to obtain a `(credential, extras)` tuple without going through `SharePointClient`.
- `dbx_sharepoint.excel.dataframe_from_excel_bytes`, `dataframe_to_excel_bytes`: thin pandas wrappers, exported for testing convenience.
