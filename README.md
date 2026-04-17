# dbx-sharepoint

A small, dependency-light Python library for reading, writing, and templating Excel files in SharePoint from Databricks notebooks and jobs. Works on both **Azure Commercial** (`*.sharepoint.com`) and **Azure Government** (`*.sharepoint.us`), with the environment auto-detected from your site URL.

> **Positioning.** This library is a bridge/stop-gap for teams that need SharePoint and Excel integration today, especially on Azure Gov where the managed options are thinner. Once [Databricks Lakeflow Connect](docs/lakeflow-connect.md) supports SharePoint natively in your region, prefer it for ingestion workloads. This library remains useful for Excel-specific concerns such as templates, named ranges, and formatted report generation, and for lightweight file I/O in workflows that are not a good fit for a managed ingestion connector.

---

## Contents

- [Features](#features)
- [When to use this vs. Lakeflow Connect](#when-to-use-this-vs-lakeflow-connect)
- [Install](#install)
- [Quick start](#quick-start)
- [Public API at a glance](#public-api-at-a-glance)
- [Documentation](#documentation)

---

## Features

- Microsoft Graph-based SharePoint access. No SharePoint CSOM, no on-prem dependencies.
- Gov and Commercial endpoint auto-detection from the site URL. Explicit override supported.
- Excel read and write via pandas (`read_excel`, `write_excel`).
- Template workflow that preserves workbook formatting, formulas, and named ranges. Fill a range by `(sheet, start_cell)` or by `named_range`. Transpose with `orientation="columns"`. Optional bounds checking with `end_cell`.
- Raw file download and upload for any binary content (PDFs, CSVs, images) via bytes.
- Databricks secret-scope factory method: `SharePointClient.from_databricks_secrets(...)`.
- Shared-link helper for "anyone with the link" Excel reads that do not require auth.
- Typed exceptions mapped from HTTP status codes. 429s are retried automatically with `Retry-After`.

## When to use this vs. Lakeflow Connect

| Situation | Recommendation |
|---|---|
| You need to ingest large volumes of tabular SharePoint data into Unity Catalog on an ongoing basis | **Prefer Lakeflow Connect** when available in your region. Managed, incremental, with lineage. |
| You're on Azure Gov and Lakeflow Connect's SharePoint connector isn't yet GA there | Use this library as a bridge. |
| You need to produce **formatted Excel reports** (templates, named ranges, preserved styling) and publish them back to SharePoint | Use this library — report generation isn't Lakeflow Connect's job. |
| You need light, ad-hoc reads of a few SharePoint-hosted Excel files in a notebook | Use this library. |
| A notebook calls `.toPandas()` on a multi-million-row Spark DataFrame and writes it to a single `.xlsx` | Reshape the data first. See [Usage patterns](docs/usage-patterns.md). A collect of that size will exhaust driver memory, and Excel's hard ceiling is ~1.05M rows per sheet regardless. |

More on this in [docs/lakeflow-connect.md](docs/lakeflow-connect.md).

## Install

```bash
# From a built wheel
pip install dbx_sharepoint-0.1.0-py3-none-any.whl

# Or from a git ref
pip install "git+https://github.com/<org>/<repo>.git@main"
```

Inside a Databricks notebook:

```python
%pip install /dbfs/FileStore/wheels/dbx_sharepoint-0.1.0-py3-none-any.whl
dbutils.library.restartPython()
```

Python 3.9+ is required. Dependencies: `azure-identity`, `requests`, `pandas`, `openpyxl`.

## Quick start

### 1. Register an app in Azure AD and grant Graph permissions

You need an Azure AD app registration with **application permissions** for Microsoft Graph:

- `Sites.Selected` (recommended, least privilege) — then grant the app access to the specific SharePoint site, **or**
- `Sites.Read.All` / `Sites.ReadWrite.All` (broader) if your tenant policy allows it.

Full details: [docs/authentication.md](docs/authentication.md).

### 2. Store credentials in a Databricks secret scope

```bash
databricks secrets create-scope sharepoint
databricks secrets put-secret sharepoint tenant-id       # Azure AD tenant ID (a GUID)
databricks secrets put-secret sharepoint client-id       # Application (client) ID
databricks secrets put-secret sharepoint client-secret   # Client secret value
databricks secrets put-secret sharepoint site-url        # e.g. https://myorg.sharepoint.us/sites/TeamSite
```

### 3. Use it

```python
from dbx_sharepoint import SharePointClient

sp = SharePointClient.from_databricks_secrets(dbutils=dbutils, scope="sharepoint")

# Read an Excel file
df = sp.read_excel("/Shared Documents/data.xlsx")

# Write a DataFrame
sp.write_excel(df, "/Shared Documents/output.xlsx", sheet_name="Results")

# List a folder
files = sp.list_files("/Shared Documents/reports/")

# Download/upload any file
pdf_bytes = sp.download("/Shared Documents/report.pdf")
sp.upload(pdf_bytes, "/Shared Documents/archive/report.pdf")
```

Paths can be either relative to the site (`/Shared Documents/...`) or full SharePoint URLs. Both forms are accepted by every method that takes a path.

### 4. Fill an Excel template

```python
template = sp.open_template("/Shared Documents/templates/report_template.xlsx")
template.set_value("Cover", cell="B2", value="Q1 2026 Report")
template.fill_range("Data", start_cell="B3", data=summary_df)        # by cell anchor
template.fill_range(named_range="data_table", data=summary_df)        # by named range
sp.save(template, "/Shared Documents/reports/q1_2026.xlsx")
```

Details in [docs/excel-templates.md](docs/excel-templates.md).

## Public API at a glance

```python
from dbx_sharepoint import (
    SharePointClient,           # main client
    Template,                   # Excel template wrapper
    read_excel_from_shared_link,  # anon shared-link reader
    SharePointError,            # base exception
    SharePointAuthError,        # 401
    SharePointPermissionError,  # 403
    SharePointFileNotFoundError,  # 404
    SharePointThrottledError,   # 429 after retries
)
```

Full signatures in [docs/api-reference.md](docs/api-reference.md).

## Documentation

| Guide | What it covers |
|---|---|
| [Authentication & secrets](docs/authentication.md) | Azure AD app registration, Graph permission choices, Databricks secret-scope layout, bring-your-own credential |
| [Azure Gov](docs/azure-gov.md) | Gov vs Commercial endpoint table, auto-detection rules, overrides |
| [Usage patterns](docs/usage-patterns.md) | Read this before writing a job. Covers Excel limits, when not to use Excel, how to avoid collecting large tables to the driver, chunking, sampling, and write-back patterns |
| [Excel templates](docs/excel-templates.md) | Named ranges, transposition, bounds checking, preserving formulas and styles |
| [Lakeflow Connect](docs/lakeflow-connect.md) | Positioning, migration guidance, what to do once the managed connector is available |
| [Troubleshooting](docs/troubleshooting.md) | Mapping common errors to root causes and fixes |
| [API reference](docs/api-reference.md) | Full signatures for every public symbol |

## Development

```bash
# Install with dev extras
pip install -e ".[dev]"

# Run tests
pytest

# Build the wheel
pip install build && python -m build
```

Tests are fully offline and use `responses` to mock the Graph API.

## License

MIT. See [LICENSE](LICENSE).
