# Positioning vs. Lakeflow Connect

## Summary

- Lakeflow Connect is the Databricks-managed path for ingesting data from external sources (including SharePoint) into Unity Catalog. When it is available and fits the workload, prefer it.
- `dbx-sharepoint` is a small client library intended as a bridge or stop-gap for teams that need SharePoint and Excel integration today, particularly on Azure Gov, where managed-connector availability tends to lag Commercial.
- Even after Lakeflow Connect's SharePoint connector reaches GA everywhere, this library remains useful for Excel-specific concerns: template-driven report generation, named-range writes, and small-file round-trips that are not ingestion pipelines.

---

## Why Lakeflow Connect is the preferred long-term path for ingestion

Lakeflow Connect is purpose-built for moving data from SaaS sources and file systems into Delta tables governed by Unity Catalog. For ongoing ingestion workloads (such as pulling a folder of Excel files into a bronze table every hour), it provides capabilities that this library does not attempt to reproduce:

- Managed, serverless execution. No cluster to size or patch. No cold-start penalty for small batches.
- Incremental ingestion. Only changed data is read, rather than reprocessing the full folder every run.
- Schema evolution. Added or renamed columns across file versions are handled automatically.
- Unity Catalog lineage. End-to-end lineage from source file to downstream tables and dashboards.
- Built-in retries, monitoring, and observability. Failures surface in the standard Databricks UI.
- No credential handling in user code. Authentication and secret rotation are managed.
- Governance and access control through standard UC grants.

For "land SharePoint data in a lakehouse table on a schedule," Lakeflow Connect is almost certainly the right answer once it is available for the target cloud and region.

## Why this library exists anyway

The list above describes an ingestion pipeline. Not every SharePoint workload is one. This library handles the cases that are not, and the cases where Lakeflow Connect is not yet an option.

### 1. Availability gaps

Managed connectors roll out to Azure Commercial first and to Azure Government on a delay. For teams running on Gov today who need SharePoint-to-Databricks integration now, the options are to wait or to use a small library until the managed option arrives. This is the primary reason this library exists.

### 2. Producing formatted reports, not ingesting data

Lakeflow Connect reads from sources into the lakehouse. It is not a general file-I/O layer, and it is not built for writing formatted Excel back out. Workflows such as these are output and round-trip workloads rather than ingestion:

- Computing metrics in Spark, filling a branded Excel template, and publishing it to a SharePoint folder for a stakeholder audience.
- Pulling a single config file from SharePoint, applying it in a job, and writing an updated version back.

This library handles such workloads directly.

### 3. Excel-specific features

The `Template` class in `dbx-sharepoint` (named ranges, cell-level writes, orientation transposition, bounds checking) addresses a different problem than ingestion. Producing formatted reports is a common analytics workflow, and an ingestion connector has no equivalent.

### 4. Small, ad-hoc, non-pipeline use

For a one-off notebook that pulls a single Excel file and plots it, standing up a managed ingestion pipeline is overkill. `sp.read_excel(path)` is two lines.

## Migration guidance

When Lakeflow Connect's SharePoint connector becomes available in the target environment, revisit any jobs built on `dbx-sharepoint` and evaluate:

| Question | If yes | If no |
|---|---|---|
| Is this job a recurring ingestion of SharePoint files into a Databricks table? | Migrate to Lakeflow Connect. Delete the `dbx-sharepoint` code. | Keep it on this library. |
| Does the job produce formatted Excel output (templates, branded reports)? | Keep this library. Lakeflow Connect does not target this. | — |
| Is the job round-tripping configuration or control files? | Keep this library, or evaluate whether UC Volumes is a better home for those files. | — |
| Is the job small, ad-hoc, or non-scheduled? | Keep this library when it is simpler to maintain. Managed-connector overhead is rarely justified for one-shot work. | — |

The library's exception types and return shapes are intentionally conventional (`pd.DataFrame`, `bytes`, standard Python exceptions). Swapping an ingestion call for a Lakeflow Connect-sourced table is typically a local change.

## Feature comparison (as of writing)

> Availability moves; treat this as indicative, not authoritative. Check Databricks docs for current regional availability of Lakeflow Connect's SharePoint connector.

| Concern | Lakeflow Connect | `dbx-sharepoint` |
|---|---|---|
| Primary use | Managed ingestion into UC | Ad-hoc file I/O and Excel report generation |
| Serverless | Yes | Runs in your cluster / job |
| Incremental | Yes | No — full download/upload per call |
| Lineage in UC | Yes | No |
| Gov region availability | Varies; may lag Commercial | Gov + Commercial, auto-detected |
| Excel template/named-range writes | Not its purpose | Yes |
| Producing formatted reports back to SharePoint | Not its purpose | Yes |
| Shared-link Excel reads (no auth) | No | Yes |
| Credential management | Managed | Caller provides (Databricks secrets or BYO credential) |

## A note on direction

This library is intentionally small. It is not trying to become a managed connector, and it does not carry a long-term roadmap. Expect it to stabilize at its current scope, receive bug fixes and Gov-specific adjustments, and eventually step aside for Lakeflow Connect on the ingestion path. It will continue to handle the report-writing and small-file cases that a managed connector does not target.
