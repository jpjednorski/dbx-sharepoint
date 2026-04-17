# Usage patterns

This guide covers how to use `dbx-sharepoint` effectively inside Databricks workloads, and, more importantly, how to avoid a few easy-to-miss pitfalls. Excel is a presentation format, Spark is a distributed compute engine, and combining the two without a few specific guardrails tends to fail in predictable ways.

Read this before writing a job that touches SharePoint.

---

## Contents

- [The mental model](#the-mental-model)
- [Excel's hard limits](#excels-hard-limits)
- [Anti-pattern: collect a huge Spark table and write to xlsx](#anti-pattern-collect-a-huge-spark-table-and-write-to-xlsx)
- [Patterns that work](#patterns-that-work)
  - [1. Aggregate in Spark, export the summary](#1-aggregate-in-spark-export-the-summary)
  - [2. Filter or sample before exporting](#2-filter-or-sample-before-exporting)
  - [3. Split across sheets (workbook batching)](#3-split-across-sheets-workbook-batching)
  - [4. One file per logical partition](#4-one-file-per-logical-partition)
  - [5. Use CSV when you only think you need Excel](#5-use-csv-when-you-only-think-you-need-excel)
  - [6. Templates for formatted reports](#6-templates-for-formatted-reports)
- [Reading: small files, big files, many files](#reading-small-files-big-files-many-files)
- [Writing: idempotency and partial failures](#writing-idempotency-and-partial-failures)
- [Rate limiting](#rate-limiting)
- [Checklist](#checklist)

---

## The mental model

Every operation in this library runs on the driver, through a single `requests` session, against the Microsoft Graph REST API. There is no Spark distribution of SharePoint I/O, and there cannot be: Graph is a per-tenant API with per-app rate limits.

That has two consequences:

1. Data has to fit in driver memory to be written to or read from an Excel file through this library. Not "fit on the cluster." Fit on the driver, with headroom for pandas and openpyxl's in-memory workbook representation.
2. Throughput is bounded by Graph, not by cluster size. Larger clusters do not accelerate SharePoint I/O.

Use Spark to prepare the payload (filter, aggregate, deduplicate, shape). Use this library to move the final, small payload to or from SharePoint.

## Excel's hard limits

These are Excel file format limits, not library limits:

| Limit | Value |
|---|---|
| Rows per sheet | 1,048,576 |
| Columns per sheet | 16,384 |
| Characters per cell | 32,767 |
| Sheets per workbook | limited by memory |

If you try to write more than ~1.05M rows to one sheet, openpyxl will fail, and even if it didn't, Excel wouldn't open the file.

Practical rules of thumb:

- Under ~100k rows: Excel handles this comfortably.
- 100k to 500k rows: works, but workbook opening becomes noticeably slow. Consider pivoting or filtering first.
- 500k to 1M rows: near the upper limit of what Excel is designed for. Evaluate whether Excel is the right format.
- Over 1M rows: Excel is the wrong format. Use CSV, Parquet, or Delta, or aggregate first.

## Anti-pattern: collect a huge Spark table and write to xlsx

This is the most common pitfall, and the primary reason this guide exists. The pattern looks like this:

```python
# Don't do this.
df_spark = spark.table("main.sales.transactions")   # 40M rows
pdf = df_spark.toPandas()                            # collects to driver -> OOM
sp.write_excel(pdf, "/Shared Documents/transactions.xlsx")
```

What goes wrong, in order:

1. `.toPandas()` pulls every row to the driver. At 40M rows and a handful of columns, driver memory is exhausted long before any SharePoint call runs.
2. If the data did fit, `DataFrame.to_excel` with openpyxl holds the whole workbook in memory as cells. That representation is roughly an order of magnitude larger than the pandas DataFrame.
3. If that step also succeeded, the write hits Excel's 1,048,576-row ceiling at row 1,048,577.
4. Batching across sheets does not help. The upload is still one in-memory workbook serialization, and memory is exhausted again.
5. Even if the upload succeeded, the resulting file would be unusable when opened in Excel.

The fix is almost always to reshape the data before exporting.

## Patterns that work

### 1. Aggregate in Spark, export the summary

When the consumer wants to understand the data in Excel, they rarely need every row. Aggregate in Spark first, then export the small result:

```python
summary = (
    spark.table("main.sales.transactions")
      .where("transaction_date >= '2026-01-01'")
      .groupBy("region", "product_category")
      .agg(F.sum("amount").alias("total"), F.count("*").alias("txn_count"))
      .orderBy("region", "product_category")
)

pdf = summary.toPandas()          # e.g. a few hundred rows
sp.write_excel(pdf, "/Shared Documents/reports/sales_summary.xlsx", sheet_name="Summary")
```

Aggregation typically reduces millions of rows to hundreds or thousands, which Excel handles comfortably.

### 2. Filter or sample before exporting

When you really do need row-level data, filter to the rows that matter:

```python
pdf = (
    spark.table("main.sales.transactions")
      .where("region = 'NW' AND transaction_date >= '2026-04-01'")
      .limit(50_000)                         # hard cap as a safety net
      .toPandas()
)
sp.write_excel(pdf, "/Shared Documents/reports/nw_april.xlsx")
```

For exploratory sharing, a sample is often enough:

```python
sample = spark.table("main.sales.transactions").sample(fraction=0.001, seed=42).limit(25_000)
pdf = sample.toPandas()
sp.write_excel(pdf, "/Shared Documents/reports/sample.xlsx")
```

### 3. Split across sheets (workbook batching)

If the requirement is genuinely "all the rows in one workbook" and the total is manageable (a few hundred thousand, say), split by a natural key so each sheet stays under Excel's limit:

```python
import pandas as pd

pdf = spark_df.toPandas()   # already filtered down to a size that fits on the driver

buf = io.BytesIO()
with pd.ExcelWriter(buf, engine="openpyxl") as writer:
    for region, chunk in pdf.groupby("region"):
        chunk.to_excel(writer, sheet_name=region, index=False)

sp.upload(buf.getvalue(), "/Shared Documents/reports/by_region.xlsx")
```

This uses `upload` + pandas directly because `write_excel` is a one-sheet convenience wrapper.

### 4. One file per logical partition

Often a per-region, per-month, or per-team file is a better deliverable than one large workbook. Drive the loop from Spark metadata rather than from a collect of the full dataset:

```python
regions = [row["region"] for row in spark_df.select("region").distinct().collect()]

for region in regions:
    part = spark_df.where(F.col("region") == region).toPandas()
    sp.write_excel(part, f"/Shared Documents/reports/by_region/{region}.xlsx")
```

Two things to note:

- Each iteration holds one partition in driver memory, not the whole table.
- The calls to SharePoint are serial. With 200 regions and a one-second upload each, the loop takes minutes, which is usually acceptable. If it is not, parallelize with a bounded thread pool, but expect to hit Graph rate limits (see [Rate limiting](#rate-limiting)).

### 5. Use CSV when you only think you need Excel

If the consumer just needs to open the file in a spreadsheet and the data has no formatting requirement, CSV is often better:

- No 1M-row limit. Modern Excel opens multi-million-row CSVs, slowly, but without failing.
- Streams natively. No workbook has to be materialized in memory.
- Can be written from Spark directly (`df.write.csv(...)`) when the output lives in object storage.

If it needs to live on SharePoint, write CSV bytes and upload:

```python
csv_bytes = pdf.to_csv(index=False).encode("utf-8")
sp.upload(csv_bytes, "/Shared Documents/reports/transactions.csv")
```

### 6. Templates for formatted reports

When the consumer wants a formatted report (branded header, pre-built pivot charts, specific cells for totals), build an Excel template once, check it into version control or store it in SharePoint, and have the job fill in the data each run:

```python
template = sp.open_template("/Shared Documents/templates/monthly_report.xlsx")

template.set_value("Cover", cell="B2", value=f"Report for {period}")
template.set_value("Cover", cell="B3", value=pd.Timestamp.utcnow())

template.fill_range(named_range="metrics_table", data=metrics_df)
template.fill_range("Details", start_cell="A3", end_cell="F1000", data=details_df)

sp.save(template, f"/Shared Documents/reports/{period}.xlsx")
```

The template preserves styling, formulas that reference the filled cells, and named ranges. Only cell values are written, not styles, so the output consistently matches the template.

Details in [excel-templates.md](excel-templates.md).

## Reading: small files, big files, many files

For reading, the same principle applies: the whole file comes to the driver via `download` / `read_excel`.

- Small or medium files (< ~50 MB): call `sp.read_excel(path)` directly.
- Large Excel files: read once, then work on the resulting pandas DataFrame. To move it to Spark scale, use `spark_df = spark.createDataFrame(pdf)`. This only helps if the file was small enough to fit on the driver in the first place.
- Many files (for example, every xlsx in a folder): use `list_files` to enumerate, then loop. Be aware of rate limits, and consider adding a short sleep between calls when the folder has hundreds of files.

```python
files = sp.list_files("/Shared Documents/inbox/")
excel_files = files[
    (~files["is_folder"]) & (files["name"].str.endswith(".xlsx"))
]

frames = []
for path in excel_files["path"]:
    frames.append(sp.read_excel(path))

combined = pd.concat(frames, ignore_index=True)
```

## Writing: idempotency and partial failures

- `upload`, `write_excel`, and `save` all issue a PUT to the drive-item path. PUTs are idempotent: re-running the same write overwrites the same file.
- There is no transaction across files. If a job writes ten files and fails on the fifth, the first four are on SharePoint. Structure jobs so that re-running from scratch is safe, which idempotent PUTs make straightforward.
- There is no "create-if-not-exists" primitive. Writing to a target path creates it, or overwrites it if it already exists.
- **Large files are handled automatically.** For payloads larger than 4 MiB, the library switches from a single PUT to a Microsoft Graph [upload session](https://learn.microsoft.com/graph/api/driveitem-createuploadsession), uploading in 10 MiB chunks. Graph's simple-upload endpoint is unreliable above 4 MiB, which is easy to hit with a moderately sized Excel file. There is no action required on your side — call `upload`/`write_excel`/`save` as normal. Graph's hard ceiling is 250 GiB per file, well beyond anything Excel can open.

## Rate limiting

Microsoft Graph enforces per-app throttling. The library already:

- Retries on HTTP 429 up to 3 times, honoring the `Retry-After` header.
- Raises `SharePointThrottledError` if retries are exhausted.

You can still hit the ceiling if you fire many concurrent requests from a large cluster. Rules of thumb:

- Serial loops are usually safe. A tight loop of several hundred sequential downloads rarely triggers 429s.
- Parallel fan-outs need care. If you use a thread pool for uploads, keep it small (4 to 8 workers) and add jitter.
- Do not run SharePoint I/O from executors. Wrapping `sp.upload` in a pandas UDF may look convenient, but every executor then issues concurrent requests to Graph under the same app credential. That is the access pattern Graph throttling is designed to catch.

## Checklist

Before you ship a job that uses this library, walk through:

- [ ] Does the data fit in driver memory with at least 3–5× headroom for pandas/openpyxl?
- [ ] If the data is large, have I filtered, aggregated, or sampled in Spark before collecting?
- [ ] Is Excel the right format, or would CSV / Parquet / Delta serve the consumer better?
- [ ] If the output is a polished report, am I using a template rather than building formatting in code?
- [ ] Are my SharePoint operations on the driver (not inside a UDF or executor task)?
- [ ] Is the job idempotent? Does re-running it from scratch produce the same result?
- [ ] Do I handle `SharePointFileNotFoundError` / `SharePointPermissionError` / `SharePointThrottledError` where appropriate?
- [ ] For ongoing ingestion workloads, have I checked whether [Lakeflow Connect](lakeflow-connect.md) is now available in my region?
