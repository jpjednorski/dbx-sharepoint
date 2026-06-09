# Excel templates

The `Template` class opens an `.xlsx` file from SharePoint, writes values into specific cells, named ranges, or cell ranges (preserving the rest of the workbook), and saves the result back. It is the recommended path for producing formatted reports without reimplementing the formatting in Python.

Internally, `Template` wraps an `openpyxl` workbook, so anything openpyxl supports works. The class itself exposes a small surface area focused on the common "fill the blanks" workflow.

## When to use templates

- You have a report with branded styling, colored headers, merged cells, charts, or logos, and you want to keep all of that intact while updating the numbers each run.
- You have formulas that reference the cells being filled, and those formulas need to recompute when the workbook is opened.
- You want consumers to open the file in Excel and see something that looks like a real report, not a dumped DataFrame.

If none of the above apply, use `write_excel` instead. It is simpler.

## Anatomy of a template workflow

```python
# 1. Download the template from SharePoint
template = sp.open_template("/Shared Documents/templates/monthly_report.xlsx")

# 2. Fill specific cells (titles, dates, totals)
template.set_value("Cover", cell="B2", value="April 2026 Report")
template.set_value("Cover", cell="B3", value=pd.Timestamp.utcnow())

# 3. Fill ranges with DataFrames
template.fill_range("Metrics", start_cell="A3", data=metrics_df)
template.fill_range(named_range="details_table", data=details_df)

# 4. Save back to SharePoint
sp.save(template, "/Shared Documents/reports/2026-04.xlsx")
```

## `fill_range` in detail

```python
template.fill_range(
    sheet=None,             # sheet name; required if using start_cell
    start_cell=None,        # e.g. "B3"
    end_cell=None,          # optional bound; raises if data exceeds
    named_range=None,       # named range, alternative to sheet+start_cell
    data=None,              # DataFrame (required)
    orientation="rows",     # "rows" | "columns"
    allow_expand=False,     # for named ranges, allow writing beyond bounds
    include_header=False,   # write DataFrame column names before values
)
```

Exactly one of `(sheet + start_cell)` **or** `named_range` must be provided. `data` is always required.

### By cell anchor

Write a DataFrame starting at a specific cell. Expands right and down from the anchor:

```python
template.fill_range("Data", start_cell="B3", data=df)
```

If `df` has 5 rows and 3 columns, values land in B3:D7.

### With an end-cell bound

To guard against a case where the data grows larger than the template expects:

```python
template.fill_range("Data", start_cell="B3", end_cell="F23", data=df)
```

If `df` would overflow B3:F23, the call raises `ValueError` before writing anything. Nothing is half-written.

### By named range

If your template defines named ranges (Formulas → Name Manager in Excel), you can fill by name:

```python
template.fill_range(named_range="metrics_table", data=df)
```

By default, writing beyond the named range's bounds raises `ValueError`. To allow overflow (e.g., the table row count grows naturally), pass `allow_expand=True`:

```python
template.fill_range(named_range="metrics_table", data=df, allow_expand=True)
```

Why bounds matter: named ranges are often referenced by formulas elsewhere in the workbook. A job that silently writes past them leaves downstream calculations inconsistent.

### Transposition: `orientation="columns"`

By default, each DataFrame row becomes an Excel row. To have each DataFrame row become an Excel column instead, pass `orientation="columns"`. This is common when the template lays out metrics vertically with periods across the top:

```python
# df:
#   period    revenue   cost
#   Q1 2026   100000    45000
#   Q2 2026   125000    50000

template.fill_range("Metrics", start_cell="B3", data=df, orientation="columns")

# Excel receives:
#   B3: Q1 2026    C3: Q2 2026
#   B4: 100000     C4: 125000
#   B5: 45000      C5: 50000
```

### What `fill_range` does not do by default

- Does not write column headers unless `include_header=True`. Most styled templates already have header cells. The one-call `write_excel_from_template` helper does include headers by default because it mirrors `write_excel`.
- Does not coerce types. Each DataFrame cell value is written as-is. pandas `Timestamp` values become Excel dates, strings remain strings, and `NaN` values land as blank cells.
- Does not reset styles. Only cell values change. The template's fills, fonts, borders, and number formats are preserved.

## `set_value`

For single-cell writes such as title bars, report dates, or signature lines:

```python
template.set_value("Cover", cell="B2", value="April 2026 Report")
template.set_value("Summary", cell="E1", value=42)
```

Accepts any value openpyxl can serialize (strings, numbers, dates, bools, formulas as strings prefixed with `=`).

## `to_bytes` and saving

`template.to_bytes()` returns the serialized workbook. In most cases, pass the `Template` directly to `sp.save`:

```python
sp.save(template, "/Shared Documents/reports/out.xlsx")
```

`sp.save` is equivalent to `sp.upload(template.to_bytes(), path)`. Call `to_bytes` directly when the file needs to go somewhere other than SharePoint.

## Macro-enabled templates (`.xlsm`)

Templates that contain VBA macros, buttons, or `Workbook_Open` handlers are stored as `.xlsm`. These work the same way as `.xlsx` templates — the only difference is that the VBA project must be preserved on save, or Excel will report the file as corrupt when opening it.

This is handled automatically. `open_template` (and the `Template` constructor) inspect the file for an embedded VBA project and, when present, load the workbook with `keep_vba=True`:

```python
# .xlsm is detected automatically — macros are preserved
t = sp.open_template("/Shared Documents/templates/inspection.xlsm")
t.fill_range("Data", start_cell="A2", data=results_df)
t.set_value("Cover", cell="B2", value="June 2026")
sp.save(t, "/Shared Documents/reports/2026-06.xlsm")
```

Keep the `.xlsm` extension on the destination path so Excel treats the saved file as macro-enabled.

For the common "fill one sheet from a DataFrame and upload" case, there is a one-call helper:

```python
sp.write_excel_from_template(
    df=results_df,
    template_url_or_path="/Shared Documents/templates/inspection.xlsm",
    dest_url_or_path="/Shared Documents/reports/2026-06.xlsm",
    sheet_name="Data",     # defaults to the template's active sheet
    start_cell="A2",       # defaults to "A1"
)
```

If `sheet_name` does not exist and the template is a single blank sheet, the helper renames that blank sheet. This supports a reusable blank macro-enabled template such as `MasterTemplate.xlsm` while still producing output sheets like `ReportData`, `Orders`, or `Summary`. If the workbook already has content, a missing `sheet_name` creates a new sheet instead.

If you have the template bytes in hand (not from SharePoint), use the underlying function directly:

```python
from dbx_sharepoint import dataframe_to_excel_bytes_from_template

out_bytes = dataframe_to_excel_bytes_from_template(template_bytes, df, start_cell="A2")
```

Notes:
- Macro **code** is preserved verbatim; it is not executed. Macros run when a user opens the file in Excel, exactly as designed.
- To force the behavior either way, pass `keep_vba=True`/`False` to the `Template` constructor.

## Common patterns

### Parameterized report

One template, one output per period, driven by a list:

```python
periods = ["2026-01", "2026-02", "2026-03"]

for period in periods:
    t = sp.open_template("/Shared Documents/templates/monthly_report.xlsx")
    metrics = compute_metrics(period).toPandas()
    t.set_value("Cover", cell="B2", value=period)
    t.fill_range(named_range="metrics_table", data=metrics)
    sp.save(t, f"/Shared Documents/reports/{period}.xlsx")
```

Each iteration re-downloads the template. This is intentional: it guarantees a clean starting point and avoids mutating a single `Template` instance across runs.

### Multi-sheet fill

One workbook, multiple named ranges:

```python
t = sp.open_template("/Shared Documents/templates/quarterly.xlsx")

t.fill_range(named_range="revenue_table", data=revenue_df)
t.fill_range(named_range="cost_table", data=cost_df)
t.fill_range(named_range="headcount_table", data=headcount_df)
t.set_value("Cover", cell="B2", value="Q1 2026")

sp.save(t, "/Shared Documents/reports/Q1_2026.xlsx")
```

### Guarding against data growth

When a range has a fixed-size block backing a chart:

```python
# The chart references A3:F23 — if we write more rows, the chart lies.
t.fill_range("Data", start_cell="A3", end_cell="F23", data=df)
```

If `df` grows beyond 21 rows, the call raises `ValueError`. This is usually the desired behavior: fix the template, not the output.

## Limits and caveats

- openpyxl is pure-Python and entirely in-memory. Templates with thousands of heavily-styled cells take longer to load. Very large templates (~100 MB) may be slow.
- Macro-enabled workbooks (`.xlsm`) are supported. `Template` and `open_template` auto-detect an embedded VBA project and load it with `keep_vba=True`, so macros survive the round-trip and the saved file opens cleanly in Excel. See [Macro-enabled templates](#macro-enabled-templates-xlsm) below. Loading a `.xlsm` without preserving its VBA project (the openpyxl default) produces a file Excel reports as corrupt.
- Formulas are preserved but not recalculated. openpyxl writes the formula string; Excel recalculates when the file is opened. A job that reads back a saved template and expects formula results will not find them unless the file is opened in Excel first or the values are computed in Python.
- Charts that reference filled cells update automatically when the file is opened. Charts that reference cells outside the filled range do not update.
- Conditional formatting that references filled cells is applied as expected when the file is opened in Excel.
