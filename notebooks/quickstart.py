# Databricks notebook source
# MAGIC %md
# MAGIC # dbx-sharepoint Quickstart
# MAGIC
# MAGIC Simplified SharePoint file and Excel interfaces for Azure Gov Databricks.

# COMMAND ----------

# MAGIC %pip install /path/to/dbx_sharepoint-0.3.0-py3-none-any.whl
# MAGIC dbutils.library.restartPython()

# COMMAND ----------

# MAGIC %md
# MAGIC ## 1. Connect to SharePoint
# MAGIC
# MAGIC Expects these secrets in your scope:
# MAGIC - `tenant-id`
# MAGIC - `client-id`
# MAGIC - `client-secret`
# MAGIC - `site-url` (optional — can pass as parameter instead)

# COMMAND ----------

from dbx_sharepoint import SharePointClient

sp = SharePointClient.from_databricks_secrets(dbutils=dbutils, scope="sharepoint")

# COMMAND ----------

# MAGIC %md
# MAGIC ## 2. List Files

# COMMAND ----------

files = sp.list_files("/Shared Documents/")
display(files)

# COMMAND ----------

# MAGIC %md
# MAGIC ## 3. Read an Excel File

# COMMAND ----------

df = sp.read_excel("/Shared Documents/data.xlsx")
display(df)

# COMMAND ----------

# MAGIC %md
# MAGIC ## 4. Write a DataFrame to Excel

# COMMAND ----------

import pandas as pd

output_df = pd.DataFrame({"metric": ["revenue", "cost"], "value": [100000, 45000]})
sp.write_excel(output_df, "/Shared Documents/output.xlsx", sheet_name="Metrics")
print("Uploaded successfully!")

# COMMAND ----------

# MAGIC %md
# MAGIC ## 5. Template Workflow

# COMMAND ----------

template = sp.open_template("/Shared Documents/templates/report_template.xlsx")

# Fill data into the template
template.fill_range("Summary", start_cell="B3", data=output_df)
template.set_value("Summary", cell="A1", value="Q1 2026 Report")

# Save to a new location
sp.save(template, "/Shared Documents/reports/q1_2026.xlsx")
print("Template saved!")

# COMMAND ----------

# MAGIC %md
# MAGIC ## 6. Sensitivity Labels (Microsoft Information Protection)
# MAGIC
# MAGIC If your tenant enforces **mandatory sensitivity labeling**, an unlabeled
# MAGIC file opens fine on double-click but can't be read from another workbook
# MAGIC (external references, Power Query, linked workbooks). openpyxl drops the
# MAGIC label when it saves, so files written here need the label re-applied.
# MAGIC
# MAGIC A label is identified by **two GUIDs**, not by its display name. You do
# MAGIC **not** pass the human name like `"INTERNAL"` — you pass the label's GUID
# MAGIC and your tenant/site GUID, wrapped in a `SensitivityLabel`. (The `name`
# MAGIC field is just a readable tag for your own code; it isn't written to the
# MAGIC file.) This applies classification-only labels — it does not encrypt.

# COMMAND ----------

# MAGIC %md
# MAGIC ### Finding your label's GUIDs
# MAGIC
# MAGIC The easiest way: label one file by hand in Excel, then read the IDs back
# MAGIC out of it. (Alternatively, `Get-Label` in Security & Compliance
# MAGIC PowerShell lists every label's `DisplayName` and `Guid`.)

# COMMAND ----------

from dbx_sharepoint import extract_sensitivity_label

# Download a file you labeled by hand in Excel and read its label identifiers.
sample = sp.download("/Shared Documents/hand-labeled-sample.xlsx")
discovered = extract_sensitivity_label(sample)
print(discovered)  # -> SensitivityLabel(label_id='{...}', site_id='{...}', ...)

# COMMAND ----------

# MAGIC %md
# MAGIC ### Defining a label
# MAGIC
# MAGIC The GUIDs below are placeholders — replace them with the values you
# MAGIC discovered above (or from your Purview admin).

# COMMAND ----------

from dbx_sharepoint import SensitivityLabel

label = SensitivityLabel(
    label_id="{00000000-1111-2222-3333-444444444444}",  # your label's GUID
    site_id="{aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee}",    # your tenant/site GUID
    name="INTERNAL",                                      # informational only
)

# COMMAND ----------

# MAGIC %md
# MAGIC ### Applying it
# MAGIC
# MAGIC A DataFrame carries no label, so pass one explicitly when writing:

# COMMAND ----------

sp.write_excel(
    output_df,
    "/Shared Documents/labeled-output.xlsx",
    sheet_name="Metrics",
    sensitivity_label=label,
)
print("Wrote a labeled file.")

# COMMAND ----------

# MAGIC %md
# MAGIC **Templates preserve their own label automatically** — label the template
# MAGIC once in Excel and every generated file inherits it, with no code change:

# COMMAND ----------

# The output keeps whatever label the template was saved with.
sp.write_excel_from_template(
    output_df,
    "/Shared Documents/templates/MasterTemplate.xlsm",
    "/Shared Documents/reports/report.xlsm",
    sheet_name="ReportData",
)

# To override the template's label (or label a template that has none), pass one:
sp.write_excel_from_template(
    output_df,
    "/Shared Documents/templates/MasterTemplate.xlsm",
    "/Shared Documents/reports/report.xlsm",
    sheet_name="ReportData",
    sensitivity_label=label,
)

# COMMAND ----------

# MAGIC %md
# MAGIC To label every `write_excel` output without repeating yourself, set a
# MAGIC default on the client:

# COMMAND ----------

sp_labeled = SharePointClient.from_databricks_secrets(
    dbutils=dbutils, scope="sharepoint", default_sensitivity_label=label
)
sp_labeled.write_excel(output_df, "/Shared Documents/auto-labeled.xlsx")

# COMMAND ----------

# MAGIC %md
# MAGIC ## 7. Quick Test with Shared Link (No Auth)

# COMMAND ----------

from dbx_sharepoint import read_excel_from_shared_link

# Paste any "anyone with the link" URL here
# df = read_excel_from_shared_link("https://myorg.sharepoint.us/:x:/s/Team/EaBcDeFg...")
# display(df)
