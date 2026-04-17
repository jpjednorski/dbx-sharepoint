# Databricks notebook source
# MAGIC %md
# MAGIC # dbx-sharepoint Quickstart
# MAGIC
# MAGIC Simplified SharePoint file and Excel interfaces for Azure Gov Databricks.

# COMMAND ----------

# MAGIC %pip install /path/to/dbx_sharepoint-0.1.0-py3-none-any.whl
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
# MAGIC ## 6. Quick Test with Shared Link (No Auth)

# COMMAND ----------

from dbx_sharepoint import read_excel_from_shared_link

# Paste any "anyone with the link" URL here
# df = read_excel_from_shared_link("https://myorg.sharepoint.us/:x:/s/Team/EaBcDeFg...")
# display(df)
