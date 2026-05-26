# =============================================================================
# OSI MISSION – AnaCredit Exposure Analysis Tool
# =============================================================================
# Description:
# Python-based AnaCredit OSI tool for SQL data cleaning,
# aggregation and exposure visualisation.
#
# Purpose:
# - Connect to OSI Mission SQL database
# - Pull AnaCredit-style loan tape data
# - Clean and enrich datasets
# - Calculate debtor/group exposures
# - Create Top-20 exposure visualisations
# - Export results to Excel
#
# =============================================================================

# =============================================================================
# IMPORT LIBRARIES
# =============================================================================

import pandas as pd
import matplotlib.pyplot as plt
from sqlalchemy import create_engine

# =============================================================================
# DATABASE CONNECTION
# =============================================================================
# Example SQL Server connection
# Replace placeholders with actual credentials

SERVER = "ECB-OSI-SQL01"
DATABASE = "OSI_MISSION"
USERNAME = "your_username"
PASSWORD = "your_password"

connection_string = (
    f"mssql+pyodbc://{USERNAME}:{PASSWORD}@{SERVER}/{DATABASE}"
    "?driver=ODBC+Driver+17+for+SQL+Server"
)

engine = create_engine(connection_string)

# =============================================================================
# SQL QUERY
# =============================================================================
# Pull AnaCredit-style OSI mission data

query = """

SELECT
    l.loan_id,
    l.debtor_id,
    c.debtor_name,
    c.group_id,
    c.group_name,
    l.exposure_amount,
    l.collateral_value,
    l.default_status,
    l.forbearance_status,
    l.nace_code,
    l.country_code,
    l.product_type_code,
    l.stage,
    l.origination_date

FROM LOAN_DATA l

LEFT JOIN COUNTERPARTY_DATA c
    ON l.debtor_id = c.debtor_id

WHERE l.reference_date = '2025-12-31'

"""

# =============================================================================
# LOAD DATA
# =============================================================================

df = pd.read_sql(query, engine)

print("Rows loaded:", len(df))

# =============================================================================
# DATA CLEANING
# =============================================================================

# Remove duplicate rows
df = df.drop_duplicates()

# Remove rows without debtor ID
df = df[df["debtor_id"].notna()]

# Convert numeric columns
numeric_cols = [
    "exposure_amount",
    "collateral_value"
]

for col in numeric_cols:
    df[col] = pd.to_numeric(df[col], errors="coerce")

# Replace missing numeric values with zero
df[numeric_cols] = df[numeric_cols].fillna(0)

# Standardise text columns
text_cols = [
    "debtor_name",
    "group_name",
    "country_code",
    "product_type_code"
]

for col in text_cols:
    df[col] = (
        df[col]
        .astype(str)
        .str.strip()
        .str.upper()
    )

# =============================================================================
# CODELIST ENRICHMENT
# =============================================================================
# Create business-friendly descriptions

product_type_map = {
    "1000": "Term Loan",
    "2000": "Overdraft",
    "3000": "Credit Card",
    "4000": "Mortgage",
    "5000": "Leasing"
}

default_status_map = {
    "0": "Performing",
    "1": "Defaulted"
}

forbearance_map = {
    "0": "Non-Forborne",
    "1": "Forborne"
}

# Add description fields
df["product_type_desc"] = (
    df["product_type_code"]
    .map(product_type_map)
    .fillna("Other")
)

df["default_status_desc"] = (
    df["default_status"]
    .astype(str)
    .map(default_status_map)
    .fillna("Unknown")
)

df["forbearance_desc"] = (
    df["forbearance_status"]
    .astype(str)
    .map(forbearance_map)
    .fillna("Unknown")
)

# =============================================================================
# CALCULATED FIELDS
# =============================================================================

# Loan-to-value ratio
df["ltv_ratio"] = (
    df["exposure_amount"] /
    df["collateral_value"].replace(0, pd.NA)
)

# Exposure buckets
df["exposure_bucket"] = pd.cut(
    df["exposure_amount"],
    bins=[0, 1e6, 5e6, 20e6, 100e6, float("inf")],
    labels=[
        "<1m",
        "1m-5m",
        "5m-20m",
        "20m-100m",
        ">100m"
    ]
)

# =============================================================================
# DEBTOR EXPOSURE AGGREGATION
# =============================================================================

debtor_summary = (
    df.groupby(
        ["debtor_id", "debtor_name"],
        as_index=False
    )
    .agg({
        "exposure_amount": "sum",
        "collateral_value": "sum"
    })
)

# Rename columns
debtor_summary = debtor_summary.rename(columns={
    "exposure_amount": "total_exposure",
    "collateral_value": "total_collateral"
})

# Sort descending
debtor_summary = debtor_summary.sort_values(
    by="total_exposure",
    ascending=False
)

# Top 20 debtors
top20_debtors = debtor_summary.head(20)

# =============================================================================
# GROUP EXPOSURE AGGREGATION
# =============================================================================

group_summary = (
    df.groupby(
        ["group_id", "group_name"],
        as_index=False
    )
    .agg({
        "exposure_amount": "sum",
        "collateral_value": "sum"
    })
)

# Rename columns
group_summary = group_summary.rename(columns={
    "exposure_amount": "total_exposure",
    "collateral_value": "total_collateral"
})

# Sort descending
group_summary = group_summary.sort_values(
    by="total_exposure",
    ascending=False
)

# Top 20 groups
top20_groups = group_summary.head(20)

# =============================================================================
# VISUALISATION – TOP 20 DEBTORS
# =============================================================================

plt.figure(figsize=(14, 8))

plt.barh(
    top20_debtors["debtor_name"],
    top20_debtors["total_exposure"]
)

plt.gca().invert_yaxis()

plt.title("Top 20 Debtors by Total Exposure")
plt.xlabel("Exposure Amount")
plt.ylabel("Debtor")

plt.tight_layout()
plt.show()

# =============================================================================
# VISUALISATION – TOP 20 GROUPS
# =============================================================================

plt.figure(figsize=(14, 8))

plt.barh(
    top20_groups["group_name"],
    top20_groups["total_exposure"]
)

plt.gca().invert_yaxis()

plt.title("Top 20 Groups by Total Exposure")
plt.xlabel("Exposure Amount")
plt.ylabel("Group")

plt.tight_layout()
plt.show()

# =============================================================================
# EXPORT RESULTS TO EXCEL
# =============================================================================

output_file = "OSI_Mission_Exposure_Analysis.xlsx"

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:

    df.to_excel(
        writer,
        sheet_name="Cleaned_Loan_Data",
        index=False
    )

    debtor_summary.to_excel(
        writer,
        sheet_name="Debtor_Summary",
        index=False
    )

    group_summary.to_excel(
        writer,
        sheet_name="Group_Summary",
        index=False
    )

print("Analysis completed successfully.")
print(f"Output exported to: {output_file}")

# =============================================================================
# END OF SCRIPT
# =============================================================================
