# -*- coding: utf-8 -*-
"""
ETO / Fluoro / PM invoicing script

Updates included:
- Reads monthly PM cost from the Excel tab named "PM Cost".
- Adds a PM000 true-up charge when monthly PM cost exceeds the total monthly
  fluoroscopy charges for that month. ETO charges are not included in the PM
  offset calculation.
- Creates monthly invoice files that include PM000 when a PM true-up is due.
- Creates/updates fiscal-year running totals automatically based on the invoice
  month. Example: July 2026 through June 2027 is FY2027.
"""

import calendar
import os
from typing import Optional

import pandas as pd

# ======================================================================
# Configuration
# ======================================================================

data = "ETO_Fluoro_Use.xlsx"

ETO_RATE = 40.00
FLUORO_RATE = 250.00
PM_ACCOUNT = "PM000"
OFFICIAL_ACCOUNT = "CL000"

ETO_SHEET = "eto_use_alt"
FLUORO_SHEET = "fluoro_use"
CL_CODES_SHEET = "CL Codes"
PM_COST_SHEET = "PM Cost"

dept_col = "Group"  # change this if your department/group column has a different name

parent_dir = "../eto_billing"


# ======================================================================
# Helper functions
# ======================================================================

def month_name_to_number(month_name: str) -> int:
    """Convert a month name to its month number."""
    try:
        return list(calendar.month_name).index(month_name.strip().capitalize())
    except ValueError as exc:
        raise ValueError(f"Invalid month name: {month_name!r}") from exc


def normalize_account_series(series: pd.Series) -> pd.Series:
    """Normalize comma-separated account fields and return lists of accounts."""
    return (
        series
        .fillna("")
        .astype(str)
        .str.replace(" ", "", regex=False)
        .str.split(",")
        .apply(lambda accounts: [acct for acct in accounts if acct and acct.lower() != "nan"])
    )


def expand_usage_df(df: pd.DataFrame, date_col: str = "Date", account_col: str = "Account") -> pd.DataFrame:
    """
    Expand usage records so one shared run becomes one row per account.
    The Percent_Charge column divides each run evenly across listed accounts.
    """
    expanded = df.copy()
    expanded[account_col] = normalize_account_series(expanded[account_col])
    expanded["Account_Count"] = expanded[account_col].apply(len)

    # Drop rows without usable accounts before calculating split charges.
    expanded = expanded[expanded["Account_Count"] > 0].copy()
    expanded = expanded.explode(account_col)
    expanded[date_col] = pd.to_datetime(expanded[date_col])
    expanded["Percent_Charge"] = 1 / expanded["Account_Count"]
    expanded = expanded.drop(columns="Account_Count")
    return expanded


def find_date_column(df: pd.DataFrame) -> str:
    """Find the most likely date/month column in the PM Cost sheet."""
    preferred_names = ["Date", "Month", "Month Start", "Month_Start", "Billing Month", "Billing_Month"]
    for name in preferred_names:
        if name in df.columns:
            return name

    for col in df.columns:
        parsed = pd.to_datetime(df[col], errors="coerce")
        if parsed.notna().sum() > 0:
            return col

    raise ValueError(
        "Could not find a date/month column in the 'PM Cost' sheet. "
        "Use a column named Date, Month, Month Start, or similar."
    )


def find_pm_cost_column(df: pd.DataFrame, date_col: str) -> str:
    """Find the most likely monthly PM cost column in the PM Cost sheet."""
    preferred_names = [
        "PM Cost", "PM_Cost", "PM Charge", "PM_Charge", "Monthly PM Cost",
        "Monthly_PM_Cost", "Cost", "Charge", "Amount", "Total"
    ]
    for name in preferred_names:
        if name in df.columns and name != date_col:
            return name

    numeric_candidates = []
    for col in df.columns:
        if col == date_col:
            continue
        numeric_values = pd.to_numeric(df[col], errors="coerce")
        if numeric_values.notna().sum() > 0:
            numeric_candidates.append(col)

    if len(numeric_candidates) == 1:
        return numeric_candidates[0]

    if len(numeric_candidates) > 1:
        raise ValueError(
            "Found more than one possible PM cost column in the 'PM Cost' sheet: "
            f"{numeric_candidates}. Rename the intended column to 'PM Cost'."
        )

    raise ValueError(
        "Could not find a numeric PM cost column in the 'PM Cost' sheet. "
        "Use a column named 'PM Cost' or similar."
    )


def get_monthly_pm_cost(pm_cost_df: pd.DataFrame, start_of_month: pd.Timestamp) -> float:
    """Return the PM cost for the invoice month from the PM Cost sheet."""
    if pm_cost_df.empty:
        return 0.0

    pm_df = pm_cost_df.copy()
    date_col = find_date_column(pm_df)
    cost_col = find_pm_cost_column(pm_df, date_col)

    pm_df[date_col] = pd.to_datetime(pm_df[date_col], errors="coerce")
    pm_df[cost_col] = pd.to_numeric(pm_df[cost_col], errors="coerce")
    pm_df = pm_df.dropna(subset=[date_col, cost_col]).copy()

    # Normalize all dates to first day of their month.
    pm_df["PM_Month"] = pm_df[date_col].dt.to_period("M").dt.to_timestamp()
    target_month = start_of_month.to_period("M").to_timestamp()

    matches = pm_df[pm_df["PM_Month"] == target_month]
    if matches.empty:
        return 0.0

    # Sum allows the sheet to contain more than one PM line for a month.
    return float(matches[cost_col].sum())


def get_fy_label_and_bounds(start_of_month: pd.Timestamp):
    """Return FY label, FY start, and FY end for a July-June fiscal year."""
    if start_of_month.month >= 7:
        fy_start_year = start_of_month.year
    else:
        fy_start_year = start_of_month.year - 1

    fy_start = pd.Timestamp(fy_start_year, 7, 1)
    fy_end_exclusive = pd.Timestamp(fy_start_year + 1, 7, 1)
    fy_label = fy_start_year + 1
    return fy_label, fy_start, fy_end_exclusive


def build_pm_account_row(
    out_columns,
    pm_true_up: float,
    pm_monthly_cost: float,
    pm_fluoro_offset: float,
    invoice_month_label: str,
    cl_df: pd.DataFrame,
    merge_cols,
    has_dept: bool,
) -> Optional[pd.DataFrame]:
    """Build the PM000 row for the monthly invoice, if a true-up is due."""
    if pm_true_up <= 0:
        return None

    pm_info = cl_df[cl_df["Account"] == PM_ACCOUNT]

    row = {col: 0 for col in out_columns}
    row["Account"] = PM_ACCOUNT
    row["ETO Uses"] = 0.0
    row["ETO Dates"] = ""
    row["ETO Total ($)"] = 0.0
    row["Fluoroscopy Uses"] = 0.0
    row["Fluoroscopy Dates"] = f"PM true-up for {invoice_month_label}"
    row["Fluoroscopy Total ($)"] = 0.0
    row["PM Monthly Cost ($)"] = round(pm_monthly_cost, 2)
    row["PM Fluoro Offset ($)"] = round(pm_fluoro_offset, 2)
    row["PM True-Up ($)"] = round(pm_true_up, 2)
    row["Total ($)"] = round(pm_true_up, 2)

    if not pm_info.empty:
        for col in merge_cols:
            if col != "Account" and col in out_columns and col in pm_info.columns:
                row[col] = pm_info.iloc[0][col]

    if has_dept and dept_col in out_columns and (not row.get(dept_col)):
        row[dept_col] = "Unassigned"

    return pd.DataFrame([row], columns=out_columns)


# ======================================================================
# Read Excel
# ======================================================================

use_df = pd.read_excel(data, sheet_name=ETO_SHEET)
fluoro_df = pd.read_excel(data, sheet_name=FLUORO_SHEET)
CL_df = pd.read_excel(data, sheet_name=CL_CODES_SHEET)
pm_cost_df = pd.read_excel(data, sheet_name=PM_COST_SHEET)

# Normalize CL-code column name for merging.
CL_df = CL_df.rename(columns={"CL_code": "Account"})
CL_df["Account"] = CL_df["Account"].astype(str).str.replace(" ", "", regex=False)

has_dept = dept_col in CL_df.columns

merge_cols = ["Account"]
for col in ["PI", dept_col]:
    if col in CL_df.columns:
        merge_cols.append(col)


# ======================================================================
# Inputs and date range
# ======================================================================

month_name = input("Please enter the month name: ").strip()
year = str(input("Please enter 4-digit year: ")).strip()

month_number = month_name_to_number(month_name)
formatted_month_number = f"{month_number:02d}"

start_of_month = pd.Timestamp(int(year), month_number, 1)
start_of_next_month = start_of_month + pd.DateOffset(months=1)

date_min = start_of_month
date_max = start_of_next_month
invoice_month_label = start_of_month.strftime("%B %Y")


# ======================================================================
# Data cleaning and monthly subsets
# ======================================================================

use_df_expanded = expand_usage_df(use_df)
fluoro_df_expanded = expand_usage_df(fluoro_df)

period_of_interest_df = use_df_expanded[
    (use_df_expanded["Date"] >= date_min) & (use_df_expanded["Date"] < date_max)
].copy()

period_of_interest_fluoro_df = fluoro_df_expanded[
    (fluoro_df_expanded["Date"] >= date_min) & (fluoro_df_expanded["Date"] < date_max)
].copy()


# ======================================================================
# Monthly invoice by account
# ======================================================================

eto_accounts = period_of_interest_df["Account"].dropna().unique()
fluoro_accounts = period_of_interest_fluoro_df["Account"].dropna().unique()
all_accounts = pd.Index(eto_accounts).union(fluoro_accounts)

charges_df = pd.DataFrame({"Account": all_accounts})
charges_df["ETO Uses"] = 0.0
charges_df["ETO Dates"] = ""
charges_df["ETO Total ($)"] = 0.0

fluoro_charges_df = pd.DataFrame({"Account": all_accounts})
fluoro_charges_df["Fluoroscopy Uses"] = 0.0
fluoro_charges_df["Fluoroscopy Dates"] = ""
fluoro_charges_df["Fluoroscopy Total ($)"] = 0.0

charges_df = charges_df.merge(CL_df[merge_cols], on="Account", how="left")
fluoro_charges_df = fluoro_charges_df.merge(CL_df[merge_cols], on="Account", how="left")

# ---------- Aggregate ETO ----------
if not period_of_interest_df.empty:
    eto_agg = (
        period_of_interest_df
        .groupby("Account", as_index=False)
        .agg(
            ETO_Uses=("Percent_Charge", "sum"),
            ETO_Dates=("Date", lambda s: ", ".join(sorted({str(pd.to_datetime(d).date()) for d in s}))),
        )
    )
else:
    eto_agg = pd.DataFrame(columns=["Account", "ETO_Uses", "ETO_Dates"])

charges_df = charges_df.merge(eto_agg, on="Account", how="left")
charges_df["ETO Uses"] = charges_df["ETO_Uses"].fillna(0)
charges_df["ETO Dates"] = charges_df["ETO_Dates"].fillna("")
charges_df["ETO Total ($)"] = (charges_df["ETO Uses"] * ETO_RATE).round(2)
charges_df = charges_df.drop(columns=[c for c in ["ETO_Uses", "ETO_Dates"] if c in charges_df.columns])

# ---------- Aggregate Fluoro ----------
if not period_of_interest_fluoro_df.empty:
    fluoro_agg = (
        period_of_interest_fluoro_df
        .groupby("Account", as_index=False)
        .agg(
            Fluoroscopy_Uses=("Percent_Charge", "sum"),
            Fluoroscopy_Dates=("Date", lambda s: ", ".join(sorted({str(pd.to_datetime(d).date()) for d in s}))),
        )
    )
else:
    fluoro_agg = pd.DataFrame(columns=["Account", "Fluoroscopy_Uses", "Fluoroscopy_Dates"])

fluoro_charges_df = fluoro_charges_df.merge(fluoro_agg, on="Account", how="left")
fluoro_charges_df["Fluoroscopy Uses"] = fluoro_charges_df["Fluoroscopy_Uses"].fillna(0)
fluoro_charges_df["Fluoroscopy Dates"] = fluoro_charges_df["Fluoroscopy_Dates"].fillna("")
fluoro_charges_df["Fluoroscopy Total ($)"] = (fluoro_charges_df["Fluoroscopy Uses"] * FLUORO_RATE).round(2)
fluoro_charges_df = fluoro_charges_df.drop(
    columns=[c for c in ["Fluoroscopy_Uses", "Fluoroscopy_Dates"] if c in fluoro_charges_df.columns]
)

# Drop official code before merging. PM000 is intentionally retained/added below.
charges_df = charges_df[charges_df["Account"] != OFFICIAL_ACCOUNT]
fluoro_charges_df = fluoro_charges_df[fluoro_charges_df["Account"] != OFFICIAL_ACCOUNT]

out = charges_df.merge(
    fluoro_charges_df.drop(columns=[c for c in ["PI", dept_col] if c in fluoro_charges_df.columns]),
    on="Account",
    how="outer",
    suffixes=("", "_fluoro"),
)

for col in ["PI", dept_col]:
    fluoro_col = f"{col}_fluoro"
    if fluoro_col in out.columns:
        if col not in out.columns:
            out[col] = out[fluoro_col]
        else:
            out[col] = out[col].fillna(out[fluoro_col])
        out = out.drop(columns=[fluoro_col])

# ---------- PM true-up for monthly invoice ----------
monthly_pm_cost = round(get_monthly_pm_cost(pm_cost_df, start_of_month), 2)
monthly_fluoro_offset = round(out["Fluoroscopy Total ($)"].fillna(0).sum(), 2) if "Fluoroscopy Total ($)" in out.columns else 0.0
monthly_pm_true_up = round(max(monthly_pm_cost - monthly_fluoro_offset, 0), 2)

out["PM Monthly Cost ($)"] = 0.0
out["PM Fluoro Offset ($)"] = 0.0
out["PM True-Up ($)"] = 0.0

out["Total ($)"] = out[["ETO Total ($)", "Fluoroscopy Total ($)", "PM True-Up ($)"]].fillna(0).sum(axis=1).round(2)

cols_order = ["Account"]
if "PI" in out.columns:
    cols_order.append("PI")
if has_dept and dept_col in out.columns:
    cols_order.append(dept_col)
cols_order += [
    "ETO Uses", "ETO Dates", "ETO Total ($)",
    "Fluoroscopy Uses", "Fluoroscopy Dates", "Fluoroscopy Total ($)",
    "PM Monthly Cost ($)", "PM Fluoro Offset ($)", "PM True-Up ($)",
    "Total ($)",
]
out = out[cols_order]

pm_row = build_pm_account_row(
    out_columns=out.columns,
    pm_true_up=monthly_pm_true_up,
    pm_monthly_cost=monthly_pm_cost,
    pm_fluoro_offset=monthly_fluoro_offset,
    invoice_month_label=invoice_month_label,
    cl_df=CL_df,
    merge_cols=merge_cols,
    has_dept=has_dept,
)
if pm_row is not None:
    # Avoid duplicate PM000 rows if PM000 appeared in raw usage by mistake.
    out = out[out["Account"] != PM_ACCOUNT]
    out = pd.concat([out, pm_row], ignore_index=True)

out = out.sort_values("Account").reset_index(drop=True)


# ======================================================================
# Save monthly invoice
# ======================================================================

year_dir = os.path.join(parent_dir, year)
month_dir_name = f"{year}.{formatted_month_number} Invoicing"
target_dir = os.path.join(year_dir, month_dir_name)
os.makedirs(target_dir, exist_ok=True)

outfile = os.path.join(target_dir, f"ethylene_oxide_invoice_{month_name}_{year}.xlsx")
out.to_excel(outfile, index=False)

print(f"Saved invoice: {outfile}")
print(f"Monthly PM cost: ${monthly_pm_cost:,.2f}")
print(f"Monthly fluoroscopy offset: ${monthly_fluoro_offset:,.2f}")
print(f"PM000 true-up charge: ${monthly_pm_true_up:,.2f}")


# ======================================================================
# Fiscal-year running totals by CL and by Department
# ======================================================================

fy_label, fy_start, fy_end_exclusive = get_fy_label_and_bounds(start_of_month)

fy_eto_df = use_df_expanded[
    (use_df_expanded["Date"] >= fy_start) & (use_df_expanded["Date"] < fy_end_exclusive)
].copy()
fy_fluoro_df = fluoro_df_expanded[
    (fluoro_df_expanded["Date"] >= fy_start) & (fluoro_df_expanded["Date"] < fy_end_exclusive)
].copy()

if not fy_eto_df.empty:
    fy_eto_agg = fy_eto_df.groupby("Account", as_index=False).agg(ETO_Uses=("Percent_Charge", "sum"))
else:
    fy_eto_agg = pd.DataFrame(columns=["Account", "ETO_Uses"])

if not fy_fluoro_df.empty:
    fy_fluoro_agg = fy_fluoro_df.groupby("Account", as_index=False).agg(Fluoroscopy_Uses=("Percent_Charge", "sum"))
else:
    fy_fluoro_agg = pd.DataFrame(columns=["Account", "Fluoroscopy_Uses"])

fy_accounts = pd.Index(fy_eto_agg["Account"].unique()).union(fy_fluoro_agg["Account"].unique())
fy_cl_df = pd.DataFrame({"Account": fy_accounts})
fy_cl_df = fy_cl_df.merge(CL_df[merge_cols], on="Account", how="left")
fy_cl_df = fy_cl_df.merge(fy_eto_agg, on="Account", how="left")
fy_cl_df = fy_cl_df.merge(fy_fluoro_agg, on="Account", how="left")

fy_cl_df["ETO_Uses"] = fy_cl_df["ETO_Uses"].fillna(0)
fy_cl_df["Fluoroscopy_Uses"] = fy_cl_df["Fluoroscopy_Uses"].fillna(0)
fy_cl_df["ETO Total ($)"] = (fy_cl_df["ETO_Uses"] * ETO_RATE).round(2)
fy_cl_df["Fluoroscopy Total ($)"] = (fy_cl_df["Fluoroscopy_Uses"] * FLUORO_RATE).round(2)
fy_cl_df["PM True-Up ($)"] = 0.0

# ---------- PM true-up for FY running totals ----------
# For each PM Cost month in the current FY, compare that month's PM cost to
# that month's fluoroscopy charges only. ETO never offsets PM cost.
pm_for_fy = pm_cost_df.copy()
try:
    pm_date_col = find_date_column(pm_for_fy)
    pm_cost_col = find_pm_cost_column(pm_for_fy, pm_date_col)
    pm_for_fy[pm_date_col] = pd.to_datetime(pm_for_fy[pm_date_col], errors="coerce")
    pm_for_fy[pm_cost_col] = pd.to_numeric(pm_for_fy[pm_cost_col], errors="coerce")
    pm_for_fy = pm_for_fy.dropna(subset=[pm_date_col, pm_cost_col]).copy()
    pm_for_fy["PM_Month"] = pm_for_fy[pm_date_col].dt.to_period("M").dt.to_timestamp()
    pm_for_fy = pm_for_fy[(pm_for_fy["PM_Month"] >= fy_start) & (pm_for_fy["PM_Month"] < fy_end_exclusive)]

    monthly_pm_costs_fy = pm_for_fy.groupby("PM_Month", as_index=False)[pm_cost_col].sum()

    if not fy_fluoro_df.empty:
        fy_fluoro_monthly = fy_fluoro_df.copy()
        fy_fluoro_monthly["PM_Month"] = fy_fluoro_monthly["Date"].dt.to_period("M").dt.to_timestamp()
        fy_fluoro_monthly = (
            fy_fluoro_monthly
            .groupby("PM_Month", as_index=False)
            .agg(Fluoro_Offset=("Percent_Charge", "sum"))
        )
        fy_fluoro_monthly["Fluoro_Offset"] = fy_fluoro_monthly["Fluoro_Offset"] * FLUORO_RATE
    else:
        fy_fluoro_monthly = pd.DataFrame(columns=["PM_Month", "Fluoro_Offset"])

    fy_pm_monthly = monthly_pm_costs_fy.merge(fy_fluoro_monthly, on="PM_Month", how="left")
    fy_pm_monthly["Fluoro_Offset"] = fy_pm_monthly["Fluoro_Offset"].fillna(0)
    fy_pm_monthly["PM True-Up ($)"] = (fy_pm_monthly[pm_cost_col] - fy_pm_monthly["Fluoro_Offset"]).clip(lower=0).round(2)
    fy_pm_true_up_total = round(float(fy_pm_monthly["PM True-Up ($)"].sum()), 2)
except ValueError:
    fy_pm_monthly = pd.DataFrame()
    fy_pm_true_up_total = 0.0

if fy_pm_true_up_total > 0:
    # Ensure PM000 is present in the CL-level FY sheet.
    if PM_ACCOUNT in fy_cl_df["Account"].values:
        fy_cl_df.loc[fy_cl_df["Account"] == PM_ACCOUNT, "PM True-Up ($)"] = fy_pm_true_up_total
    else:
        pm_info = CL_df[CL_df["Account"] == PM_ACCOUNT]
        pm_fy_row = {col: None for col in fy_cl_df.columns}
        pm_fy_row["Account"] = PM_ACCOUNT
        pm_fy_row["ETO_Uses"] = 0.0
        pm_fy_row["Fluoroscopy_Uses"] = 0.0
        pm_fy_row["ETO Total ($)"] = 0.0
        pm_fy_row["Fluoroscopy Total ($)"] = 0.0
        pm_fy_row["PM True-Up ($)"] = fy_pm_true_up_total
        if not pm_info.empty:
            for col in merge_cols:
                if col != "Account" and col in pm_fy_row and col in pm_info.columns:
                    pm_fy_row[col] = pm_info.iloc[0][col]
        fy_cl_df = pd.concat([fy_cl_df, pd.DataFrame([pm_fy_row])], ignore_index=True)

fy_cl_df["Total ($)"] = (
    fy_cl_df["ETO Total ($)"].fillna(0)
    + fy_cl_df["Fluoroscopy Total ($)"].fillna(0)
    + fy_cl_df["PM True-Up ($)"].fillna(0)
).round(2)

fy_cl_df = fy_cl_df[fy_cl_df["Account"] != OFFICIAL_ACCOUNT]

fy_cols_order = ["Account"]
if "PI" in fy_cl_df.columns:
    fy_cols_order.append("PI")
if has_dept and dept_col in fy_cl_df.columns:
    fy_cols_order.append(dept_col)
fy_cols_order += [
    "ETO_Uses", "ETO Total ($)",
    "Fluoroscopy_Uses", "Fluoroscopy Total ($)",
    "PM True-Up ($)", "Total ($)",
]
fy_cl_df = fy_cl_df[fy_cols_order].sort_values("Account").reset_index(drop=True)

if has_dept and dept_col in fy_cl_df.columns:
    fy_cl_df[dept_col] = fy_cl_df[dept_col].fillna("Unassigned")
    fy_dept_df = (
        fy_cl_df
        .groupby(dept_col, as_index=False)
        .agg(
            ETO_Uses=("ETO_Uses", "sum"),
            ETO_Total=("ETO Total ($)", "sum"),
            Fluoroscopy_Uses=("Fluoroscopy_Uses", "sum"),
            Fluoroscopy_Total=("Fluoroscopy Total ($)", "sum"),
            PM_True_Up=("PM True-Up ($)", "sum"),
            Total=("Total ($)", "sum"),
        )
        .sort_values(dept_col)
        .reset_index(drop=True)
    )
else:
    fy_dept_df = pd.DataFrame()

running_file = os.path.join(parent_dir, f"ETO_Fluoro_running_totals_FY{fy_label}.xlsx")
file_exists = os.path.exists(running_file)

if file_exists:
    with pd.ExcelWriter(running_file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        fy_cl_df.to_excel(writer, sheet_name="FY_running_by_CL", index=False)
        if not fy_dept_df.empty:
            fy_dept_df.to_excel(writer, sheet_name="FY_running_by_dept", index=False)
        if not fy_pm_monthly.empty:
            fy_pm_monthly.to_excel(writer, sheet_name="FY_PM_monthly_detail", index=False)
else:
    with pd.ExcelWriter(running_file, engine="openpyxl", mode="w") as writer:
        fy_cl_df.to_excel(writer, sheet_name="FY_running_by_CL", index=False)
        if not fy_dept_df.empty:
            fy_dept_df.to_excel(writer, sheet_name="FY_running_by_dept", index=False)
        if not fy_pm_monthly.empty:
            fy_pm_monthly.to_excel(writer, sheet_name="FY_PM_monthly_detail", index=False)

print(f"Updated FY running totals workbook: {running_file}")
print(f"Fiscal year label: FY{fy_label}")
