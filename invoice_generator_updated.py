# -*- coding: utf-8 -*-
# Updated ETO/Fluoro invoicing script

import pandas as pd
import os
import calendar

data = 'ETO_Fluoro_Use.xlsx'

# ---------- Read Excel ----------
use_df = pd.read_excel(data, sheet_name='eto_use_alt')      # ETO
fluoro_df = pd.read_excel(data, sheet_name='fluoro_use')    # Fluoro
CL_df = pd.read_excel(data, sheet_name='CL Codes')          # Code-to-PI mapping
CL_df = CL_df.rename(columns={'CL_code': 'Account'})        # normalize column name for merging

# If your department column has a different name, change this:
dept_col = 'Group'
has_dept = dept_col in CL_df.columns

# ---------- Inputs ----------
def month_name_to_number(month_name: str) -> int:
    return list(calendar.month_name).index(month_name.capitalize())

month_name = input("Please enter the month name: ")
year = str(input("Please enter 4-digit year: "))

month_number = month_name_to_number(month_name)
formatted_month_number = f"{month_number:02d}"
formatted_month_number_max = f"{month_number + 1:02d}"

date_min = f"{year}-{formatted_month_number}-01"
if month_name.lower() == 'december':
    date_max = f"{year}-12-31"
else:
    date_max = f"{year}-{formatted_month_number_max}-01"

# ---------- Data Cleaning: ETO ----------
# ---------- Data Cleaning: ETO ----------
use_df['Account'] = use_df['Account'].astype(str).str.replace(' ', '', regex=False)
use_df['Account'] = use_df['Account'].str.split(',')

# number of CL codes on each run (per row)
use_df['Account_Count'] = use_df['Account'].apply(len)

# explode to one row per (run, account)
use_df_expanded = use_df.explode('Account')

# ensure Date is datetime
use_df_expanded['Date'] = pd.to_datetime(use_df_expanded['Date'])

# per-run split: each run contributes 1.0 use, divided by number of accounts on that run
use_df_expanded['Percent_Charge'] = 1 / use_df_expanded['Account_Count']

use_df_expanded = use_df_expanded.drop(columns='Account_Count')


# period subset (monthly invoice)
period_of_interest_df = use_df_expanded[
    (use_df_expanded['Date'] >= date_min) & (use_df_expanded['Date'] < date_max)
]

# ---------- Data Cleaning: FLUORO ----------
fluoro_df['Account'] = fluoro_df['Account'].astype(str).str.replace(' ', '', regex=False)
fluoro_df['Account'] = fluoro_df['Account'].str.split(',')

# number of CL codes on each run (per row)
fluoro_df['Account_Count'] = fluoro_df['Account'].apply(len)

# explode to one row per (run, account)
fluoro_df_expanded = fluoro_df.explode('Account')

fluoro_df_expanded['Date'] = pd.to_datetime(fluoro_df_expanded['Date'])

# per-run split: each run contributes 1.0 use, divided by number of accounts on that run
fluoro_df_expanded['Percent_Charge'] = 1 / fluoro_df_expanded['Account_Count']

fluoro_df_expanded = fluoro_df_expanded.drop(columns='Account_Count')

period_of_interest_fluoro_df = fluoro_df_expanded[
    (fluoro_df_expanded['Date'] >= date_min) & (fluoro_df_expanded['Date'] < date_max)
]

# ---------- Determine accounts to include ----------
# Use the union of ETO + Fluoro accounts that actually appear in the selected period
eto_accounts = period_of_interest_df['Account'].dropna().unique()
fluoro_accounts = period_of_interest_fluoro_df['Account'].dropna().unique()
all_accounts = pd.Index(eto_accounts).union(fluoro_accounts)

# ---------- Seed summary frames from the union ----------
charges_df = pd.DataFrame({'Account': all_accounts})
charges_df['ETO Uses'] = 0.0
charges_df['ETO Dates'] = ''
charges_df['ETO Total ($)'] = 0.0

fluoro_charges_df = pd.DataFrame({'Account': all_accounts})
fluoro_charges_df['Fluoroscopy Uses'] = 0.0
fluoro_charges_df['Fluoroscopy Dates'] = ''
fluoro_charges_df['Fluoroscopy Total ($)'] = 0.0

# ---------- Merge PI (and Department) info ----------
merge_cols = ['Account']
for col in ['PI', dept_col]:
    if col in CL_df.columns:
        merge_cols.append(col)

charges_df = charges_df.merge(CL_df[merge_cols], on='Account', how='left')
fluoro_charges_df = fluoro_charges_df.merge(CL_df[merge_cols], on='Account', how='left')

# ---------- Aggregate ETO (monthly) ----------
if not period_of_interest_df.empty:
    eto_agg = (period_of_interest_df
               .groupby('Account', as_index=False)
               .agg(ETO_Uses=('Percent_Charge', 'sum'),
                    ETO_Dates=('Date', lambda s: ", ".join(sorted({str(pd.to_datetime(d).date()) for d in s})))))
else:
    eto_agg = pd.DataFrame(columns=['Account', 'ETO_Uses', 'ETO_Dates'])

charges_df = charges_df.merge(eto_agg, on='Account', how='left')
charges_df['ETO Uses'] = charges_df['ETO_Uses'].fillna(0)
charges_df['ETO Dates'] = charges_df['ETO_Dates'].fillna('')
charges_df['ETO Total ($)'] = (charges_df['ETO Uses'] * 40).round(2)
charges_df = charges_df.drop(columns=[c for c in ['ETO_Uses', 'ETO_Dates'] if c in charges_df.columns])

# ---------- Aggregate FLUORO (monthly) ----------
if not period_of_interest_fluoro_df.empty:
    fluoro_agg = (
        period_of_interest_fluoro_df
        .groupby('Account', as_index=False)
        .agg(
            Fluoroscopy_Uses=('Percent_Charge', 'sum'),
            Fluoroscopy_Dates=('Date', 
                lambda s: ", ".join(str(pd.to_datetime(d).date()) for d in sorted(s))
            )
        )
    )

else:
    fluoro_agg = pd.DataFrame(columns=['Account', 'Fluoroscopy_Uses', 'Fluoroscopy_Dates'])

fluoro_charges_df = fluoro_charges_df.merge(fluoro_agg, on='Account', how='left')
fluoro_charges_df['Fluoroscopy Uses'] = fluoro_charges_df['Fluoroscopy_Uses'].fillna(0)
fluoro_charges_df['Fluoroscopy Dates'] = fluoro_charges_df['Fluoroscopy_Dates'].fillna('')
fluoro_charges_df['Fluoroscopy Total ($)'] = (fluoro_charges_df['Fluoroscopy Uses'] * 250).round(2)
fluoro_charges_df = fluoro_charges_df.drop(columns=[c for c in ['Fluoroscopy_Uses', 'Fluoroscopy_Dates'] if c in fluoro_charges_df.columns])

# ---------- Drop official code (if present) before merge ----------
charges_df = charges_df[charges_df['Account'] != 'CL000']
fluoro_charges_df = fluoro_charges_df[fluoro_charges_df['Account'] != 'CL000']

# ---------- Merge ETO + Fluoro and compute grand total (monthly) ----------
out = (charges_df
       .merge(
           fluoro_charges_df.drop(columns=[c for c in ['PI', dept_col] if c in fluoro_charges_df.columns]),
           on='Account', how='outer', suffixes=('', '_fluoro'))
      )

# Coalesce PI / Department if needed
for col in ['PI', dept_col]:
    fluoro_col = f"{col}_fluoro"
    if fluoro_col in out.columns:
        if col not in out.columns:
            out[col] = out[fluoro_col]
        else:
            out[col] = out[col].fillna(out[fluoro_col])
        out = out.drop(columns=[fluoro_col])

out['Total ($)'] = out[['ETO Total ($)', 'Fluoroscopy Total ($)']].fillna(0).sum(axis=1).round(2)

# Order columns nicely
cols_order = ['Account']
if 'PI' in out.columns:
    cols_order.append('PI')
if has_dept and dept_col in out.columns:
    cols_order.append(dept_col)
cols_order += [
    'ETO Uses', 'ETO Dates', 'ETO Total ($)',
    'Fluoroscopy Uses', 'Fluoroscopy Dates', 'Fluoroscopy Total ($)',
    'Total ($)'
]
out = out[cols_order].sort_values('Account').reset_index(drop=True)

# ---------- Create directories ----------
parent_dir = "../eto_billing"
year_dir = os.path.join(parent_dir, year)
month_dir_name = f"{year}.{formatted_month_number} Invoicing"
target_dir = os.path.join(year_dir, month_dir_name)

os.makedirs(target_dir, exist_ok=True)

# ---------- Save monthly invoice ----------
outfile = os.path.join(target_dir, f"ethylene_oxide_invoice_{month_name}_{year}.xlsx")
out.to_excel(outfile, index=False)

print(f"Saved invoice: {outfile}")

# ======================================================================
# NEW: Fiscal-year (July–June) running totals by CL and by Department
# ======================================================================

# Figure out which fiscal year this month belongs to
start_of_month = pd.to_datetime(date_min)
if start_of_month.month >= 7:
    fy_start_year = start_of_month.year
else:
    fy_start_year = start_of_month.year - 1

fy_start = pd.Timestamp(fy_start_year, 7, 1)
fy_end = pd.Timestamp(fy_start_year + 1, 6, 30)

# Subset full expanded data to the current FY
fy_eto_df = use_df_expanded[(use_df_expanded['Date'] >= fy_start) & (use_df_expanded['Date'] <= fy_end)]
fy_fluoro_df = fluoro_df_expanded[(fluoro_df_expanded['Date'] >= fy_start) & (fluoro_df_expanded['Date'] <= fy_end)]

# ---------- Aggregate ETO (FY) ----------
if not fy_eto_df.empty:
    fy_eto_agg = (fy_eto_df
                  .groupby('Account', as_index=False)
                  .agg(ETO_Uses=('Percent_Charge', 'sum')))
else:
    fy_eto_agg = pd.DataFrame(columns=['Account', 'ETO_Uses'])

# ---------- Aggregate Fluoro (FY) ----------
if not fy_fluoro_df.empty:
    fy_fluoro_agg = (fy_fluoro_df
                     .groupby('Account', as_index=False)
                     .agg(Fluoroscopy_Uses=('Percent_Charge', 'sum')))
else:
    fy_fluoro_agg = pd.DataFrame(columns=['Account', 'Fluoroscopy_Uses'])

# Combine all accounts seen in FY
fy_accounts = pd.Index(fy_eto_agg['Account'].unique()).union(fy_fluoro_agg['Account'].unique())
fy_cl_df = pd.DataFrame({'Account': fy_accounts})

# Merge CL info (PI / Department)
fy_cl_df = fy_cl_df.merge(CL_df[merge_cols], on='Account', how='left')

# Merge ETO & Fluoro FY uses
fy_cl_df = fy_cl_df.merge(fy_eto_agg, on='Account', how='left')
fy_cl_df = fy_cl_df.merge(fy_fluoro_agg, on='Account', how='left')

fy_cl_df['ETO_Uses'] = fy_cl_df['ETO_Uses'].fillna(0)
fy_cl_df['Fluoroscopy_Uses'] = fy_cl_df['Fluoroscopy_Uses'].fillna(0)

fy_cl_df['ETO Total ($)'] = (fy_cl_df['ETO_Uses'] * 40).round(2)
fy_cl_df['Fluoroscopy Total ($)'] = (fy_cl_df['Fluoroscopy_Uses'] * 250).round(2)
fy_cl_df['Total ($)'] = (fy_cl_df['ETO Total ($)'] + fy_cl_df['Fluoroscopy Total ($)']).round(2)

# Drop official code if needed
fy_cl_df = fy_cl_df[fy_cl_df['Account'] != 'CL000']

# Nice column order for FY CL-level sheet
fy_cols_order = ['Account']
if 'PI' in fy_cl_df.columns:
    fy_cols_order.append('PI')
if has_dept and dept_col in fy_cl_df.columns:
    fy_cols_order.append(dept_col)
fy_cols_order += ['ETO_Uses', 'ETO Total ($)', 'Fluoroscopy_Uses', 'Fluoroscopy Total ($)', 'Total ($)']
fy_cl_df = fy_cl_df[fy_cols_order].sort_values('Account').reset_index(drop=True)

# ---------- FY totals by Department ----------
if has_dept and dept_col in fy_cl_df.columns:
    fy_cl_df[dept_col] = fy_cl_df[dept_col].fillna('Unassigned')

    fy_dept_df = (fy_cl_df
                  .groupby(dept_col, as_index=False)
                  .agg(
                      ETO_Uses=('ETO_Uses', 'sum'),
                      ETO_Total=('ETO Total ($)', 'sum'),
                      Fluoroscopy_Uses=('Fluoroscopy_Uses', 'sum'),
                      Fluoroscopy_Total=('Fluoroscopy Total ($)', 'sum'),
                      Total=('Total ($)', 'sum')
                  )
                  .sort_values(dept_col)
                  .reset_index(drop=True)
                  )
else:
    fy_dept_df = pd.DataFrame()  # no department info available

# ---------- Save / update FY running total workbook ----------
fy_label = fy_start_year + 1  # e.g., FY2025 for July 2024–June 2025
running_file = os.path.join(parent_dir, f"ETO_Fluoro_running_totals_FY{fy_label}.xlsx")

file_exists = os.path.exists(running_file)

if file_exists:
    # Append mode — allow replacing sheets
    with pd.ExcelWriter(
        running_file,
        engine="openpyxl",
        mode="a",
        if_sheet_exists="replace"
    ) as writer:
        fy_cl_df.to_excel(writer, sheet_name="FY_running_by_CL", index=False)

        if not fy_dept_df.empty:
            fy_dept_df.to_excel(writer, sheet_name="FY_running_by_dept", index=False)

else:
    # Write mode — CANNOT use if_sheet_exists
    with pd.ExcelWriter(
        running_file,
        engine="openpyxl",
        mode="w"
    ) as writer:
        fy_cl_df.to_excel(writer, sheet_name="FY_running_by_CL", index=False)

        if not fy_dept_df.empty:
            fy_dept_df.to_excel(writer, sheet_name="FY_running_by_dept", index=False)

print(f"Updated FY running totals workbook: {running_file}")

