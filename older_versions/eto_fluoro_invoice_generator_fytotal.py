# -*- coding: utf-8 -*-
# Updated ETO/Fluoro invoicing script with fiscal-year running totals

import pandas as pd
import os
import calendar
from datetime import datetime

data = 'ETO_Fluoro_Use.xlsx'

# ---------- Read Excel ----------
use_df = pd.read_excel(data, sheet_name='eto_use_alt')      # ETO
fluoro_df = pd.read_excel(data, sheet_name='fluoro_use')    # Fluoro
CL_df = pd.read_excel(data, sheet_name='CL Codes')          # Code-to-PI mapping
CL_df = CL_df.rename(columns={'CL_code': 'Account'})        # normalize column name for merging

# ---------- Inputs ----------
def month_name_to_number(month_name: str) -> int:
    return list(calendar.month_name).index(month_name.capitalize())

month_name = input("Please enter the month name: ")
year = str(input("Please enter 4-digit year: "))

month_number = month_name_to_number(month_name)
formatted_month_number = f"{month_number:02d}"
formatted_month_number_max = f"{month_number + 1:02d}"

# We'll work with datetime for robust comparisons
date_min = pd.to_datetime(f"{year}-{formatted_month_number}-01")

# exclusive upper bound (first day of next month)
if month_name.lower() == 'december':
    date_max = pd.to_datetime(f"{year}-12-31") + pd.offsets.Day(1)  # 1 day past end -> exclusive
else:
    date_max = pd.to_datetime(f"{year}-{formatted_month_number_max}-01")

# ---------- Helper: fiscal year start (July 1) and label ----------
def fiscal_year_start(dt: pd.Timestamp) -> pd.Timestamp:
    # If month >= July, FY starts July 1 of same year; else July 1 of prev year
    if dt.month >= 7:
        return pd.Timestamp(year=dt.year, month=7, day=1)
    else:
        return pd.Timestamp(year=dt.year - 1, month=7, day=1)

def fiscal_year_label(dt: pd.Timestamp) -> str:
    # FY YYYY (e.g., July 2025–June 2026 -> "FY2026")
    if dt.month >= 7:
        return f"FY{dt.year + 1}"
    else:
        return f"FY{dt.year}"

fy_start = fiscal_year_start(date_min)
fy_label = fiscal_year_label(date_min)

# ---------- Clean & explode: ETO ----------
use_df['Date'] = pd.to_datetime(use_df['Date'])
use_df['Account'] = use_df['Account'].astype(str).str.replace(' ', '', regex=False)
use_df['Account'] = use_df['Account'].str.split(',')
use_df_expanded = use_df.explode('Account')

# per-run account count (for split percentage)
date_counts = use_df_expanded.groupby('Date').size().reset_index(name='Account_Count')
use_df_expanded = use_df_expanded.merge(date_counts, on='Date', how='left')
use_df_expanded['Percent_Charge'] = 1 / use_df_expanded['Account_Count']
use_df_expanded = use_df_expanded.drop(columns='Account_Count')

# month subset
period_of_interest_df = use_df_expanded[
    (use_df_expanded['Date'] >= date_min) & (use_df_expanded['Date'] < date_max)
].copy()

# ---------- Clean & explode: Fluoro ----------
fluoro_df['Date'] = pd.to_datetime(fluoro_df['Date'])
fluoro_df['Account'] = fluoro_df['Account'].astype(str).str.replace(' ', '', regex=False)
fluoro_df['Account'] = fluoro_df['Account'].str.split(',')
fluoro_df_expanded = fluoro_df.explode('Account')

date_counts_fluoro = fluoro_df_expanded.groupby('Date').size().reset_index(name='Account_Count')
fluoro_df_expanded = fluoro_df_expanded.merge(date_counts_fluoro, on='Date', how='left')
fluoro_df_expanded['Percent_Charge'] = 1 / fluoro_df_expanded['Account_Count']
fluoro_df_expanded = fluoro_df_expanded.drop(columns='Account_Count')

period_of_interest_fluoro_df = fluoro_df_expanded[
    (fluoro_df_expanded['Date'] >= date_min) & (fluoro_df_expanded['Date'] < date_max)
].copy()

# ---------- Determine accounts to include (only those seen in the month) ----------
eto_accounts = period_of_interest_df['Account'].dropna().unique()
fluoro_accounts = period_of_interest_fluoro_df['Account'].dropna().unique()
all_accounts = pd.Index(eto_accounts).union(fluoro_accounts)

# ---------- Seed summary frames ----------
charges_df = pd.DataFrame({'Account': all_accounts})
charges_df['ETO Uses'] = 0.0
charges_df['ETO Dates'] = ''
charges_df['ETO Total ($)'] = 0.0

fluoro_charges_df = pd.DataFrame({'Account': all_accounts})
fluoro_charges_df['Fluoroscopy Uses'] = 0.0
fluoro_charges_df['Fluoroscopy Dates'] = ''
fluoro_charges_df['Fluoroscopy Total ($)'] = 0.0

# ---------- Merge PI info ----------
charges_df = charges_df.merge(CL_df[['Account', 'PI']], on='Account', how='left')
fluoro_charges_df = fluoro_charges_df.merge(CL_df[['Account', 'PI']], on='Account', how='left')

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

# ---------- Aggregate Fluoro (monthly) ----------
if not period_of_interest_fluoro_df.empty:
    fluoro_agg = (period_of_interest_fluoro_df
                  .groupby('Account', as_index=False)
                  .agg(Fluoroscopy_Uses=('Percent_Charge', 'sum'),
                       Fluoroscopy_Dates=('Date', lambda s: ", ".join(sorted({str(pd.to_datetime(d).date()) for d in s})))))
else:
    fluoro_agg = pd.DataFrame(columns=['Account', 'Fluoroscopy_Uses', 'Fluoroscopy_Dates'])

fluoro_charges_df = fluoro_charges_df.merge(fluoro_agg, on='Account', how='left')
fluoro_charges_df['Fluoroscopy Uses'] = fluoro_charges_df['Fluoroscopy_Uses'].fillna(0)
fluoro_charges_df['Fluoroscopy Dates'] = fluoro_charges_df['Fluoroscopy_Dates'].fillna('')
fluoro_charges_df['Fluoroscopy Total ($)'] = (fluoro_charges_df['Fluoroscopy Uses'] * 250).round(2)
fluoro_charges_df = fluoro_charges_df.drop(columns=[c for c in ['Fluoroscopy_Uses', 'Fluoroscopy_Dates'] if c in fluoro_charges_df.columns])

# ---------- Drop official code (if present) ----------
charges_df = charges_df[charges_df['Account'] != 'CL000']
fluoro_charges_df = fluoro_charges_df[fluoro_charges_df['Account'] != 'CL000']

# ---------- Merge monthly ETO + Fluoro and compute grand total ----------
out = (charges_df
       .merge(fluoro_charges_df.drop(columns=['PI']), on='Account', how='outer', suffixes=('', '_fluoro'))
       .rename(columns={'PI': 'PI'}))

if 'PI_fluoro' in out.columns:
    out['PI'] = out['PI'].fillna(out['PI_fluoro'])
    out = out.drop(columns=[c for c in ['PI_fluoro'] if c in out.columns])

out['Total ($)'] = out[['ETO Total ($)', 'Fluoroscopy Total ($)']].fillna(0).sum(axis=1).round(2)

cols_order = [
    'Account', 'PI',
    'ETO Uses', 'ETO Dates', 'ETO Total ($)',
    'Fluoroscopy Uses', 'Fluoroscopy Dates', 'Fluoroscopy Total ($)',
    'Total ($)'
]
out = out.reindex(columns=cols_order).sort_values('Account').reset_index(drop=True)

# ---------- Fiscal-year RUNNING TOTALS (July -> June) ----------
# Filter full datasets to FY-to-date (from FY start through the end of the selected month)
fy_end_exclusive = date_max  # through selected month
eto_fy = use_df_expanded[(use_df_expanded['Date'] >= fy_start) & (use_df_expanded['Date'] < fy_end_exclusive)].copy()
fluoro_fy = fluoro_df_expanded[(fluoro_df_expanded['Date'] >= fy_start) & (fluoro_df_expanded['Date'] < fy_end_exclusive)].copy()

# Aggregate by Account
eto_fy_agg = (eto_fy.groupby('Account', as_index=False)
              .agg(ETO_Uses_YTD=('Percent_Charge', 'sum')))
fluoro_fy_agg = (fluoro_fy.groupby('Account', as_index=False)
                 .agg(Fluoroscopy_Uses_YTD=('Percent_Charge', 'sum')))

# Build account set from FY data (you can switch to CL_df['Account'] to include all)
fy_accounts = pd.Index(eto_fy_agg['Account']).union(fluoro_fy_agg['Account'])
running_df = pd.DataFrame({'Account': fy_accounts}).merge(CL_df[['Account', 'PI']], on='Account', how='left')

running_df = (running_df
              .merge(eto_fy_agg, on='Account', how='left')
              .merge(fluoro_fy_agg, on='Account', how='left'))

running_df['ETO_Uses_YTD'] = running_df['ETO_Uses_YTD'].fillna(0).round(2)
running_df['Fluoroscopy_Uses_YTD'] = running_df['Fluoroscopy_Uses_YTD'].fillna(0).round(2)

running_df['ETO_Total_YTD_($)'] = (running_df['ETO_Uses_YTD'] * 40).round(2)
running_df['Fluoroscopy_Total_YTD_($)'] = (running_df['Fluoroscopy_Uses_YTD'] * 250).round(2)
running_df['Grand_Total_YTD_($)'] = (running_df['ETO_Total_YTD_($)'] +
                                     running_df['Fluoroscopy_Total_YTD_($)']).round(2)

running_cols = [
    'Account', 'PI',
    'ETO_Uses_YTD', 'ETO_Total_YTD_($)',
    'Fluoroscopy_Uses_YTD', 'Fluoroscopy_Total_YTD_($)',
    'Grand_Total_YTD_($)'
]
running_df = running_df.reindex(columns=running_cols).sort_values('Account').reset_index(drop=True)

# ---------- Create directories ----------
parent_dir = "../eto_billing"
year_dir = os.path.join(parent_dir, year)
month_dir_name = f"{year}.{formatted_month_number} Invoicing"
target_dir = os.path.join(year_dir, month_dir_name)
os.makedirs(target_dir, exist_ok=True)

# ---------- Save monthly invoice (existing behavior) ----------
invoice_xlsx = os.path.join(target_dir, f"ethylene_oxide_invoice_{month_name}_{year}.xlsx")
out.to_excel(invoice_xlsx, index=False)

# ---------- Save RUNNING TOTALS ----------
# 1) Excel (sheet per modality summary combined)
running_xlsx_name = f"running_totals_{fy_label}_through_{year}-{formatted_month_number}.xlsx"
running_xlsx_path = os.path.join(target_dir, running_xlsx_name)
with pd.ExcelWriter(running_xlsx_path, engine='xlsxwriter') as xw:
    running_df.to_excel(xw, sheet_name='RunningTotals', index=False)

# 2) CSV (same content)
running_csv_name = f"running_totals_{fy_label}_through_{year}-{formatted_month_number}.csv"
running_csv_path = os.path.join(target_dir, running_csv_name)
running_df.to_csv(running_csv_path, index=False)

print(f"Saved invoice: {invoice_xlsx}")
print(f"Saved FY running totals (Excel): {running_xlsx_path}")
print(f"Saved FY running totals (CSV):   {running_csv_path}")
