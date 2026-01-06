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
use_df['Account'] = use_df['Account'].astype(str).str.replace(' ', '', regex=False)
use_df['Account'] = use_df['Account'].str.split(',')
use_df_expanded = use_df.explode('Account')

# per-run account count (for split percentage)
date_counts = use_df_expanded.groupby('Date').size().reset_index(name='Account_Count')
use_df_expanded = use_df_expanded.merge(date_counts, on='Date', how='left')
use_df_expanded['Percent_Charge'] = 1 / use_df_expanded['Account_Count']
use_df_expanded = use_df_expanded.drop(columns='Account_Count')

# period subset
period_of_interest_df = use_df_expanded[
    (use_df_expanded['Date'] >= date_min) & (use_df_expanded['Date'] < date_max)
]

# ---------- Data Cleaning: FLUORO ----------
fluoro_df['Account'] = fluoro_df['Account'].astype(str).str.replace(' ', '', regex=False)
fluoro_df['Account'] = fluoro_df['Account'].str.split(',')
fluoro_df_expanded = fluoro_df.explode('Account')

date_counts_fluoro = fluoro_df_expanded.groupby('Date').size().reset_index(name='Account_Count')
fluoro_df_expanded = fluoro_df_expanded.merge(date_counts_fluoro, on='Date', how='left')
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

# If you prefer to include ALL CL codes regardless of use in the month, replace with:
# all_accounts = CL_df['Account'].dropna().astype(str).unique()

# ---------- Seed summary frames from the union ----------
charges_df = pd.DataFrame({'Account': all_accounts})
charges_df['ETO Uses'] = 0.0
charges_df['ETO Dates'] = ''
charges_df['ETO Total ($)'] = 0.0

fluoro_charges_df = pd.DataFrame({'Account': all_accounts})
fluoro_charges_df['Fluoroscopy Uses'] = 0.0
fluoro_charges_df['Fluoroscopy Dates'] = ''
fluoro_charges_df['Fluoroscopy Total ($)'] = 0.0

# ---------- Merge PI info (left join; may be NaN if code not in CL sheet) ----------
charges_df = charges_df.merge(CL_df[['Account', 'PI']], on='Account', how='left')
fluoro_charges_df = fluoro_charges_df.merge(CL_df[['Account', 'PI']], on='Account', how='left')

# ---------- Aggregate ETO ----------
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

# ---------- Aggregate FLUORO ----------
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

# ---------- Drop official code (if present) before merge ----------
charges_df = charges_df[charges_df['Account'] != 'CL000']
fluoro_charges_df = fluoro_charges_df[fluoro_charges_df['Account'] != 'CL000']

# ---------- Merge ETO + Fluoro and compute grand total ----------
out = (charges_df
       .merge(fluoro_charges_df.drop(columns=['PI']), on='Account', how='outer', suffixes=('', '_fluoro'))
       .rename(columns={'PI': 'PI'})  # keep PI from ETO frame; they match
      )

# If PI is NaN on the left but present on the right (rare), coalesce:
if 'PI_fluoro' in out.columns:
    out['PI'] = out['PI'].fillna(out['PI_fluoro'])
    out = out.drop(columns=[c for c in ['PI_fluoro'] if c in out.columns])

out['Total ($)'] = out[['ETO Total ($)', 'Fluoroscopy Total ($)']].fillna(0).sum(axis=1).round(2)

# Order columns nicely
cols_order = [
    'Account', 'PI',
    'ETO Uses', 'ETO Dates', 'ETO Total ($)',
    'Fluoroscopy Uses', 'Fluoroscopy Dates', 'Fluoroscopy Total ($)',
    'Total ($)'
]
out = out.reindex(columns=cols_order).sort_values('Account').reset_index(drop=True)

# ---------- Create directories and save ----------
parent_dir = "../eto_billing"
year_dir = os.path.join(parent_dir, year)
month_dir_name = f"{year}.{formatted_month_number} Invoicing"
target_dir = os.path.join(year_dir, month_dir_name)

os.makedirs(target_dir, exist_ok=True)

outfile = os.path.join(target_dir, f"ethylene_oxide_invoice_{month_name}_{year}.xlsx")
out.to_excel(outfile, index=False)

print(f"Saved invoice: {outfile}")
