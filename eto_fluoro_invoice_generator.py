# Reading in excel spreadsheet
import pandas as pd
import os 
import calendar

data = 'ETO_Fluoro_Use.xlsx'

# read in eto use data
use_df = pd.read_excel(data, sheet_name='eto_use_alt')

# read in fluoro use data
fluoro_df = pd.read_excel(data, sheet_name='fluoro_use')

# read in users data
CL_df = pd.read_excel(data, sheet_name='CL Codes')

### INPUTS ###

# Function to convert month name to its corresponding number
def month_name_to_number(month_name):
    # Get the month number as an integer
    month_number = list(calendar.month_name).index(month_name.capitalize())
    return month_number

# Get input from the user
month_name = input("Please enter the month name: ")
year = str(input("Please enter 4-digit year: "))

# Convert the month name to its corresponding number
month_number = month_name_to_number(month_name)

formatted_month_number = '{:02d}'.format(month_number)
formatted_month_number_max = '{:02d}'.format(month_number + 1)
formatted_month_number = str(formatted_month_number)
formatted_month_number_max = str(formatted_month_number_max)

date_min = f"{year}-{formatted_month_number}-01"

if month_name == 'december':
    date_max = f"{year}-12-31"
else:
    date_max = f"{year}-{formatted_month_number_max}-01"

### Data Cleaning ### ETO
# drop spaces from account lists
use_df['Account'] = use_df['Account'].str.replace(' ', '')

# Split the 'Account' column into a list of accounts
use_df['Account'] = use_df['Account'].str.split(',')

# Use explode() to expand the accounts into separate rows while keeping the associated date
use_df_expanded = use_df.explode('Account')

# determine number of accounts during a given run
date_counts = use_df_expanded.groupby('Date').size().reset_index(name='Account_Count')

# Merge the count information back into the original DataFrame
use_df_expanded = use_df_expanded.merge(date_counts, on='Date')

# assign % charge based off the counts
percent_charge = []
for i, count in use_df_expanded.iterrows():
    percent_charge.append(1 / use_df_expanded['Account_Count'][i])

use_df_expanded['Percent_Charge'] = percent_charge
use_df_expanded = use_df_expanded.drop(columns=('Account_Count'))

# create subset df for period of interest

period_of_interest_df = use_df_expanded[(use_df_expanded['Date'] >= date_min) & (use_df_expanded['Date'] < date_max)]





### Data Cleaning ### FLUORO
# drop spaces from account lists
fluoro_df['Account'] = fluoro_df['Account'].str.replace(' ', '')

# Split the 'Account' column into a list of accounts
fluoro_df['Account'] = fluoro_df['Account'].str.split(',')

# Use explode() to expand the accounts into separate rows while keeping the associated date
fluoro_df_expanded = fluoro_df.explode('Account')

# determine number of accounts during a given run
date_counts_fluoro = fluoro_df_expanded.groupby('Date').size().reset_index(name='Account_Count')

# Merge the count information back into the original DataFrame
fluoro_df_expanded = fluoro_df_expanded.merge(date_counts_fluoro, on='Date')

# assign % charge based off the counts
percent_charge = []
for i, count in fluoro_df_expanded.iterrows():
    percent_charge.append(1 / fluoro_df_expanded['Account_Count'][i])

fluoro_df_expanded['Percent_Charge'] = percent_charge
fluoro_df_expanded = fluoro_df_expanded.drop(columns=('Account_Count'))

# create subset df for period of interest

period_of_interest_fluoro_df = fluoro_df_expanded[(fluoro_df_expanded['Date'] >= date_min) & (fluoro_df_expanded['Date'] < date_max)]


# creating new dataframe to total uses and charge for eto

charges_df = pd.DataFrame()
charges_df['Account'] = use_df_expanded['Account'].unique()
charges_df['PI'] = 0
charges_df['ETO Uses'] = 0
charges_df['ETO Dates'] = str(0)
charges_df['Total ($)'] = 0

charges_df = charges_df.sort_values(by='Account')


# creating new dataframe to total uses and charge for fluoro

fluoro_charges_df = pd.DataFrame()
fluoro_charges_df['Account'] = use_df_expanded['Account'].unique()
fluoro_charges_df['PI'] = 0
fluoro_charges_df['Fluoroscopy Uses'] = 0
fluoro_charges_df['Fluoroscopy Dates'] = str(0)
fluoro_charges_df['Total ($)'] = 0

fluoro_charges_df = fluoro_charges_df.sort_values(by='Account')


# fill in ETO PI by referencng CL codes df
PI = []

for i, row in CL_df.iterrows():
    for j, event in charges_df.iterrows():
        if CL_df['CL_code'][i] == charges_df['Account'][j]:
            PI.append(CL_df['PI'][i])

charges_df['PI'] = PI
charges_df = charges_df.reset_index().drop(columns='index')


# filling in ETO charges df for period of interest

for i, row in charges_df.iterrows():
    owed = []
    num_uses = []
    eto_dates = []
    
    for j, rows in period_of_interest_df.iterrows():
        if charges_df['Account'][i] == period_of_interest_df['Account'][j]:
            num_uses.append(period_of_interest_df['Percent_Charge'][j])
            eto_dates.append(period_of_interest_df['Date'][j])

        charges_df['ETO Uses'][i] = sum(num_uses)
        charges_df['ETO Dates'][i] = ", ".join(str(d.date()) for d in eto_dates)
        
for i, row in charges_df.iterrows():
    charges_df['Total ($)'][i] = charges_df['ETO Uses'][i]*40

charges_df = charges_df.round(2)



# fill in FLUORO PI by referencng CL codes df
PI = []

for i, row in CL_df.iterrows():
    for j, event in fluoro_charges_df.iterrows():
        if CL_df['CL_code'][i] == fluoro_charges_df['Account'][j]:
            PI.append(CL_df['PI'][i])

fluoro_charges_df['PI'] = PI
fluoro_charges_df = fluoro_charges_df.reset_index().drop(columns='index')


# filling in FLUORO charges df for period of interest

for i, row in fluoro_charges_df.iterrows():
    owed = []
    num_uses = []
    fluoro_dates = []
    
    for j, rows in period_of_interest_fluoro_df.iterrows():
        if fluoro_charges_df['Account'][i] == period_of_interest_fluoro_df['Account'][j]:
            num_uses.append(period_of_interest_fluoro_df['Percent_Charge'][j])
            fluoro_dates.append(period_of_interest_fluoro_df['Date'][j])

        fluoro_charges_df['Fluoroscopy Uses'][i] = sum(num_uses)
        fluoro_charges_df['Fluoroscopy Dates'][i] = ", ".join(str(d.date()) for d in fluoro_dates)
        
for i, row in fluoro_charges_df.iterrows():
    fluoro_charges_df['Total ($)'][i] = fluoro_charges_df['Fluoroscopy Uses'][i]*250

fluoro_charges_df = fluoro_charges_df.round(2)


# drop official use row

charges_df = charges_df[charges_df['Account'] != 'CL000']
fluoro_charges_df = fluoro_charges_df[fluoro_charges_df['Account'] != 'CL000']

# only keep rows with uses/charges
charges_df = charges_df[charges_df['ETO Uses'] != 0]
fluoro_charges_df = fluoro_charges_df[fluoro_charges_df['Fluoroscopy Uses'] != 0]

# dropping unnecessary columns pre-merge
# fluoro_charges_df = fluoro_charges_df.drop(columns=['PI'])


charges_df = charges_df.merge(fluoro_charges_df, on='Account', how='outer')

charges_df['PI_x'] = charges_df['PI_x'].combine_first(charges_df['PI_y'])

charges_df = charges_df.drop(columns=['PI_y'])
charges_df.rename(columns={'PI_x': 'PI'}, inplace=True)

# adding fluoro and eto charges
charges_df['Total ($)'] = charges_df['Total ($)_x'].fillna(0) + charges_df['Total ($)_y'].fillna(0)
charges_df = charges_df.drop(columns=['Total ($)_x', 'Total ($)_y'])

# creat year directory
directory_year = year

# Parent Directory path 
parent_dir = "../eto_billing"

# Path 
path = os.path.join(parent_dir, directory_year) 

if not os.path.exists(path):
    os.makedirs(path) 

# Create Directory 
directory = f"{year}/{year}.{formatted_month_number} Invoicing"
  
# Parent Directory path 
parent_dir2 = f"../eto_billing/{year}"
  
# Path 
path2 = os.path.join(parent_dir, directory) 

if not os.path.exists(path2):
    os.makedirs(path2) 

charges_df.to_excel(f"{directory}/ethylene_oxide_invoice_{month_name}_{year}.xlsx", index=False)

    