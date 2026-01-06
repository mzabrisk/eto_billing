# Reading in excel spreadsheet
import pandas as pd
import os 
import calendar

data = 'ETO use.xlsx'

# read in use data
use_df = pd.read_excel(data, sheet_name='eto_use_alt')

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

### Data Cleaning ###
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


# creating new dataframe to total uses and charge

charges_df = pd.DataFrame()
charges_df['Account'] = use_df_expanded['Account'].unique()
charges_df['PI'] = 0
charges_df['Number Uses'] = 0
charges_df['Total ($)'] = 0

charges_df = charges_df.sort_values(by='Account')


# fill in PI by referencng CL codes df
PI = []

for i, row in CL_df.iterrows():
    for j, event in charges_df.iterrows():
        if CL_df['CL_code'][i] == charges_df['Account'][j]:
            PI.append(CL_df['PI'][i])

charges_df['PI'] = PI
charges_df = charges_df.reset_index().drop(columns='index')


# filling in charges df for period of interest

for i, row in charges_df.iterrows():
    owed = []
    num_uses = []
    
    for j, rows in period_of_interest_df.iterrows():
        if charges_df['Account'][i] == period_of_interest_df['Account'][j]:
            num_uses.append(period_of_interest_df['Percent_Charge'][j])

        charges_df['Number Uses'][i] = sum(num_uses)
        
for i, row in charges_df.iterrows():
    charges_df['Total ($)'][i] = charges_df['Number Uses'][i]*40

charges_df = charges_df.round(2)

# drop official use row

charges_df = charges_df[charges_df['Account'] != 'CL000']

# only keep rows with uses/charges
charges_df = charges_df[charges_df['Number Uses'] != 0]


# Create Directory 
directory = f"{year}.{formatted_month_number} Invoicing"
  
# Parent Directory path 
parent_dir = "../eto_billing"
  
# Path 
path = os.path.join(parent_dir, directory) 

if not os.path.exists(path):
    os.makedirs(path) 

charges_df.to_excel(f"{directory}/ethylene_oxide_invoice_{month_name}_{year}.xlsx", index=False)
    