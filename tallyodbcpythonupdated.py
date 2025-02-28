import pyodbc
import warnings
import pandas as pd
import time
from datetime import datetime
import pathlib
import tkinter as tk
from tkinter import ttk, filedialog
import subprocess
import os


warnings.filterwarnings('ignore')


def VendorPartyName(input_df):
        df = input_df.copy()  # Create a copy of the input dataframe to avoid modifying the original
        
        vendor_filter = df['$Led_Group'].isin(['Sundry Debtors', 'Sundry Creditors'])
        vendor_df = df[vendor_filter]
        
        vendor_party_name = vendor_df.groupby('$Key')['$LedgerName'].apply(lambda x: ', '.join(x.unique())).reset_index()
        vendor_party_name.rename(columns={'$LedgerName': 'VendorPartyName'}, inplace=True)
        
        df = df.merge(vendor_party_name, on='$Key', how='left')
        
        return df

def ExpenseName(input_df):
        df = input_df.copy()  # Create a copy of the input dataframe to avoid modifying the original
        
        vendor_filter = df['$Led_Group'].isin(['Direct Expenses', 'Purchases','Indirect Expenses','Fixed Assets'])
        vendor_df = df[vendor_filter]
        
        vendor_party_name = vendor_df.groupby('$Key')['$LedgerName'].apply(lambda x: ', '.join(x.unique())).reset_index()
        vendor_party_name.rename(columns={'$LedgerName': 'ExpenseLedgerName'}, inplace=True)
        
        df = df.merge(vendor_party_name, on='$Key', how='left')
        
        return df
    
def SaleName(input_df):
        df = input_df.copy()  # Create a copy of the input dataframe to avoid modifying the original
        
        vendor_filter = df['$Led_Group'].isin(['Sales Accounts', 'Direct Incomes','Indirect Incomes'])
        vendor_df = df[vendor_filter]
        
        vendor_party_name = vendor_df.groupby('$Key')['$LedgerName'].apply(lambda x: ', '.join(x.unique())).reset_index()
        vendor_party_name.rename(columns={'$LedgerName': 'SaleLedgerName'}, inplace=True)
        
        df = df.merge(vendor_party_name, on='$Key', how='left')
        
        return df




pyodbc.drivers()
import pandas as pd
server = 'localhost,9000'
#database = 'database_name' # enter database name
cnxn = pyodbc.connect('DRIVER={Tally ODBC Driver64};SERVER='+server+';Trusted_Connection=yes;')
cursor = cnxn.cursor()
# select command
query = '''SELECT $Key, $MasterId, $AlterID, $VoucherNumber, $Date, $VoucherTypeName, $Led_Lineno, $Type, $LedgerName, $Amount, $Led_Parent, $Led_Group, $Party_LedName, $Vch_GSTIN, $Led_GSTIN, $Party_GST_Type, $GST_Classification, $Narration, $EnteredBy, $LastEventinVoucher, $UpdatedDate, $UpdatedTime, $Nature_Led, $Led_MID, $CompanyName, $Year_from, $Year_to, $Company_number, $Path FROM A__DayBook'''
query2 = '''SELECT $Name, $_PrimaryGroup, $Parent, $OpeningBalance, $ClosingBalance, $_PrevYearBalance, $IsRevenue, $PartyGSTIN, $MasterId, $AlterID, $Nature_Led, $UpdatedDate, $UpdatedTime, $CreatedBy, $CreatedDate, $AlteredDate, $AlteredBy, $LastVoucherDate, $CompanyName, $Year_from, $Year_to, $Company_number, $Path, $Amount, $Date, $VoucherNumber, $Type FROM A__M_Ledger'''
query3 = '''SELECT $Name,$OpeningBalance, $ClosingBalance FROM Ledger'''

data = pd.read_sql(query, cnxn)

# Retrieve the Company Name from the first row of the '$CompanyName' column
if not data.empty and '$CompanyName' in data.columns:
    company_name = data['$CompanyName'].iloc[0]
    # Clean the company name to be file-system safe
    company_name = "".join(x for x in company_name if x.isalnum() or x in "._- ")
else:
    company_name = "DefaultCompany"  # Default name if the column doesn't exist or dataframe is empty.

# Get the current timestamp
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

# Define output file names with timestamps
daybook_filename = f"daybook_{company_name}_{timestamp}.csv"
trial_filename = f"trial_{company_name}_{timestamp}.csv"
balances_filename = f"balances_{timestamp}.csv"

# Export data to CSV files with timestamps
data = pd.read_sql(query, cnxn)
data = VendorPartyName(data)
data = ExpenseName(data)
data = SaleName(data)
data.head()
data.to_csv(daybook_filename)

data2 = pd.read_sql(query2,cnxn)
data2.head()
data2.to_csv(trial_filename)

data3 = pd.read_sql(query3,cnxn)
data3.head()
data3.to_csv(balances_filename)

print(f"Data exported to: {daybook_filename}, {trial_filename}, {balances_filename}")




