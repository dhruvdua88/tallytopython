#requires two TDLS


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

pyodbc.drivers()
import pandas as pd
server = 'localhost,9000'
#database = 'database_name' # enter database name
cnxn = pyodbc.connect('DRIVER={Tally ODBC Driver64};SERVER='+server+';Trusted_Connection=yes;')
cursor = cnxn.cursor()
# select command
query = '''SELECT $Key, $MasterId, $AlterID, $VoucherNumber, $Date, $VoucherTypeName, $Led_Lineno, $Type, $LedgerName, $Amount, $Led_Parent, $Led_Group, $Party_LedName, $Vch_GSTIN, $Led_GSTIN, $Party_GST_Type, $GST_Classification, $Narration, $EnteredBy, $LastEventinVoucher, $UpdatedDate, $UpdatedTime, $Nature_Led, $Led_MID, $CompanyName, $Year_from, $Year_to, $Company_number, $Path FROM A__DayBook'''
query2 = '''SELECT $Name, $_PrimaryGroup, $Parent, $OpeningBalance, $ClosingBalance, $_PrevYearBalance, $IsRevenue, $PartyGSTIN, $MasterId, $AlterID, $Nature_Led, $UpdatedDate, $UpdatedTime, $CreatedBy, $CreatedDate, $AlteredDate, $AlteredBy, $LastVoucherDate, $CompanyName, $Year_from, $Year_to, $Company_number, $Path, $Amount, $Date, $VoucherNumber, $Type FROM A__M_Ledger'''
query3 = '''SELECT * FROM TLedger'''

data = pd.read_sql(query, cnxn)
data.head()
data.to_csv('daybook.csv')

data2 = pd.read_sql(query2,cnxn)
data2.head()
data2.to_csv('trial.csv')

data3 = pd.read_sql(query2,cnxn)
data3.head()
data3.to_csv('balances.csv')

