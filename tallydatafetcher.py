
import requests
import tkinter as tk
from tkinter import filedialog
import xml.etree.ElementTree as ET
import pandas as pd
from lxml import etree
import pyodbc
from datetime import datetime
import os
import requests
from bs4 import BeautifulSoup
import xlsxwriter





import tkinter as tk
from tkinter import messagebox
import pandas as pd

class OptionSelector:
    def __init__(self, dataframe):
        self.df = dataframe
        self.TDSselection = []

    def launch_selection_gui(self):
        self.root = tk.Tk()
        self.root.title("Select Options")

        self.search_var = tk.StringVar()
        self.search_var.trace("w", self.update_list)

        search_entry = tk.Entry(self.root, textvariable=self.search_var)
        search_entry.pack(pady=10)

        self.listbox = tk.Listbox(self.root, selectmode=tk.MULTIPLE, width=50, height=10)
        self.listbox.pack(pady=10)

        select_btn = tk.Button(self.root, text="Select", command=self.on_select)
        select_btn.pack(pady=10)

        self.update_list()  # Initially populate the listbox
        self.root.mainloop()

    def on_select(self):
        selections = self.listbox.curselection()
        self.TDSselection = [self.listbox.get(i) for i in selections]
        messagebox.showinfo("Selections for TDS and GST (first time TDS will come and then GST)", "\n".join(self.TDSselection))
        self.root.destroy()

    def update_list(self, *args):
        search_term = self.search_var.get().lower()
        self.listbox.delete(0, tk.END)
        # Get unique values from the LEDGERNAME column
        unique_options = self.df['LEDGERNAME'].unique()
        for option in unique_options:
            if search_term in option.lower():
                self.listbox.insert(tk.END, option)


def select_tds_options(dataframe):
    selector = OptionSelector(dataframe)
    selector.launch_selection_gui()
    return selector.TDSselection

def select_gst_options(dataframe):
    selector = OptionSelector(dataframe)
    selector.launch_selection_gui()
    return selector.TDSselection




class App:
    def __init__(self, root):
        self.root = root
        self.data_frame = None

        self.button_fetch_voucher_data = tk.Button(root, text='Fetch Voucher Data', command=self.fetch_voucher_data)
        self.button_fetch_trial_data = tk.Button(root, text='Fetch Ledger Data', command=self.fetch_ledger_data)
        self.button_fetch_voucher_data.pack()
        self.button_fetch_trial_data.pack()

    def get_save_path(self):
        root = tk.Tk()
        root.withdraw()
        file_path = filedialog.asksaveasfilename(defaultextension=".xml", filetypes=[("XML files", "*.xml"), ("All files", "*.*")])
        return file_path

    @staticmethod
    def fetch_voucher_data():
        data_frame = App.fetch_data_from_tally()
        if data_frame is not None:
            print("Data successfully fetched.")
        else:
            print("Failed to fetch data.")
        return data_frame

    @staticmethod
    def fetch_ledger_data():
        data_frame = App.fetch_ledger_from_tally()
        if data_frame is not None:
            print("Data successfully fetched.")
        else:
            print("Failed to fetch data.")
        return data_frame
    
    @staticmethod
    def fetch_stock_data():
        data_frame = App.fetch_stock_data_from_tally()
        if data_frame is not None:
            print("Data successfully fetched.")
        else:
            print("Failed to fetch data.")
        return data_frame
    
    
   
    
    @staticmethod
    def fetch_ledger_from_tally():
    
        conn = pyodbc.connect('DRIVER={Tally ODBC Driver64};'
                              'SERVER=(local);'  # replace with your server name
                              'PORT=9000')  

        sql_query = 'Select $Name, $Parent,$OpeningBalance,$ClosingBalance,$_PrimaryGroup from Ledger'
        ledger_data=pd.read_sql(sql_query, conn)
        conn.close()
        
        return pd.DataFrame(ledger_data)
    
    
    
    
    @staticmethod
    def fetch_data_from_tally():
        TALLY_URL = "http://localhost:9000"
        xml_request = '''
        <ENVELOPE>
            <HEADER>
                <VERSION>1</VERSION>
                <TALLYREQUEST>Export</TALLYREQUEST>
                <TYPE>Collection</TYPE>
                <ID>LVoucher</ID>
            </HEADER>
            <BODY>
                <DESC>
                    <STATICVARIABLES>
                        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
                    </STATICVARIABLES>
                </DESC>
            </BODY>
        </ENVELOPE>
        '''

        
        
                # Send the POST request and handle potential errors
        response = requests.post(TALLY_URL, data=xml_request, headers={'Content-Type': 'application/xml'})
        response.raise_for_status()  # Raise an exception for non-200 status codes

        # Parse the XML response using BeautifulSoup
        soup = BeautifulSoup(response.content, 'xml')

        # Extract ledger entries using BeautifulSoup's find_all method
        ledger_entries = []
        for ledger_entry in soup.find_all('LEDGERENTRY'):  # Use find_all to get multiple entries
            entry_data = {}
            for child in ledger_entry.children:
                entry_data[child.name] = child.text.strip()  # Use name for tag and strip whitespace
            ledger_entries.append(entry_data)
        
        data_frame = pd.DataFrame(ledger_entries)

        data_frame['DATE'] = pd.to_datetime(data_frame['DATE'], format='%Y%m%d').dt.strftime('%d/%m/%Y')
        data_frame['AMOUNT'] = pd.to_numeric(data_frame['AMOUNT'])
        data_frame['MONTH'] = pd.to_datetime(data_frame['DATE'], format='%d/%m/%Y').dt.strftime('%B')
        data_frame['YEAR'] = pd.to_datetime(data_frame['DATE'], format='%d/%m/%Y').dt.year

        return data_frame

    @staticmethod
    def fetch_stock_data_from_tally():
        TALLY_URL = "http://localhost:9000"
        xml_request = '''
        <ENVELOPE>
            <HEADER>
                <VERSION>1</VERSION>
                <TALLYREQUEST>Export</TALLYREQUEST>
                <TYPE>Collection</TYPE>
                <ID>SVoucher</ID>
            </HEADER>
            <BODY>
                <DESC>
                    <STATICVARIABLES>
                        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
                    </STATICVARIABLES>
                </DESC>
            </BODY>
        </ENVELOPE>
        '''

        response = requests.post(TALLY_URL, data=xml_request, headers={'Content-Type': 'application/xml'})
        if response.status_code != 200:
            return None

        root = ET.fromstring(response.content)
        ledger_entries = []
        for ledger in root.findall('.//LEDGERENTRY'):
            entry_data = {}
            for child in ledger:
                entry_data[child.tag] = child.text
            ledger_entries.append(entry_data)
        
        data_frame = pd.DataFrame(ledger_entries)
        #data_frame['DATE'] = pd.to_datetime(data_frame['DATE'], format='%Y%m%d').dt.strftime('%d/%m/%Y')
        #data_frame['AMOUNT'] = pd.to_numeric(data_frame['AMOUNT'])
        #data_frame['MONTH'] = pd.to_datetime(data_frame['DATE'], format='%d/%m/%Y').dt.strftime('%B')
        #data_frame['YEAR'] = pd.to_datetime(data_frame['DATE'], format='%d/%m/%Y').dt.year

        return data_frame
    
class DataAnalysis:
    def VendorPartyName(input_df):
        df = input_df.copy()  # Create a copy of the input dataframe to avoid modifying the original
        
        vendor_filter = df['Primary'].isin(['Sundry Debtors', 'Sundry Creditors'])
        vendor_df = df[vendor_filter]
        
        vendor_party_name = vendor_df.groupby('GUID')['LEDGERNAME'].apply(lambda x: ', '.join(x.unique())).reset_index()
        vendor_party_name.rename(columns={'LEDGERNAME': 'VendorPartyName'}, inplace=True)
        
        df = df.merge(vendor_party_name, on='GUID', how='left')
        
        return df
    
    def ExpenseName(input_df):
        df = input_df.copy()  # Create a copy of the input dataframe to avoid modifying the original
        
        vendor_filter = df['Primary'].isin(['Direct Expenses', 'Purchases','Indirect Expenses','Fixed Assets'])
        vendor_df = df[vendor_filter]
        
        vendor_party_name = vendor_df.groupby('GUID')['LEDGERNAME'].apply(lambda x: ', '.join(x.unique())).reset_index()
        vendor_party_name.rename(columns={'LEDGERNAME': 'ExpenseLedgerName'}, inplace=True)
        
        df = df.merge(vendor_party_name, on='GUID', how='left')
        
        return df
    
    def SaleName(input_df):
        df = input_df.copy()  # Create a copy of the input dataframe to avoid modifying the original
        
        vendor_filter = df['Primary'].isin(['Sales Accounts', 'Direct Incomes','Indirect Incomes'])
        vendor_df = df[vendor_filter]
        
        vendor_party_name = vendor_df.groupby('GUID')['LEDGERNAME'].apply(lambda x: ', '.join(x.unique())).reset_index()
        vendor_party_name.rename(columns={'LEDGERNAME': 'SaleLedgerName'}, inplace=True)
        
        df = df.merge(vendor_party_name, on='GUID', how='left')
        
        return df
        
    def TDSandGSTLedgerNames(input_df):
        df = input_df.copy()  # Create a copy of the input dataframe to avoid modifying the original
        
        # Additional step for 'TDS Ledger Name'
        tdsfilterlist = select_tds_options(df)
        TDS_filter = df['LEDGERNAME'].isin(tdsfilterlist)
        TDS_df = df[TDS_filter]
        TDS_ledgers_name = TDS_df.groupby('GUID')['LEDGERNAME'].apply(lambda x: ', '.join(x.unique())).reset_index()
        TDS_ledgers_name.rename(columns={'LEDGERNAME': 'TDSLedgersName'}, inplace=True)
        df = df.merge(TDS_ledgers_name, on='GUID', how='left')

        # Additional step for 'GST Ledger Name'
        gstfilterlist = select_gst_options(df)
        gst_types = df['LEDGERNAME'].isin(gstfilterlist)
        GST_df = df[gst_types]
        GST_ledgers_name = GST_df.groupby('GUID')['LEDGERNAME'].apply(lambda x: ', '.join(x.unique())).reset_index()
        GST_ledgers_name.rename(columns={'LEDGERNAME': 'GSTLedgersName'}, inplace=True)
        df = df.merge(GST_ledgers_name, on='GUID', how='left')

        # Group by 'GUID' and 'Primary' and calculate the sum of 'Amount'
        grouped = df.groupby(['GUID', 'Primary'])['AMOUNT'].sum().unstack(fill_value=0).reset_index()
        df = pd.merge(df, grouped, on='GUID', how='left')

        def calculate_tds_amount(row):
            if pd.notnull(row['LEDGERNAME']) and row['LEDGERNAME'] in tdsfilterlist:
                return row['AMOUNT']
            else:
                return 0
            
        def calculate_GST_amount(row):
            if pd.notnull(row['LEDGERNAME']) and row['LEDGERNAME'] in gstfilterlist:
                return row['AMOUNT']
            else:
                return 0
        
                
        df['TDS Amount'] = df.apply(calculate_tds_amount, axis=1)
        df['TDS Amount Sum'] = df.groupby('GUID')['TDS Amount'].transform('sum')
        df.drop('TDS Amount', axis=1, inplace=True)

        df['GST Amount'] = df.apply(calculate_GST_amount, axis=1)
        df['GST Amount Sum'] = df.groupby('GUID')['GST Amount'].transform('sum')
        df.drop('GST Amount', axis=1, inplace=True)
              
        return df
    
    def clean_dataframe(df):
    # Define the list of characters to remove
    # Add any additional characters to this list as needed
        characters_to_remove = ['♦']

        # Function to remove the characters from a string, ignoring non-string values
        def remove_characters(value):
            if isinstance(value, str):
                for char in characters_to_remove:
                    value = value.replace(char, 'Primary')  # Replace with empty string if value is a string
            return value

        # Apply the function to clean only columns detected as containing string data
        for column in df.columns:
            if df[column].dtype == object:  # Checks if column is of object type, typically strings
                df[column] = df[column].apply(remove_characters)
        
        return df
            







#root = tk.Tk()
#app = App(root)
#root.mainloop()

trial_data = App.fetch_ledger_data()
trial_data.rename(columns={'$Name': 'LEDGERNAME'}, inplace=True)
trial_data.rename(columns={'$Parent': 'Parent'}, inplace=True)
trial_data.rename(columns={'$_PrimaryGroup': 'Primary'}, inplace=True)
trial_data = DataAnalysis.clean_dataframe(trial_data)


voucher_data = App.fetch_voucher_data()
#voucher_data = DataAnalysis.clean_dataframe(voucher_data)
merged_data = pd.merge(voucher_data, trial_data, on='LEDGERNAME', how='left')
voucher_data = merged_data

voucher_data = DataAnalysis.VendorPartyName(voucher_data)
voucher_data = DataAnalysis.ExpenseName(voucher_data)
voucher_data = DataAnalysis.SaleName(voucher_data)
voucher_data = DataAnalysis.TDSandGSTLedgerNames(voucher_data)




#stock_data = App.fetch_stock_data()
# Get the current date and time in the format ddmmyyhhss
timestamp = datetime.now().strftime("%d%m%y%H%M%S")





folder_path = os.getcwd()


# Export voucher_data, trial_data, and stock_data to CSV files in the current directory
voucher_data=DataAnalysis.clean_dataframe(voucher_data)
voucher_data.to_excel(os.path.join(folder_path, f'voucher_data_{timestamp}.xlsx'), index=False)
trial_data.to_excel(os.path.join(folder_path, f'trial_data_{timestamp}.xlsx'), index=False)
#stock_data.to_excel(os.path.join(folder_path, f'stock_data_{timestamp}.xlsx'), index=False)



