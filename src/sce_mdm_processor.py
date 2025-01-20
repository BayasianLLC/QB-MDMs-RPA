import re
import html
import pandas as pd
import os
from datetime import datetime
import requests
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
import time
from quickbase_client.orm.table import QuickbaseTable
from quickbase_client.orm.app import QuickbaseApp
from quickbase_client import QuickbaseTableClient
from io import BytesIO


def get_sharepoint_context():
   
   #SharePoint credentials and site URL
   # sharepoint_url = "https://wescodist.sharepoint.com/sites/UtilityMDMs-SCE"
   # username = "juan.bayas@wescodist.com"
   # password = "DhkofiL@512345"
   
   sharepoint_url = "https://stdntpartners.sharepoint.com/sites/MDMQB"
   username = "Victor.Sabare@studentambassadors.com"
   password = "ni2b:+AANpP?N7w"


   try:
    auth_context = AuthenticationContext(sharepoint_url)
    auth_context.acquire_token_for_user(username, password)
    ctx = ClientContext(sharepoint_url, auth_context)
    return ctx
   
   except Exception as e:
    print(f"Error connecting to SharePoint: {str(e)}")
    return None

def check_new_files(ctx, last_check_time):
    try:
        # Get the web's server relative URL first
        ctx.load(ctx.web)
        ctx.execute_query()
        web_url = ctx.web.properties['ServerRelativeUrl']
        
        # Construct the full folder path
        # folder_path = f"{web_url}/Shared%20Documents/SCE/MDM%20Files"
        folder_path = f"{web_url}/Shared%20Documents"
        
        print(f"Accessing folder: {folder_path}")
        
        # Get files from SharePoint folder
        folder = ctx.web.get_folder_by_server_relative_url(folder_path)
        files = folder.files
        ctx.load(files)
        files.execute_query()
        
        # Look for new XLSB or XLSM files
        new_files = [f for f in files 
                    if "SCE WCDM" in f.properties["Name"] 
                    and (f.properties["Name"].lower().endswith('.xlsb') 
                         or f.properties["Name"].lower().endswith('.xlsm'))]
        
        print(f"Found {len(new_files)} new files")
        return new_files
    except Exception as e:
        print(f"Error checking SharePoint: {str(e)}")
        return []


def transform_mdm_file(file_content, output_file):
    try:
        print("Starting file transformation...")
        excel_data = BytesIO(file_content)
        
        # Detect file type and use appropriate engine
        if output_file.lower().endswith('.xlsb'):
            df = pd.read_excel(excel_data, engine='pyxlsb', header=None)
        elif output_file.lower().endswith('.xlsm'):
            df = pd.read_excel(excel_data, engine='openpyxl', header=None)
        else:
            raise ValueError("Unsupported file format. Only .xlsb and .xlsm files are supported.")

        print("\nOriginal columns:")
        for i, col in enumerate(df.columns):
            print(f"Column {i}: {col}")

        # Create mapping based on QuickBase field IDs
        column_mapping = {
            0: 'MDM Sort',              # Field ID: 6
            1: 'Added By',              # Field ID: 7
            2: 'Date Added',            # Field ID: 8
            3: 'In Scope',              # Field ID: 9
            4: 'Servicing Business Unit', # Field ID: 10
            5: 'Pricing Category / Owner', # Field ID: 11
            6: 'Product Category',       # Field ID: 12
            7: 'Product Sub-Category',   # Field ID: 13
            8: 'Cust. ID #',            # Field ID: 14
            9: 'Main Category',          # Field ID: 15
            10: 'Long Description',      # Field ID: 16
            11: 'UOP',                  # Field ID: 17
            12: 'Last 12 Usage',         # Field ID: 18
            13: 'Annual Times Purchased', # Field ID: 19
            14: 'Manufacturer',          # Field ID: 20
            15: 'Manufacturer Part #',   # Field ID: 21
            16: 'Manufacturer Status',   # Field ID: 22
            17: 'Customer Info Change Date', # Field ID: 23
            18: 'Customer Reqd Status',  # Field ID: 24
            19: 'VMI (Y/N)',            # Field ID: 25
            20: 'Customer Comments',     # Field ID: 26
            21: 'Sugg. Sell Price',     # Field ID: 27
            22: 'Sugg. Sell Price Extended', # Field ID: 28
            23: 'Markup',               # Field ID: 29
            24: 'Billing Margin %',      # Field ID: 30
            25: 'Extended Billing Margin $', # Field ID: 31
            26: 'Item Review Notes',     # Field ID: 32
            27: 'Vendor Name',          # Field ID: 33
            28: 'Vendor Code',          # Field ID: 34
            29: 'Blanket #',            # Field ID: 35
            30: 'Blanket Load Price',    # Field ID: 36
            31: 'Blanket Load Standard Pack', # Field ID: 37
            32: 'Blanket Load Leadtime', # Field ID: 38
            33: 'Blanket Load Date',     # Field ID: 39
            34: 'Source',               # Field ID: 40
            35: 'Source Manufacturer',   # Field ID: 41
            36: 'Source Supplier #',     # Field ID: 42
            37: 'SIM',                  # Field ID: 43
            38: 'SIM MFR',              # Field ID: 44
            39: 'SIM Item',             # Field ID: 45
            40: 'Wesnet Catalog #',      # Field ID: 46
            41: 'Wesnet SIM Description', # Field ID: 47
            42: 'Wesnet UOM',           # Field ID: 48
            43: 'Source Count',         # Field ID: 49
            44: 'Rank',                 # Field ID: 50
            45: 'Low Cost',             # Field ID: 51
            46: 'Cost Source',          # Field ID: 52
            47: 'Cost Extended',        # Field ID: 53
            48: 'UOP Multiplier Factor', # Field ID: 54
            49: 'UOP Divider Factor',   # Field ID: 55
            50: 'Spa Cost',             # Field ID: 56
            51: 'Spa Into Stock Cost',   # Field ID: 57
            52: 'Spa Number',           # Field ID: 58
            53: 'Spa Start Date',       # Field ID: 59
            54: 'Spa End Date',         # Field ID: 60
            55: 'DC Xfer',              # Field ID: 61
            56: '8500 Low Repl Cost',    # Field ID: 62
            57: '8500 Low Repl Cost Extended', # Field ID: 63
            58: '8570 Low Repl Cost',    # Field ID: 64
            59: '8570 Low Repl Cost Extended', # Field ID: 65
            60: 'Future Quote Loaded',   # Field ID: 66
            61: 'Last Date Quote Modified', # Field ID: 67
            62: 'Quoted Mfr / Brand',    # Field ID: 68
            63: 'Quoted Mfr Part Number', # Field ID: 69
            64: 'Direct Equal',         # Field ID: 70
            65: 'Returnable',           # Field ID: 71
            66: 'Supplier Comments',     # Field ID: 72
            67: 'Quoted Price',         # Field ID: 73
            68: 'List Price',           # Field ID: 74
            69: 'Unit of Measure',       # Field ID: 75
            70: 'Qty per Unit of Measure', # Field ID: 76
            71: 'Std Purchase Qty',      # Field ID: 77
            72: 'Lead Time (Calendar Days)', # Field ID: 78
            73: 'Quote #',              # Field ID: 79
            74: 'Quote End Date',       # Field ID: 80
            75: 'Minimum Order',        # Field ID: 81
            76: 'Freight Terms',        # Field ID: 82
            77: 'Quote - Contact / Preparer Name', # Field ID: 83
            78: 'Quote - Contact Phone', # Field ID: 84
            79: 'Quote - Contact E-mail', # Field ID: 85
            80: 'Purchasing - Contact Name', # Field ID: 86
            81: 'Purchasing - Contact Phone', # Field ID: 87
            82: 'Purchasing - Contact E-mail', # Field ID: 88
            83: 'Last 12',              # Field ID: 89
            84: 'VC',                   # Field ID: 90
            85: 'CC',                   # Field ID: 91
            86: 'Loaded ORP',           # Field ID: 92
            87: 'Loaded EOQ',           # Field ID: 93
            88: 'On Hand',              # Field ID: 94
            89: 'On Order',             # Field ID: 95
            90: 'On Backorder',         # Field ID: 96
            91: 'Net Stock',            # Field ID: 97
            92: 'Region Low Repl Cost',  # Field ID: 131
            93: 'Region Low Repl Cost Extended', # Field ID: 132
            94: 'Tier Cost',            # Field ID: 133
            95: 'UOM',                  # Field ID: 134
            96: 'Standard Pack',        # Field ID: 135
            97: 'Leadtime',            # Field ID: 136
            98: 'List Price',          # Field ID: 137
            99: 'Quote Start Date',     # Field ID: 138
            100: 'Purchasing - Contact Phone' # Field ID: 139
        }

        # Rename columns
        df = df.rename(columns=column_mapping)

        # Keep only the mapped columns
        df = df[list(column_mapping.values())]
        
        # Skip the first two rows (headers) and reset index
        df = df.iloc[2:].reset_index(drop=True)

        # Define column types for proper formatting
        date_columns = ['Date Added', 'Customer Info Change Date', 'Blanket Load Date', 
                       'Spa Start Date', 'Spa End Date', 'Last Date Quote Modified',
                       'Quote End Date', 'Quote Start Date']
        
        numeric_columns = ['MDM Sort', 'Last 12 Usage', 'Annual Times Purchased', 
                         'Source Count', 'Rank', 'UOP Multiplier Factor', 
                         'UOP Divider Factor', 'Qty per Unit of Measure', 
                         'Std Purchase Qty', 'Lead Time (Calendar Days)', 
                         'Last 12', 'VC', 'Loaded ORP', 'Loaded EOQ', 
                         'On Hand', 'On Order', 'On Backorder', 'Net Stock']
        
        currency_columns = ['Sugg. Sell Price', 'Sugg. Sell Price Extended', 'Markup',
                          'Extended Billing Margin $', 'Blanket Load Price', 
                          'Low Cost', 'Cost Extended', 'Spa Cost', 
                          'Spa Into Stock Cost', '8500 Low Repl Cost',
                          '8500 Low Repl Cost Extended', '8570 Low Repl Cost',
                          '8570 Low Repl Cost Extended', 'Quoted Price', 
                          'List Price', 'Region Low Repl Cost',
                          'Region Low Repl Cost Extended', 'Tier Cost']
        
        checkbox_columns = ['In Scope', 'VMI (Y/N)', 'Future Quote Loaded', 
                          'Direct Equal', 'Returnable']

        # Apply data type formatting
        for col in df.columns:
            if col in date_columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')
            elif col in numeric_columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')
            elif col in currency_columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')
            elif col in checkbox_columns:
                df[col] = df[col].map({'Yes': 'Y', 'No': 'N', 'TRUE': 'Y', 'FALSE': 'N', 
                                      'Y': 'Y', 'N': 'N', True: 'Y', False: 'N'})

        print(f"Saving processed file to: {output_file}")
        df.to_csv(output_file, index=False)
        
        print("\nSample of data being saved:")
        print(df.head())
        
        if upload_to_quickbase(output_file):
            print("File successfully uploaded to QuickBase")
            return True
        else:
            print("Failed to upload to QuickBase")
            return False
        
    except Exception as e:
        print(f"Error processing file: {str(e)}")
        print("Full error traceback:")
        import traceback
        print(traceback.format_exc())
        return False
    
def delete_quickbase_records():
    try:
        print("Deleting existing QuickBase records...")
        
        headers = {
            'QB-Realm-Hostname': 'wesco.quickbase.com',
            'Authorization': 'QB-USER-TOKEN cacrrx_vcs_0_ezvd3icw7ds8wdegdjbwbigxm45',
            'Content-Type': 'application/xml'
        }
        
        # Create XML request to delete all records where MDM Sort > 0
        xml_request = """<?xml version="1.0" ?>
        <qdbapi>
            <apptoken>None</apptoken>
            <query>{6.GT.0}</query>
        </qdbapi>"""
        
        # Send request to purge records
        api_url = 'https://wesco.quickbase.com/db/bfdix6cda'
        
        response = requests.post(
            f"{api_url}?a=API_PurgeRecords",
            headers=headers,
            data=xml_request.encode('utf-8'),
            verify=False
        )
        
        if response.status_code in [200, 201]:
            # Check for successful deletion
            if '<errcode>0</errcode>' in response.text:
                # Extract number of deleted records
                match = re.search(r'<num_records_deleted>(\d+)</num_records_deleted>', response.text)
                if match:
                    num_deleted = match.group(1)
                    print(f"Successfully deleted {num_deleted} records")
                else:
                    print("Successfully deleted records (count unknown)")
                print("Response:", response.text)
                return True
            else:
                print(f"Failed to delete records. Error in response: {response.text}")
                return False
        else:
            print(f"Failed to delete records. Status code: {response.status_code}")
            print("Error response:", response.text)
            return False
            
    except Exception as e:
        print(f"Error deleting records: {str(e)}")
        print("Full error traceback:")
        import traceback
        print(traceback.format_exc())
        return False


def clean_xml_string(value):
    if pd.isna(value):
        return ''
    # Convert to string and escape special characters
    value = str(value)
    # Remove angle brackets < >
    value = re.sub(r'[<>]', '', value)
    # Escape ampersands and quotes
    value = html.escape(value)
    return value

def create_record_xml(row):
    # Clean each field value
    record_xml = """<?xml version="1.0" ?>
    <qdbapi>"""
    
    # Add each field with cleaned value
    for field_name, value in row.items():
        clean_value = clean_xml_string(value)
        record_xml += f'\n        <field name="{field_name}">{clean_value}</field>'
    
    record_xml += "\n    </qdbapi>"
    return record_xml
    
def clean_xml_string(value):
    if pd.isna(value):
        return ''
    value = str(value)
    value = re.sub(r'[<>]', '', value)
    value = html.escape(value)
    return value

def create_record_xml(row):
    record_xml = """<?xml version="1.0" ?>
    <qdbapi>"""
    for field_name, value in row.items():
        clean_value = clean_xml_string(value)
        record_xml += f'\n        <field name="{field_name}">{clean_value}</field>'
    record_xml += "\n    </qdbapi>"
    return record_xml

def upload_to_quickbase(csv_file, batch_size=1000):
    try:
        print("Starting QuickBase update process...")
        
        # Updated field mapping based on the QuickBase field IDs from images
        field_mapping = {
            'MDM Sort': 6,
            'Added By': 7,
            'Date Added': 8,
            'In Scope': 9,
            'Servicing Business Unit': 10,
            'Pricing Category / Owner': 11,
            'Product Category': 12,
            'Product Sub-Category': 13,
            'Cust. ID #': 14,
            'Main Category': 15,
            'Long Description': 16,
            'UOP': 17,
            'Last 12': 18,
            'Annual Times Purchased': 19,
            'Manufacturer': 20,
            'Manufacturer Part #': 21,
            'Manufacturer Status': 22,
            'Customer Info Change Date': 23,
            'Customer Reqd Status': 24,
            'VMI (Y/N)': 25,
            'Customer Comments': 26,
            'Sugg. Sell Price': 27,
            'Sugg. Sell Price Extended': 28,
            'Markup': 29,
            'Billing Margin %': 30,
            'Extended Billing Margin $': 31,
            'Item Review Notes': 32,
            'Vendor Name': 33,
            'Vendor Code': 34,
            'Blanket #': 35,
            'Blanket Load Price': 36,
            'Blanket Load Standard Pack': 37,
            'Blanket Load Leadtime': 38,
            'Blanket Load Date': 39,
            'Source': 40,
            'Source Manufacturer': 41,
            'Source Supplier #': 42,
            'SIM': 43,
            'SIM MFR': 44,
            'SIM Item': 45,
            'Wesnet Catalog #': 46,
            'Wesnet SIM Description': 47,
            'Wesnet UOM': 48,
            'Source Count': 49,
            'Rank': 50,
            'Low Cost': 51,
            'Cost Source': 52,
            'Cost Extended': 53,
            'UOP Multiplier Factor': 54,
            'UOP Divider Factor': 55,
            'Spa Cost': 56,
            'Spa Into Stock Cost': 57,
            'Spa Number': 58,
            'Spa Start Date': 59,
            'Spa End Date': 60,
            'DC Xfer': 61,
            '8500 Repl Cost': 62,
            '8500 Repl Cost Extended': 63,
            '8570 Repl Cost': 64,
            '8570 Repl Cost Extended': 65,
            'Future Quote Loaded': 66,
            'Last Date Quote Modified': 67,
            'Quoted Mfr / Brand': 68,
            'Quoted Mfr Part Number': 69,
            'Direct Equal': 70,
            'Returnable': 71,
            'Supplier Comments': 72,
            'Quoted Price': 73,
            'List Price': 74,
            'Unit of Measure': 75,
            'Qty per Unit of Measure': 76,
            'Std Purchase Qty': 77,
            'Lead Time (Calendar Days)': 78,
            'Quote #': 79,
            'Quote End Date': 80,
            'Minimum Order': 81,
            'Freight Terms': 82,
            'Quote - Contact / Preparer Name': 83,
            'Quote - Contact Phone': 84,
            'Quote - Contact E-mail': 85,
            'Purchasing - Contact Name': 86,
            'Purchasing - Contact Email': 87,
            'Purchasing - Contact Fax': 88,
            'Last 12': 89,
            'VC': 90,
            'CC': 91,
            'Loaded ORP': 92,
            'Loaded EOQ': 93,
            'On Hand': 94,
            'On Order': 95,
            'On Backorder': 96,
            'Net Stock': 97,
            'WESCO Stocking Item': 98,
            'WESCO Linked Cust ID': 99,
            'Combined Last 12 Purchases': 100,
            'Combined Last 12 Count': 101,
            'MMP Rank': 102,
            'SIM (Y/N)': 103,
            'Supplier Number (Y/N)': 104,
            'Cost (Y/N)': 105,
            'Ready to load (Y/N)': 106,
            'ORP': 107,
            'EOQ': 108,
            'Inventory Max Value': 109,
            'Quote Start Date': 110,
            'Region Low Repl Cost': 131,
            'Region Low Repl Cost Extended': 132,
            'Tier Cost': 133,
            'UOM': 134,
            'Standard Pack': 135,
            'Leadtime': 136,
            'List Price': 137,
            'Quote Start Date': 138,
            'Purchasing - Contact Phone': 139
        }

        # Read CSV into DataFrame
        df = pd.read_csv(csv_file, dtype=str)
        total_records = len(df)
        print(f"Read {total_records} records from CSV")

        # Define column groups based on data types from images
        date_columns = [
            'Date Added', 'Customer Info Change Date', 'Blanket Load Date',
            'Spa Start Date', 'Spa End Date', 'Last Date Quote Modified',
            'Quote End Date', 'Quote Start Date'
        ]
        
        numeric_columns = [
            'MDM Sort', 'Source Count', 'Rank', 'UOP Multiplier Factor',
            'UOP Divider Factor', 'Qty per Unit of Measure', 'Std Purchase Qty',
            'Lead Time (Calendar Days)', 'Last 12', 'VC', 'Loaded ORP',
            'Loaded EOQ', 'On Hand', 'On Order', 'On Backorder', 'Net Stock',
            'Standard Pack', 'Leadtime'
        ]
        
        currency_columns = [
            'Sugg. Sell Price', 'Sugg. Sell Price Extended', 'Markup',
            'Extended Billing Margin $', 'Blanket Load Price', 'Low Cost',
            'Cost Extended', 'Spa Cost', 'Spa Into Stock Cost',
            '8500 Repl Cost', '8500 Repl Cost Extended',
            '8570 Repl Cost', '8570 Repl Cost Extended',
            'Quoted Price', 'List Price', 'Region Low Repl Cost',
            'Region Low Repl Cost Extended', 'Tier Cost',
            'Inventory Max Value'
        ]
        
        percent_columns = ['Billing Margin %']
        
        checkbox_columns = [
            'In Scope', 'Future Quote Loaded', 'Direct Equal',
            'Returnable', 'VMI (Y/N)', 'SIM (Y/N)',
            'Supplier Number (Y/N)', 'Cost (Y/N)',
            'Ready to load (Y/N)', 'ORP', 'EOQ'
        ]
        
        phone_columns = [
            'Quote - Contact Phone', 'Purchasing - Contact Phone',
            'Purchasing - Contact Fax'
        ]
        
        email_columns = [
            'Quote - Contact E-mail', 'Purchasing - Contact Email'
        ]

        def format_date(date_str):
            if pd.isna(date_str) or date_str == '':
                return ''
            try:
                if str(date_str).replace('.', '').isdigit():
                    date_val = float(date_str)
                    date_obj = pd.Timestamp('1899-12-30') + pd.Timedelta(days=date_val)
                    return date_obj.strftime('%Y-%m-%d')
                return date_str
            except:
                return date_str

        def format_number(val):
            if pd.isna(val) or val == '':
                return ''
            try:
                num = float(val)
                if num.is_integer():
                    return str(int(num))
                return f"{num:.2f}"
            except:
                return val

        def format_currency(val):
            if pd.isna(val) or val == '':
                return ''
            try:
                num = float(val)
                return f"{num:.2f}"
            except:
                return val

        def format_percent(val):
            if pd.isna(val) or val == '':
                return ''
            try:
                num = float(val)
                return f"{num:.2f}"
            except:
                return val

        def format_checkbox(val):
            if pd.isna(val) or val == '':
                return ''
            val = str(val).upper()
            if val in ['Y', 'YES', 'TRUE', '1']:
                return 'Y'
            if val in ['N', 'NO', 'FALSE', '0']:
                return 'N'
            return val

        def format_phone(phone_str):
            if pd.isna(phone_str) or phone_str == '':
                return ''
            phone = re.sub(r'[^\d]', '', str(phone_str))
            if len(phone) == 10:
                return f"({phone[:3]}) {phone[3:6]}-{phone[6:]}"
            return phone_str

        def format_email(email_str):
            if pd.isna(email_str) or email_str == '':
                return ''
            return str(email_str).strip().lower()

        def clean_text(val):
            if pd.isna(val):
                return ''
            val = str(val)
            val = val.replace('&', '&amp;')
            val = val.replace('<', '&lt;')
            val = val.replace('>', '&gt;')
            val = val.replace('"', '&quot;')
            val = val.replace("'", '&apos;')
            val = ''.join(char for char in val if ord(char) >= 32 or char in '\n\r\t')
            return val.strip()

        # Clean data
        df = df.replace({pd.NA: '', 'nan': '', 'NaN': '', None: ''})
        df = df.fillna('')

        # Apply formatting based on column groups
        for col in df.columns:
            if col in date_columns:
                df[col] = df[col].apply(format_date)
            elif col in numeric_columns:
                df[col] = df[col].apply(format_number)
            elif col in currency_columns:
                df[col] = df[col].apply(format_currency)
            elif col in percent_columns:
                df[col] = df[col].apply(format_percent)
            elif col in checkbox_columns:
                df[col] = df[col].apply(format_checkbox)
            elif col in phone_columns:
                df[col] = df[col].apply(format_phone)
            elif col in email_columns:
                df[col] = df[col].apply(format_email)
            else:
                df[col] = df[col].apply(clean_text)

        # Keep only first occurrence of each MDM Sort value
        print("\nRemoving duplicate MDM Sort values...")
        original_count = len(df)
        df = df.drop_duplicates(subset=['MDM Sort'], keep='first')
        removed_count = original_count - len(df)
        print(f"Removed {removed_count} duplicate records. {len(df)} unique records remaining.")

        # Convert DataFrame to CSV string
        csv_data = df.to_csv(index=False, escapechar='\\', doublequote=True)
        
        # Set up API request
        headers = {
            'QB-Realm-Hostname': 'wesco.quickbase.com',
            'Authorization': 'QB-USER-TOKEN cacrrx_vcs_0_ezvd3icw7ds8wdegdjbwbigxm45',
            'Content-Type': 'application/xml'
        }

        # Create the clist parameter
        column_list = df.columns.tolist()
        clist = '.'.join([str(field_mapping[col]) for col in column_list])
        
        # Create XML request
        xml_request = f"""<?xml version="1.0" encoding="UTF-8" ?>
        <qdbapi>
            <apptoken>None</apptoken>
            <udata>mydata</udata>
            <records_csv><![CDATA[{csv_data}]]></records_csv>
            <clist>{clist}</clist>
            <skipfirst>1</skipfirst>
        </qdbapi>"""
        
        # Send request
        api_url = 'https://wesco.quickbase.com/db/brrb3vdk5'
        
        response = requests.post(
            f"{api_url}?a=API_ImportFromCSV",
            headers=headers,
            data=xml_request.encode('utf-8'),
            verify=False
        )
        
        print(f"\nResponse Status: {response.status_code}")
        print(f"Response Content: {response.text}")
        
        if response.status_code in [200, 201]:
            if '<errcode>0</errcode>' in response.text:
                records_added = re.search(r'<num_recs_added>(\d+)</num_recs_added>', response.text)
                if records_added:
                    num_added = records_added.group(1)
                    print(f"\nSuccess! Added {num_added} records to QuickBase")
                    return True
                else:
                    print("\nSuccess! Records added to QuickBase")
                    return True
            else:
                error_text = re.search(r'<errtext>(.*?)</errtext>', response.text)
                if error_text:
                    print(f"Upload failed: {error_text.group(1)}")
                else:
                    print("Upload failed with unknown error")
                return False
        else:
            print(f"Upload failed with status code: {response.status_code}")
            print("Error response:", response.text)
            return False
            
    except Exception as e:
        print(f"Error in upload process: {str(e)}")
        print("Full error traceback:")
        import traceback
        print(traceback.format_exc())
        return False

def main():
   print("Initializing SharePoint connection...")
   ctx = get_sharepoint_context()
   
   if ctx is None:
       print("Failed to connect to SharePoint. Exiting...")
       return
       
   last_check_time = datetime.now()
   print(f"Starting monitoring at: {last_check_time}")
   print("Monitoring SharePoint folder for new MDM files...")
   
   while True:
       try:
           print(f"\nChecking for new files at: {datetime.now()}")
           
           # First delete existing QuickBase records
           if delete_quickbase_records():
               print("Successfully deleted existing QuickBase records")
               
               # Then check for new files
               new_files = check_new_files(ctx, last_check_time)
               
               if new_files:
                   for file in new_files:
                       print(f"\nProcessing new file: {file.properties['Name']}")
                       
                       try:
                           # Download file content
                           file_content = file.read()
                           
                           # Create output filename
                           output_file = os.path.join(
                               # r"\\Wshqnt4sdata\dira\General Data and Automation\Quickbase2024\QB Update Files\QB MDM Files\PSEG",
                               r"C:\Users\sabar\Documents\QB MDM Updates",                              
                               file.properties["Name"].replace('.xlsb', '.csv')
                           )
                           
                           # Transform and upload file
                           if transform_mdm_file(file_content, output_file):
                               print(f"Successfully processed and uploaded file: {file.properties['Name']}")
                           else:
                               print(f"Failed to process file: {file.properties['Name']}")
                               
                       except Exception as file_error:
                           print(f"Error processing file {file.properties['Name']}: {str(file_error)}")
                           continue
               else:
                   print("No new files found to process")
           else:
               print("Failed to delete existing QuickBase records. Skipping this cycle.")
           
           # Update last check time
           last_check_time = datetime.now()
           
           print(f"Waiting 5 minutes before next check...")
           time.sleep(300)  # Check every 5 minutes
           
       except Exception as e:
           print(f"Error in main loop: {str(e)}")
           print("Full error details:")
           import traceback
           print(traceback.format_exc())
           print("Waiting 1 minute before retrying...")
           time.sleep(60)  # Wait a minute before retrying

if __name__ == "__main__":
   main()