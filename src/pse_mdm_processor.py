import json
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
   # sharepoint_url = "https://wescodist.sharepoint.com/sites/UtilityMDMs-PSE"
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
        # folder_path = f"{web_url}/Shared%20Documents/PSE/MDMs"
        folder_path = f"{web_url}/Shared%20Documents"
        
        print(f"Accessing folder: {folder_path}")
        
        # Get files from SharePoint folder
        folder = ctx.web.get_folder_by_server_relative_url(folder_path)
        files = folder.files
        ctx.load(files)
        files.execute_query()
        
        # Look for new XLSB or XLSM files
        new_files = [f for f in files 
                    if "PSE WCDM" in f.properties["Name"] 
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

        # Get the second row which contains the actual headers (row index 1)
        headers_row = df.iloc[1]
        
        # Create mapping based on column positions
        column_mapping = {
            0: 'MDM Sort',
            1: 'Added By',
            2: 'Date Added',
            3: 'In Scope',
            4: 'Servicing Business Unit',
            5: 'Pricing Category / Owner',
            6: 'Product Category',
            7: 'Product Sub-Category',
            8: 'Cust. ID #',
            9: 'Main Category',
            10: 'Long Description',
            11: 'UOP',
            12: 'Last 12 Purchases',
            13: 'Last 12 Times Purchased',
            14: 'Manufacturer',
            15: 'Manufacturer Part #',
            16: 'Manufacturer Status',
            17: 'Customer Info Change Date',
            18: 'STK Req.',
            19: 'Strom STK Req.',
            20: 'Customer Comments',
            21: 'Sugg. Sell Price',
            22: 'Sugg. Sell Price Extended',
            23: 'Margin',
            24: 'Billing Margin %',
            25: 'Extended Billing Margin $',
            26: 'Item Review Notes',
            27: 'Vendor Name',
            28: 'Vendor Code',
            29: 'Blanket #',
            30: 'Blanket Load Price',
            31: 'Blanket Load Standard Pack',
            32: 'Blanket Load Leadtime',
            33: 'Blanket Load Date',
            34: 'Source',
            35: 'Source Manufacturer',
            36: 'Source Supplier #',
            37: 'SIM',
            38: 'Sim MFR',
            39: 'Sim Item',
            40: 'Wesnet Catalog #',
            41: 'Wesnet SIM Description',
            42: 'Wesnet UOM',
            43: 'Source Count',
            44: 'Rank',
            45: 'Low Cost',
            46: 'Cost Source',
            47: 'Cost Extended',
            48: 'UOP Multiplier Factor',
            49: 'UOP Divider Factor',
            50: 'Spa Cost',
            51: 'Spa Into Stock Cost',
            52: 'Spa Number',
            53: 'Spa Start Date',
            54: 'Spa End Date',
            55: 'DC Xfer',
            56: '8500 Low Repl Cost',
            57: '8500 Low Repl Cost Extended',
            58: '8570 Low Repl Cost',
            59: '8570 Low Repl Cost Extended',
            60: 'Future Quote Loaded',
            61: 'Last Date Quote Modified',
            62: 'Quoted Mfr / Brand',
            63: 'Quoted Mfr Part Number',
            64: 'Direct Equal',
            65: 'Returnable',
            66: 'Supplier Comments',
            67: 'Quoted Price',
            68: 'List Price',
            69: 'Unit of Measure',
            70: 'Qty per Unit of Measure',
            71: 'Std Purchase Qty',
            72: 'Lead Time (Calendar Days)',
            73: 'Quote #',
            74: 'Quote End Date',
            75: 'Minimum Order',
            76: 'Freight Terms',
            77: 'Quote - Contact / Preparer Name',
            78: 'Quote - Contact Phone',
            79: 'Quote - Contact E-mail',
            80: 'Purchasing - Contact Name',
            81: 'Purchasing - Contact Phone',
            82: 'Purchasing - Contact E-mail',
            83: 'Last 12',
            84: 'VC',
            85: 'CC',
            86: 'Loaded ORP',
            87: 'Loaded EOQ',
            88: 'On Hand',
            89: 'On Order',
            90: 'On Backorder',
            91: 'Net Stock',
            92: 'WESCO Stocking Item',
            93: 'WESCO Linked Cust ID',
            94: 'Combined Last 12 Purchases',
            95: 'Combined Last 12 Count',
            96: 'MMP Rank',
            97: 'SIM (Y/N)',
            98: 'Supplier Number (Y/N)',
            99: 'Cost (Y/N)',
            100: 'Ready to load (Y/N)',
            101: 'ORP',
            102: 'EOQ',
            103: 'Inventory Max Value',
            104: 'Quote Start Date'
        }

        # Rename columns
        df = df.rename(columns=column_mapping)

        # Keep only the mapped columns
        df = df[list(column_mapping.values())]
        
        # Skip the first two rows (headers) and reset index
        df = df.iloc[2:].reset_index(drop=True)
        
        # Convert MDM Sort to numeric
        df['MDM Sort'] = pd.to_numeric(df['MDM Sort'], errors='coerce')

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
        api_url = 'https://wesco.quickbase.com/db/bs2u8eeps'
        
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
        
        # Field mapping based on the QuickBase field IDs
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
            'Last 12 Purchases': 18,
            'Last 12 Times Purchased': 19,
            'Manufacturer': 20,
            'Manufacturer Part #': 21,
            'Manufacturer Status': 22,
            'Customer Info Change Date': 23,
            'STK Req.': 24,
            'Strom STK Req.': 25,
            'Customer Comments': 26,
            'Sugg. Sell Price': 27,
            'Sugg. Sell Price Extended': 28,
            'Margin': 29,
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
            'Sim MFR': 44,
            'Sim Item': 45,
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
            '8500 Low Repl Cost': 62,
            '8500 Low Repl Cost Extended': 63,
            '8570 Low Repl Cost': 64,
            '8570 Low Repl Cost Extended': 65,
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
            'Purchasing - Contact Phone': 87,
            'Purchasing - Contact E-mail': 88,
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
            'Quote Start Date': 110
        }

        # Read CSV into DataFrame
        df = pd.read_csv(csv_file, dtype=str)
        total_records = len(df)
        print(f"Read {total_records} records from CSV")

        # Define column groups based on data types
        date_columns = [
            'Date Added', 'Customer Info Change Date', 'Blanket Load Date',
            'Spa Start Date', 'Spa End Date', 'Last Date Quote Modified',
            'Quote End Date', 'Quote Start Date'
        ]
        
        numeric_columns = [
            'MDM Sort', 'Last 12 Purchases', 'Last 12 Times Purchased',
            'STK Req.', 'Strom STK Req.', 'Source Count', 'Rank',
            'Qty per Unit of Measure', 'Std Purchase Qty',
            'Lead Time (Calendar Days)', 'Last 12', 'VC',
            'Loaded ORP', 'Loaded EOQ', 'On Hand', 'On Order',
            'On Backorder', 'Net Stock'
        ]
        
        currency_columns = [
            'Sugg. Sell Price', 'Sugg. Sell Price Extended', 'Margin',
            'Extended Billing Margin $', 'Blanket Load Price',
            'Low Cost', 'Spa Cost', 'Spa Into Stock Cost',
            '8500 Low Repl Cost', '8500 Low Repl Cost Extended',
            '8570 Low Repl Cost', '8570 Low Repl Cost Extended',
            'Quoted Price', 'List Price', 'Inventory Max Value'
        ]
        
        percent_columns = ['Billing Margin %']
        
        checkbox_columns = [
            'In Scope', 'Future Quote Loaded', 'Direct Equal', 'Returnable',
            'WESCO Stocking Item', 'Combined Last 12 Purchases',
            'Combined Last 12 Count', 'MMP Rank', 'SIM (Y/N)',
            'Supplier Number (Y/N)', 'Cost (Y/N)', 'Ready to load (Y/N)',
            'ORP', 'EOQ'
        ]
        
        phone_columns = [
            'Quote - Contact Phone', 'Purchasing - Contact Phone'
        ]
        
        email_columns = [
            'Quote - Contact E-mail', 'Purchasing - Contact E-mail'
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
        api_url = 'https://wesco.quickbase.com/db/bs2u8eeps'
        
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