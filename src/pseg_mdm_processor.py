import pandas as pd
import os
from datetime import datetime
import requests
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
import time
import html
import re
import urllib3

from io import BytesIO

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def get_sharepoint_context():
   
   #SharePoint credentials and site URL
   # sharepoint_url = "https://wescodist.sharepoint.com/sites/UtilityMDMs-PSEG"
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
        # folder_path = f"{web_url}/Shared%20Documents/PSEG/MDM%20Files"
        folder_path = f"{web_url}/Shared%20Documents"
        
        print(f"Accessing folder: {folder_path}")
   
        # Get files from SharePoint folder
        folder = ctx.web.get_folder_by_server_relative_url(folder_path)
        files = folder.files
        ctx.load(files)
        files.execute_query()
        
        # Look for new XLSB files
        new_files = [f for f in files 
                    if "PSEG MDM" in f.properties["Name"]]
        
        print(f"Found {len(new_files)} new files")
        return new_files
    except Exception as e:
        print(f"Error checking SharePoint: {str(e)}")
        return []

def transform_mdm_file(file_content, output_file):
    try:
        print("Starting file transformation...")
        excel_data = BytesIO(file_content)
        df = pd.read_excel(excel_data, engine='pyxlsb')
        
        # Print columns before any transformation
        print("\nOriginal columns:")
        for i, col in enumerate(df.columns):
            print(f"Column {i}: {col}")
        
        # Get first row for column names
        first_row = df.iloc[0]
        print("\nFirst row values:")
        for i, val in enumerate(first_row):
            print(f"Column {i}: {val}")
            
        # Set proper column names
        column_mapping = {
            df.columns[0]: 'MDM Sort',
            df.columns[1]: 'Added By',
            df.columns[2]: 'Date Added',
            df.columns[3]: 'In Scope',
            df.columns[4]: 'Servicing Business Unit',
            df.columns[5]: 'Pricing Category / Owner',
            df.columns[6]: 'Product Category',
            df.columns[7]: 'Product Sub-Category',
            df.columns[8]: 'Cust. ID #',
            df.columns[9]: 'Main Category',
            df.columns[10]: 'Short Description',
            df.columns[11]: 'Long Description',
            df.columns[12]: 'UOP',
            df.columns[13]: 'Last 12 Usage',
            df.columns[14]: 'Annual Times Purchased',
            df.columns[15]: 'Manufacturer',
            df.columns[16]: 'Manufacturer Part #',
            df.columns[17]: 'Manufacturer Status',
            df.columns[18]: 'Customer Info Change Date',
            df.columns[19]: 'VMI (Y/N)',
            df.columns[20]: 'Customer Comments',
            df.columns[21]: 'Sugg. Sell Price',
            df.columns[22]: 'Sugg. Sell Price Extended',
            df.columns[23]: 'Markup',
            df.columns[24]: 'Billing Margin %',
            df.columns[25]: 'Extended Billing Margin $',
            df.columns[26]: 'Item Review Notes',
            df.columns[27]: 'Vendor Name',
            df.columns[28]: 'Vendor Code',
            df.columns[29]: 'Blanket #',
            df.columns[30]: 'Blanket Load Price',
            df.columns[31]: 'Blanket Load Standard Pack',
            df.columns[32]: 'Blanket Load Leadtime',
            df.columns[33]: 'Blanket Load Date',
            df.columns[34]: 'Source',
            df.columns[35]: 'Source Manufacturer',
            df.columns[36]: 'Source Supplier #',
            df.columns[37]: 'SIM',
            df.columns[38]: 'Sim MFR',
            df.columns[39]: 'Sim Item',
            df.columns[40]: 'Wesnet Catalog #',
            df.columns[41]: 'Wesnet SIM Description',
            df.columns[42]: 'Wesnet UOM',
            df.columns[43]: 'Source Count',
            df.columns[44]: 'Primary Supplier',
            df.columns[45]: 'Rank',
            df.columns[46]: 'Low Cost',
            df.columns[47]: 'Cost Source',
            df.columns[48]: 'Cost Extended',
            df.columns[49]: 'Customer UOP Factor',
            df.columns[50]: 'Supplier UOP Factor',
            df.columns[51]: 'Spa Cost',
            df.columns[52]: 'Spa Into Stock Cost',
            df.columns[53]: 'Spa Number',
            df.columns[54]: 'Spa Start Date',
            df.columns[55]: 'Spa End Date',
            df.columns[56]: 'DC Xfer',
            df.columns[57]: '8500 Repl Cost',
            df.columns[58]: '8500 Repl Cost Extended',
            df.columns[59]: '8520 Repl Cost',
            df.columns[60]: '8520 Repl Cost Extended',
            df.columns[61]: 'Tier Cost',
            df.columns[62]: 'UOM',
            df.columns[63]: 'Standard Pack',
            df.columns[64]: 'Leadtime',
            df.columns[65]: 'Future Quote Loaded',
            df.columns[66]: 'Last Date Quote Modified',
            df.columns[67]: 'Quoted Mfr / Brand',
            df.columns[68]: 'Quoted Mfr Part Number',
            df.columns[69]: 'Direct Equal',
            df.columns[70]: 'Returnable',
            df.columns[71]: 'Supplier Comments',
            df.columns[72]: 'Quoted Price',
            df.columns[73]: 'List Price',
            df.columns[74]: 'Unit of Measure',
            df.columns[75]: 'Qty per Unit of Measure',
            df.columns[76]: 'Std Purchase Qty',
            df.columns[77]: 'Lead Time (Calendar Days)',
            df.columns[78]: 'Quote #',
            df.columns[79]: 'Quote End Date',
            df.columns[80]: 'Minimum Order',
            df.columns[81]: 'Freight Terms',
            df.columns[82]: 'Quote - Contact / Preparer Name',
            df.columns[83]: 'Quote - Contact Phone',
            df.columns[84]: 'Quote - Contact E-mail',
            df.columns[85]: 'Purchasing - Contact Name',
            df.columns[86]: 'Purchasing - Contact E-mail',
            df.columns[87]: 'Purchasing - Contact Fax'
         }       
        
        # Rename columns
        df = df.rename(columns=column_mapping)
        
        # Print columns after renaming
        print("\nColumns after renaming:")
        print(df.columns.tolist())
        
        # Remove header row
        df = df.iloc[1:]
    
        # Keep first 87 coumns
        df = df.iloc[:, :88]
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
        api_url = 'https://wesco.quickbase.com/db/bust272rx'
        
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
        
        # Field mapping
        field_mapping = {
            'MDM Sort': 6,
            'Date Added': 7,
            'In Scope': 8,
            'Servicing Business Unit': 9,
            'Pricing Category / Owner': 10,
            'Product Category': 11,
            'Product Sub-Category': 12,
            'Cust. ID #': 13,
            'Main Category': 14,
            'Short Description': 15,
            'Long Description': 16,
            'UOP': 17,
            'Last 12 Usage': 18,
            'Annual Times Purchased': 19,
            'Manufacturer': 20,
            'Manufacturer Part #': 21,
            'Manufacturer Status': 22,
            'Customer Info Change Date': 23,
            'VMI (Y/N)': 24,
            'Customer Comments': 25,
            'Sugg. Sell Price': 26,
            'Sugg. Sell Price Extended': 27,
            'Markup': 28,
            'Billing Margin %': 29,
            'Extended Billing Margin $': 30,
            'Item Review Notes': 31,
            'Vendor Name': 32,
            'Vendor Code': 33,
            'Blanket #': 34,
            'Blanket Load Price': 35,
            'Blanket Load Standard Pack': 36,
            'Blanket Load Leadtime': 37,
            'Blanket Load Date': 38,
            'Source': 39,
            'Source Manufacturer': 40,
            'Source Supplier #': 41,
            'SIM': 42,
            'Sim MFR': 43,
            'Sim Item': 44,
            'Wesnet Catalog #': 45,
            'Wesnet SIM Description': 46,
            'Wesnet UOM': 47,
            'Source Count': 48,
            'Primary Supplier': 49,
            'Rank': 50,
            'Low Cost': 51,
            'Cost Source': 52,
            'Cost Extended': 53,
            'Customer UOP Factor': 54,
            'Supplier UOP Factor': 55,
            'Spa Cost': 56,
            'Spa Into Stock Cost': 57,
            'Spa Number': 58,
            'Spa Start Date': 59,
            'Spa End Date': 60,
            'DC Xfer': 61,
            '8500 Repl Cost': 62,
            '8500 Repl Cost Extended': 63,
            '8520 Repl Cost': 64,
            '8520 Repl Cost Extended': 65,
            'Tier Cost': 66,
            'UOM': 67,
            'Standard Pack': 68,
            'Leadtime': 69,
            'Future Quote Loaded': 70,
            'Last Date Quote Modified': 71,
            'Quoted Mfr / Brand': 72,
            'Quoted Mfr Part Number': 73,
            'Direct Equal': 74,
            'Returnable': 75,
            'Supplier Comments': 76,
            'Quoted Price': 77,
            'List Price': 78,
            'Unit of Measure': 79,
            'Qty per Unit of Measure': 80,
            'Std Purchase Qty': 81,
            'Lead Time (Calendar Days)': 82,
            'Quote #': 83,
            'Quote End Date': 84,
            'Minimum Order': 85,
            'Freight Terms': 86,
            'Quote - Contact / Preparer Name': 87,
            'Quote - Contact Phone': 88,
            'Quote - Contact E-mail': 89,
            'Purchasing - Contact Name': 90,
            'Purchasing - Contact E-mail': 91,
            'Purchasing - Contact Fax': 92
        }

        # Read CSV into DataFrame
        df = pd.read_csv(csv_file, dtype=str)
        total_records = len(df)
        print(f"Read {total_records} records from CSV")

        # Keep only columns that exist in QuickBase
        columns_to_keep = list(field_mapping.keys())
        extra_columns = set(df.columns) - set(columns_to_keep)
        if extra_columns:
            print(f"\nDropping columns not needed in QuickBase: {', '.join(extra_columns)}")
            df = df[columns_to_keep]
        
        # Keep only first occurrence of each MDM Sort value
        print("\nRemoving duplicate MDM Sort values...")
        original_count = len(df)
        df = df.drop_duplicates(subset=['MDM Sort'], keep='first')
        removed_count = original_count - len(df)
        print(f"Removed {removed_count} duplicate records. {len(df)} unique records remaining.")

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

        def format_phone(phone_str):
            if pd.isna(phone_str) or phone_str == '':
                return ''
            phone = re.sub(r'[^\d]', '', str(phone_str))
            if len(phone) == 10:
                return f"({phone[:3]}) {phone[3:6]}-{phone[6:]}"
            return phone_str

        def format_yes_no(val):
            if pd.isna(val) or val == '':
                return ''
            val = str(val).upper()
            if val in ['Y', 'YES', 'TRUE', '1']:
                return 'Y'
            if val in ['N', 'NO', 'FALSE', '0']:
                return 'N'
            return val

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

        # Define column groups
        date_columns = [
            'Date Added', 'Customer Info Change Date', 'Blanket Load Date', 
            'Spa Start Date', 'Spa End Date', 'Last Date Quote Modified', 
            'Quote End Date'
        ]
        
        numeric_columns = [
            'Last 12 Usage', 'Annual Times Purchased', 'Sugg. Sell Price',
            'Sugg. Sell Price Extended', 'Markup', 'Billing Margin %',
            'Extended Billing Margin $', 'Blanket Load Price', 'Standard Pack',
            'Qty per Unit of Measure', 'Std Purchase Qty', 'Lead Time (Calendar Days)',
            '8500 Repl Cost', '8500 Repl Cost Extended', '8520 Repl Cost',
            '8520 Repl Cost Extended', 'Tier Cost', 'DC Xfer', 'Vendor Code',
            'Blanket #', 'SIM', 'Sim MFR', 'Sim Item'
        ]
        
        phone_columns = ['Quote - Contact Phone', 'Purchasing - Contact Fax']
        
        yes_no_columns = [
            'In Scope', 'VMI (Y/N)', 'Future Quote Loaded',
            'Direct Equal', 'Returnable'
        ]
        
        email_columns = ['Quote - Contact E-mail', 'Purchasing - Contact E-mail']

        # Clean data
        df = df.replace({pd.NA: '', 'nan': '', 'NaN': '', None: ''})
        df = df.fillna('')

        # Apply formatting
        for col in df.columns:
            if col in date_columns:
                df[col] = df[col].apply(format_date)
            elif col in numeric_columns:
                df[col] = df[col].apply(format_number)
            elif col in phone_columns:
                df[col] = df[col].apply(format_phone)
            elif col in yes_no_columns:
                df[col] = df[col].apply(format_yes_no)
            elif col in email_columns:
                df[col] = df[col].apply(lambda x: str(x).strip() if not pd.isna(x) else '')
            else:
                df[col] = df[col].apply(clean_text)

        # Convert DataFrame to CSV string
        csv_data = df.to_csv(index=False, escapechar='\\', doublequote=True)
        
        # Set up API request
        headers = {
            'QB-Realm-Hostname': 'wesco.quickbase.com',
            'Authorization': 'QB-USER-TOKEN cacrrx_vcs_0_ezvd3icw7ds8wdegdjbwbigxm45',
            'Content-Type': 'application/xml'
        }

        # Create the clist parameter based on the column order in the DataFrame
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
        api_url = 'https://wesco.quickbase.com/db/bust272rx'
        
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