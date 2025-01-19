import json
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
   sharepoint_url = "https://wescodist.sharepoint.com/sites/UtilityMDMs-PSEG"
   username = "juan.bayas@wescodist.com"
   password = "DhkofiL@512345"
   
  # sharepoint_url = "https://stdntpartners.sharepoint.com/sites/MDMQB"
  #  username = "Victor.Sabare@studentambassadors.com"
  #  password = "ni2b:+AANpP?N7w"


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
        folder_path = f"{web_url}/Shared%20Documents/PSEG/MDM%20Files"
        
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
        # Use BytesIO for Excel file
        excel_data = BytesIO(file_content)
        df = pd.read_excel(excel_data, engine='pyxlsb')
        
        print("File read successfully. Processing data...")
        df.columns = df.iloc[0]
        df = df.iloc[:, :88]
        df = df.iloc[3:].reset_index(drop=True)
        
        print(f"Saving processed file to: {output_file}")
        df.to_csv(output_file, index=False)
        
        # Upload to QuickBase
        if upload_to_quickbase(output_file):
            print("File successfully uploaded to QuickBase")
            return True
        else:
            print("Failed to upload to QuickBase")
            return False
        
    except Exception as e:
        print(f"Error processing file: {str(e)}")
        return False


def upload_to_quickbase(csv_file, batch_size=1000):
    try:
        print("Initiating QuickBase upload...")
        
        # Read CSV file with all columns as string type
        df = pd.read_csv(csv_file, dtype=str, low_memory=False)
        total_records = len(df)
        print(f"Read {total_records} records from CSV")
        
        # Replace NaN values with None
        df = df.replace({pd.NA: None, 'nan': None, 'NaN': None, '': None})
        df = df.where(pd.notnull(df), None)
        
        # Convert to QuickBase format with field IDs
        records = []
        for _, row in df.iterrows():
            record = {
                '6': row['MDM Sort'],
                '7': row['Date Added'],
                '8': row['In Scope'],
                '9': row['Servicing Business Unit'],
                '10': row['Pricing Category / Owner'],
                '11': row['Product Category'],
                '12': row['Product Sub-Category'],
                '13': row['Cust. ID #'],
                '14': row['Main Category'],
                '15': row['Short Description'],
                '16': row['Long Description'],
                '17': row['UOP'],
                '18': row['Last 12 Usage'],
                '19': row['Annual Times Purchased'],
                '20': row['Manufacturer'],
                '21': row['Manufacturer Part #'],
                '22': row['Manufacturer Status'],
                '23': row['Customer Info Change Date'],
                '24': row['VMI (Y/N)'],
                '25': row['Customer Comments'],
                '26': row['Sugg. Sell Price'],
                '27': row['Sugg. Sell Price Extended'],
                '28': row['Markup'],
                '29': row['Billing Margin %'],
                '30': row['Extended Billing Margin $'],
                '31': row['Item Review Notes'],
                '32': row['Vendor Name'],
                '33': row['Vendor Code'],
                '34': row['Blanket #'],
                '35': row['Blanket Load Price'],
                '36': row['Blanket Load Standard Pack'],
                '37': row['Blanket Load Leadtime'],
                '38': row['Blanket Load Date'],
                '39': row['Source'],
                '40': row['Source Manufacturer'],
                '41': row['Source Supplier #'],
                '42': row['SIM'],
                '43': row['Sim MFR'],
                '44': row['Sim Item'],
                '45': row['Wesnet Catalog #'],
                '46': row['Wesnet SIM Description'],
                '47': row['Wesnet UOM'],
                '48': row['Source Count'],
                '49': row['Primary Supplier'],
                '50': row['Rank'],
                '51': row['Low Cost'],
                '52': row['Cost Source'],
                '53': row['Cost Extended'],
                '54': row['Customer UOP Factor'],
                '55': row['Supplier UOP Factor'],
                '56': row['Spa Cost'],
                '57': row['Spa Into Stock Cost'],
                '58': row['Spa Number'],
                '59': row['Spa Start Date'],
                '60': row['Spa End Date'],
                '61': row['DC Xfer'],
                '62': row['8500 Repl Cost'],
                '63': row['8500 Repl Cost Extended'],
                '64': row['8520 Repl Cost'],
                '65': row['8520 Repl Cost Extended'],
                '66': row['Tier Cost'],
                '67': row['UOM'],
                '68': row['Standard Pack'],
                '69': row['Leadtime'],
                '70': row['Future Quote Loaded'],
                '71': row['Last Date Quote Modified'],
                '72': row['Quoted Mfr / Brand'],
                '73': row['Quoted Mfr Part Number'],
                '74': row['Direct Equal'],
                '75': row['Returnable'],
                '76': row['Supplier Comments'],
                '77': row['Quoted Price'],
                '78': row['List Price'],
                '79': row['Unit of Measure'],
                '80': row['Qty per Unit of Measure'],
                '81': row['Std Purchase Qty'],
                '82': row['Lead Time (Calendar Days)'],
                '83': row['Quote #'],
                '84': row['Quote End Date'],
                '85': row['Minimum Order'],
                '86': row['Freight Terms'],
                '87': row['Quote - Contact / Preparer Name'],
                '88': row['Quote - Contact Phone'],
                '89': row['Quote - Contact E-mail'],
                '90': row['Purchasing - Contact Name'],
                '91': row['Purchasing - Contact E-mail'],
                '92': row['Purchasing - Contact Fax']
            }
            records.append(record)
        
        # Split records into batches
        batches = [records[i:i + batch_size] for i in range(0, len(records), batch_size)]
        print(f"Split data into {len(batches)} batches of {batch_size} records each")
        
        headers = {
            'Content-Type': 'application/json',
            'QB-Realm-Hostname': 'wesco.quickbase.com',
            'User-Agent': 'PSEG_MDM_Integration_V1.0',
            'Authorization': 'QB-USER-TOKEN cacrrx_vcs_0_ezvd3icw7ds8wdegdjbwbigxm45'
        }
        
        # Use the correct endpoint for record creation
        api_url = 'https://api.quickbase.com/v1/records'
        
        params = {
            'tableId': 'bs2u8eeps'  # Your table ID
        }
        
        # Upload batches
        total_uploaded = 0
        for i, batch in enumerate(batches, 1):
            print(f"\nUploading batch {i} of {len(batches)}...")
            
            payload = {
                "to": "bs2u8eeps",  # Your table ID
                "data": batch
            }
            
            response = requests.post(
                api_url,
                headers=headers,
                json=payload,
                verify=False
            )
            
            if response.status_code in [200, 201]:
                total_uploaded += len(batch)
                print(f"Batch {i} uploaded successfully. Progress: {total_uploaded}/{total_records}")
                print("Response:", response.json())  # Print response for debugging
            else:
                print(f"Batch {i} upload failed with status code: {response.status_code}")
                print("Error response:", response.text)
                return False
            
            time.sleep(1)
        
        print(f"\nUpload completed successfully! Total records uploaded: {total_uploaded}")
        return True
            
    except Exception as e:
        print(f"Error uploading to QuickBase: {str(e)}")
        import traceback
        print("Full error traceback:")
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
           # Check for new files
           new_files = check_new_files(ctx, last_check_time)
           
           for file in new_files:
               print(f"\nProcessing new file: {file.properties['Name']}")
               
               # Download file content
               file_content = file.read()
               
               # Create output filename
               output_file = os.path.join(
                   r"\\Wshqnt4sdata\dira\General Data and Automation\Quickbase2024\QB Update Files\QB MDM Files\PSEG",
                   file.properties["Name"].replace('.xlsb', '.csv')
               )
               
               # Transform file
               transform_mdm_file(file_content, output_file)
           
           # Update last check time
           last_check_time = datetime.now()
           
           # Wait before next check
           print(f"Waiting 5 minutes before next check...")
           time.sleep(300)  # Check every 5 minutes
           
       except Exception as e:
           print(f"Error in main loop: {str(e)}")
           print("Waiting 1 minute before retrying...")
           time.sleep(60)  # Wait a minute before retrying


if __name__ == "__main__":
   main()