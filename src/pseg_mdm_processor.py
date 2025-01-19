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

        df['MDM Sort'] = pd.to_numeric(df['MDM_Sort'], errors='coerce').fillna(0)
        
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
        
        # Read CSV with all columns as string
        df = pd.read_csv(csv_file, dtype=str, low_memory=False)
        
        # Replace NaN values with empty string instead of None for required fields
        df = df.fillna('')  # Fill all NaN with empty string first
        
        # Convert to QuickBase format with field IDs
        records = []
        for index, row in df.iterrows():
            # Ensure MDM Sort (field 6) has a value
            mdm_sort = row.get('MDM Sort', '')
            if not mdm_sort:
                print(f"Warning: Row {index + 1} missing required MDM Sort value, setting to 0")
                mdm_sort = '0'  # Default value for required numeric field
                
            record = {
                '6': mdm_sort,  # Required field
                '7': row.get('Date Added', ''),
                '8': row.get('In Scope', ''),
                # ... rest of the fields ...
            }
            # Remove empty values to prevent API errors
            record = {k: v for k, v in record.items() if v != ''}
            records.append(record)

        print(f"Processed {len(records)} records")
        
        # Split into batches
        batches = [records[i:i + batch_size] for i in range(0, len(records), batch_size)]
        
        total_uploaded = 0
        for i, batch in enumerate(batches, 1):
            print(f"\nUploading batch {i} of {len(batches)}...")
            
            try:
                response = requests.post(
                    'https://api.quickbase.com/v1/records',
                    headers={
                        'QB-Realm-Hostname': 'wesco.quickbase.com',
                        'Authorization': 'QB-USER-TOKEN cacrrx_vcs_0_ezvd3icw7ds8wdegdjbwbigxm45',
                        'Content-Type': 'application/json'
                    },
                    json={
                        "to": "bs2u8eeps",
                        "data": batch
                    },
                    verify=False
                )
                
                if response.status_code in [200, 201]:
                    total_uploaded += len(batch)
                    print(f"Batch {i} successful: {len(batch)} records")
                else:
                    print(f"Batch {i} failed. Status: {response.status_code}")
                    print(f"Error: {response.text}")
                    return False
                    
            except Exception as e:
                print(f"Error in batch {i}: {str(e)}")
                return False
                
        print(f"Upload completed. Total records: {total_uploaded}")
        return True
        
    except Exception as e:
        print(f"Error: {str(e)}")
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