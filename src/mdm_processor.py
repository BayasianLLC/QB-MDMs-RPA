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
   
   # SharePoint credentials and site URL
   # sharepoint_url = "https://wescodist.sharepoint.com/sites/SalesOpsRPA"
   # username = "JuanCarlos.Bayas@wescodist.com"
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
        # Use BytesIO for Excel file
        excel_data = BytesIO(file_content)
        df = pd.read_excel(excel_data, engine='pyxlsb')
        
        print("File read successfully. Processing data...")
        df.columns = df.iloc[0]
        df = df.iloc[:, :88]
        df = df.iloc[2:].reset_index(drop=True)
        
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
        df = df.replace({pd.NA: None, 'nan': None})
        df = df.where(pd.notnull(df), None)
        
        # Convert to QuickBase format
        records = df.to_dict('records')
        
        # Split records into batches
        batches = [records[i:i + batch_size] for i in range(0, len(records), batch_size)]
        print(f"Split data into {len(batches)} batches of {batch_size} records each")
        
        headers = {
            'Content-Type': 'application/json',
            'QB-Realm-Hostname': 'wesco.quickbase.com',
            'User-Agent': 'PSEG_MDM_Integration_V1.0',
            'Authorization': 'QB-USER-TOKEN cacrrx_vcs_0_ezvd3icw7ds8wdegdjbwbigxm45'
        }
        
        params = {
            'appId': 'bfdix6cda'
        }
        
        table_id = 'butqctiz3'
        api_url = f'https://api.quickbase.com/v1/tables/{table_id}'
        
        # Upload batches
        total_uploaded = 0
        for i, batch in enumerate(batches, 1):
            print(f"\nUploading batch {i} of {len(batches)}...")
            
            response = requests.post(
                api_url,
                params=params,
                headers=headers,
                json={'data': batch},
                verify=False
            )
            
            if response.status_code == 200:
                total_uploaded += len(batch)
                print(f"Batch {i} uploaded successfully. Progress: {total_uploaded}/{total_records}")
            else:
                print(f"Batch {i} upload failed with status code: {response.status_code}")
                print(response.text)
                return False
            
            # Add a small delay between batches
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
                   r"C:\Users\sabar\Documents\QB MDM Updates",
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