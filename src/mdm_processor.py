import pandas as pd
import os
from datetime import datetime
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
import time
from quickbase_client import QuickBaseClient

def get_sharepoint_context():
   # SharePoint credentials and site URL
   sharepoint_url = "https://stdntpartners.sharepoint.com/sites/MasterDataRequest"
   username = "victor.sabare@studentambassadors.com"
   password = "ni2b:+AANpP?N7w"  # Remember to secure this
   
   try:
       ctx = ClientContext(sharepoint_url).with_credentials(
           UserCredential(username, password)
       )
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
        folder_path = f"{web_url}/Shared Documents/Forms/AllItems.aspx"
        
        print(f"Accessing folder: {folder_path}")
        
        # Get files from SharePoint folder
        folder = ctx.web.get_folder_by_server_relative_url(folder_path)
        files = folder.files
        ctx.load(files)
        files.execute_query()
        
        # Look for new XLSB files
        new_files = [f for f in files 
                    if f.properties["TimeCreated"] > last_check_time 
                    and "PSEG MDM" in f.properties["Name"]]
        
        print(f"Found {len(new_files)} new files")
        return new_files
    except Exception as e:
        print(f"Error checking SharePoint: {str(e)}")
        return []

def transform_mdm_file(file_content, output_file):
    try:
        print("Starting file transformation...")
        # Read XLSB from memory
        df = pd.read_excel(file_content, engine='pyxlsb')
        
        print("File read successfully. Processing data...")
        # Get the first row as column names
        df.columns = df.iloc[0]
        
        # Keep first 88 columns
        df = df.iloc[:, :88]
        
        # Remove the original header row and the row after it
        df = df.iloc[2:].reset_index(drop=True)
        
        print(f"Saving processed file to: {output_file}")
        # Save to CSV
        df.to_csv(output_file, index=False)
        print(f"Successfully transformed file to: {output_file}")
        
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
                   r"C:\Users\sabar\Desktop",
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

def upload_to_quickbase(csv_file):
    try:
        print("Initiating QuickBase upload...")
        
        # QuickBase configuration
        qb_client = QuickBaseClient({
            'realm_hostname': 'your_quickbase_domain.quickbase.com',
            'user_token': 'your_user_token',  # QB API token
            'app_id': 'your_app_id',          # QB application ID
            'table_id': 'your_table_id'       # QB table ID
        })
        
        # Read CSV file
        with open(csv_file, 'rb') as f:
            csv_data = f.read()
        
        # Upload to QuickBase
        response = qb_client.import_from_csv(
            table_id='your_table_id',
            csv_file=csv_data,
            merge_field_id=None,  # Set if you want to update existing records
            import_as_admin=True
        )
        
        print(f"QuickBase upload successful. Records processed: {response.get('number_records_processed', 0)}")
        return True
        
    except Exception as e:
        print(f"Error uploading to QuickBase: {str(e)}")
        return False

if __name__ == "__main__":
   main()