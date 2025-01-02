from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
import os
from dotenv import load_dotenv
from datetime import datetime

try:
    now = datetime.now()
    formatDateTime = now.strftime("%d/%m/%Y %H:%M")
except Exception as e:
    with open ('logfile.log', 'a') as file:
        file.write(f"""Problem with date - {str(e)}\n""")

def download_file():
    try:
        # Load env variables
        load_dotenv()
        site_url = os.getenv('site_url')
        file_url = os.getenv('file_url')
        office_username = os.getenv('office_username')
        password = os.getenv('password')
         
        # Set download path file
        current_directory = os.path.dirname(os.path.abspath(__file__))
        local_file_path = os.path.join(current_directory, "Dane_teleadresowe.xlsx")  # Save file here

        # Authenticate and connect
        credentials = UserCredential(office_username, password)
        ctx = ClientContext(site_url).with_credentials(credentials)

        # Download file
        with open(local_file_path, "wb") as file: 
            ctx.web.get_file_by_server_relative_url(file_url).download(file).execute_query()
        
        return True
    except Exception as e:
        with open ('logfile.log', 'a') as file:
            file.write(f"""{formatDateTime} Problem with downloading data - {str(e)}\n""")
        return False

# Run script only if file is downloaded
if download_file():
    pass




