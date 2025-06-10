import os
import requests
from ms_graph import generate_access_token, GRAPH_API_ENDPOINT
# === Configuration ===
APP_ID = '33e576cd-e2db-4a05-8778-71c7f799375f'
SCOPES = ['Files.Read']
FILE_PATH = "Protocol Automation EXCEL Grid.xlsx"  # path in OneDrive
# Local file location
save_location = os.getcwd()
# Get token from cache or login
access_token = generate_access_token(APP_ID, SCOPES)
headers = {
   'Authorization': 'Bearer ' + access_token['access_token']
}
# Download URL (OneDrive - self)
download_url = f"{GRAPH_API_ENDPOINT}/me/drive/root:/{FILE_PATH}:/content"
# Download request
response = requests.get(download_url, headers=headers, verify=False)
# Save the file
if response.status_code == 200:
   local_path = os.path.join(save_location, os.path.basename(FILE_PATH))
   with open(local_path, "wb") as f:
       f.write(response.content)
   print(f"✅ File downloaded and saved to: {local_path}")
else:
   print(f"❌ Download failed: {response.status_code} — {response.text}")