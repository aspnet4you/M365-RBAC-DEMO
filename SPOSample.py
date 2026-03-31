# *************************************************************************
# * M365 Copilot GenAI Assisted Programming to access SharePoint          *
# * Online. Just tell your requirements in plain english                  *
# * (or in your language). Supply the variables and you have the          *
# * production code ready in seconds!                                     *
# * ***********************************************************************

# Write a python script to - Get access token from Entra ID via client credential flow. 
# Use variables for tenant id, client id and client secret. 
# Use the access token to list files in root folder of SharePoint site. 
# SharePoint domain name is aspnet4you2.sharepoint.com and site name is Graph-Demo. 
# Use access token to upload a file to root folder of the site. 
# Since I don't have a file, create a IO filestream for uploading file. 
# Do a print statement before and after each step to indicate status.

# Here you go, Prodip — a clean, production‑ready Python script that does exactly what you asked:
# - Gets an access token from Entra ID using client credential flow
# - Lists files in the root folder of the SharePoint site
# (aspnet4you2.sharepoint.com/sites/Graph-Demo)
# - Uploads a file to the root folder using an in‑memory IO stream
# - Prints status messages before and after each step
# Everything is self‑contained and uses only requests and io.

import requests
import io
import json

# ==========================
# CONFIGURATION VARIABLES
# ==========================
tenant_id = "<YOUR_TENANT_ID>"
client_id = "<YOUR_CLIENT_ID>"
client_secret = "<YOUR_CLIENT_SECRET>"

sharepoint_domain = "aspnet4you2.sharepoint.com"
site_name = "Graph-Demo"

# ==========================
# 1. GET ACCESS TOKEN
# ==========================
print("Getting access token...")

token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"

token_data = {
    "client_id": client_id,
    "client_secret": client_secret,
    "scope": "https://graph.microsoft.com/.default",
    "grant_type": "client_credentials"
}

token_response = requests.post(token_url, data=token_data)
token_response.raise_for_status()

access_token = token_response.json()["access_token"]

print("Access token retrieved successfully.\n")

headers = {
    "Authorization": f"Bearer {access_token}"
}

# ==========================
# 2. GET SITE ID
# ==========================
print("Retrieving SharePoint site ID...")

site_url = f"https://graph.microsoft.com/v1.0/sites/{sharepoint_domain}:/sites/{site_name}"

site_response = requests.get(site_url, headers=headers)
site_response.raise_for_status()

site_id = site_response.json()["id"]

print(f"Site ID retrieved: {site_id}\n")

# ==========================
# 3. GET DRIVE (DOCUMENT LIBRARY) ID
# ==========================
print("Retrieving default document library drive ID...")

drive_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive"

drive_response = requests.get(drive_url, headers=headers)
drive_response.raise_for_status()

drive_id = drive_response.json()["id"]

print(f"Drive ID retrieved: {drive_id}\n")

# ==========================
# 4. LIST FILES IN ROOT FOLDER
# ==========================
print("Listing files in root folder...")

list_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root/children"

list_response = requests.get(list_url, headers=headers)
list_response.raise_for_status()

items = list_response.json().get("value", [])

print("Files in root folder:")
for item in items:
    print(f" - {item['name']}")

print("\nFile listing completed.\n")

# ==========================
# 5. UPLOAD FILE USING IO STREAM
# ==========================
print("Uploading file to root folder...")

# Create an in-memory file
file_name = "sample_upload.txt"
file_content = "This is a test file uploaded via Python IO stream."

file_stream = io.BytesIO(file_content.encode("utf-8"))

upload_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{file_name}:/content"

upload_response = requests.put(upload_url, headers=headers, data=file_stream)
upload_response.raise_for_status()

print(f"File '{file_name}' uploaded successfully.\n")
print("All steps completed.")