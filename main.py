import json
import gspread
from google.oauth2.service_account import Credentials
import requests
import os
from dotenv import load_dotenv

load_dotenv()

# API request for last seven days tickets
# Load environment variables for API call
api_zone = os.getenv("AUTOTASK_ZONE")
api_tracking_id = os.getenv("AUTOTASK_TRACKING_ID")
api_username = os.getenv("AUTOTASK_USERNAME")
api_secret = os.getenv("AUTOTASK_SECRET")

# Query
query_url = f"https://{api_zone}.autotask.net/atservicesrest/v1.0/tickets/query"
filter_data = {
    "filter": [
        {"op": "gte", "field": "CreateDate", "value": "2024-05-27T00:00:00Z"},
        {"op": "lte", "field": "CreateDate", "value": "2024-05-28T00:00:00Z"}
    ]
}

params = {"search": json.dumps(filter_data)}

headers = {
    "ApiIntegrationCode": api_tracking_id,
    "UserName": api_username,
    "Secret": api_secret,
    "Content-Type": "application/json"
}

# Make request
data = None
response = requests.get(query_url, params=params, headers=headers)
response.raise_for_status()
if response.status_code == 200:
    data = response.json()
    print(data)
else:
    print("Error:", response.status_code)

# Define authentication variables
scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.file'
]
credentials = Credentials.from_service_account_file('autotask-424212-324884f166e1.json', scopes=scopes)
client = gspread.authorize(credentials)

# Specify Sheet
spread_sheet_id = "1U1ef856r0528NaNLHWl9e8WFR_v-NOlF-JTEE-tvOUI"
sheet = client.open_by_key(spread_sheet_id).sheet1
values = []

# Handle dictionary response
if 'items' in data:
    for item in data['items']:
        row_values = []
        for key, value in item.items():
            # Handle nested dictionaries and lists
            if isinstance(value, dict) or isinstance(value, list):
                value = json.dumps(value)  # Convert to JSON string
            row_values.append(value)
        values.append(row_values)
else:
    print("Error: 'items' key not found in the response dictionary.")

headers = list(data['items'][0].keys())

print("Headers:", headers)
print("Values:", values)
sheet.update([headers] + values, range_name="A1")
