import pandas as pd
from datetime import datetime
import re
import requests
import os
import json
from dotenv import load_dotenv
from datetime import timedelta

load_dotenv()

START_DATE = "2025-04-21"
END_DATE = "2025-04-30"

def get_ctm_calls_for_number(phone_number, start_date_str, end_date_str):
    """Queries CTM API for calls matching a phone number within a date range."""
    api_key = os.getenv('CALL_TRACKING_METRICS_API_KEY')
    account_id = os.getenv('CTM_ACCOUNT_ID')
    if not api_key or not account_id:
        print("Error: CTM API Key or Account ID not found in environment variables.")
        return []

    headers = {"Authorization": f"Basic {api_key}"}
    url = f"https://api.calltrackingmetrics.com/api/v1/accounts/{account_id}/calls/search.json"
    
    params = {
        "search": phone_number,
        # "start_date": start_date_str,
        # "end_date": end_date_str,
    }
    
    try:
        response = requests.get(url, headers=headers, params=params)
        response.raise_for_status()
        data = response.json()
        return data.get('calls', [])
    except requests.exceptions.RequestException as e:
        print(f"Error querying CTM API for {phone_number}: {e}")
        return []
    except json.JSONDecodeError:
        print(f"Error decoding JSON response for {phone_number}. Response text: {response.text}")
        return []

# Read the CSV file
df = pd.read_csv("weekly_review_data/Leads-Reporting Export.csv")

# Filter for only "Form Fill" lead types
df = df[df['Lead Type'] == "Form Fill"]

# Convert "Date Added" to datetime format
df['Date Added'] = pd.to_datetime(df['Date Added'], format='%m/%d/%Y %I:%M%p', errors='coerce')

# Filter for dates between START_DATE and END_DATE
df = df[(df['Date Added'] >= START_DATE) & (df['Date Added'] <= END_DATE)]

# ITERATE OVER THER ORWS AND PRINT THE PHONE + DATE ADDED
for index, row in df.iterrows():
    normalized_phone = re.sub(r'\D', '', row['Phone'])
    lead_date_str = row['Date Added'].strftime('%Y-%m-%d')
    print(f"Lead: {normalized_phone} added on {lead_date_str}")

    # Search for CTM calls from lead date to END_DATE
    calls = get_ctm_calls_for_number(normalized_phone, lead_date_str, END_DATE)
    
    if calls:
        print(f"  Found {len(calls)} calls in CTM:")
        for call in calls:
            # Print relevant call details (adjust fields as needed)
            print(call)
            # save the call to a json file
            with open('call.json', 'w') as f:
                json.dump(call, f)
    else:
        print(f"  No calls found in CTM for {normalized_phone} between {lead_date_str} and {END_DATE}.")

print(len(df))