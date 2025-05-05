import pandas as pd
from datetime import datetime, timedelta, time # Removed time
import re
import requests
import os
import json
from dotenv import load_dotenv
import pytz  # Import pytz for timezone handling # Removed

load_dotenv()

START_DATE = "2025-04-21"
END_DATE = "2025-05-03" 
# END_DATE = "2025-04-21" 

# --- Business Hours Constants ---
BUSINESS_START_TIME = time(8, 0)
BUSINESS_END_TIME = time(17, 0)
BUSINESS_TIMEZONE = pytz.timezone('US/Eastern')
WEEKDAYS = set(range(5)) # Monday=0 to Friday=4
lead_types_to_include = ["Form Fill", "Thumbtack", "WTR Free Trial Request", "Email Lead"]


def is_during_business_hours(dt_aware):
    """Checks if a timezone-aware datetime falls within Mon-Fri 8:00-17:00 ET."""
    if pd.isna(dt_aware):
        return 'Unknown' # Or handle NA dates as needed

    # Convert to business timezone
    dt_local = dt_aware.astimezone(BUSINESS_TIMEZONE)

    is_weekday = dt_local.weekday() in WEEKDAYS
    is_during_hours = BUSINESS_START_TIME <= dt_local.time() < BUSINESS_END_TIME

    if is_weekday and is_during_hours:
        return 'In Hours'
    else:
        return 'Out of Hours'

def get_ctm_calls_for_number(phone_number, start_date_str, end_date_str):
    """Queries CTM API for calls matching a phone number. Note: CTM date filtering is currently commented out."""
    api_key = os.getenv('CALL_TRACKING_METRICS_API_KEY')
    account_id = os.getenv('CTM_ACCOUNT_ID')
    if not api_key or not account_id:
        print("Error: CTM API Key or Account ID not found in environment variables.")
        return []

    headers = {"Authorization": f"Basic {api_key}"}
    url = f"https://api.calltrackingmetrics.com/api/v1/accounts/{account_id}/calls/search.json"

    params = {
        "search": phone_number,
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
try:
    df = pd.read_csv("weekly_review_data/Leads-Reporting Export.csv")
except FileNotFoundError:
    print("Error: weekly_review_data/Leads-Reporting Export.csv not found.")
    exit()

# ---- Data Preparation and Filtering ----

# Filter for "Form Fill" or "Thumbtack" lead types
df = df[df['Lead Type'].isin(lead_types_to_include)].copy() # Use .copy() to avoid SettingWithCopyWarning

# Filter out rows where "Close Status" contains "Disqualified"
# Using na=False to treat NaN values as not containing the substring
df = df[~df['Close Status'].str.contains("Disqualified", na=False)].copy()

# Convert "Date Added" to datetime format, coercing errors
df['Date Added'] = pd.to_datetime(df['Date Added'], format='%m/%d/%Y %I:%M%p', errors='coerce')

# Drop rows where 'Date Added' conversion failed
df.dropna(subset=['Date Added'], inplace=True)

# Make 'Date Added' timezone-aware (assuming UTC if no timezone info, adjust if needed)
# Or localize to a specific timezone if the source data has one implicitly.
# Example: df['Date Added'] = df['Date Added'].dt.tz_localize('America/New_York')
# If 'Date Added' might already be timezone-aware from pd.to_datetime, check df['Date Added'].dt.tz
if df['Date Added'].dt.tz is None:
     df['Date Added'] = df['Date Added'].dt.tz_localize(BUSINESS_TIMEZONE) # Use Business Timezone


# Filter for dates between START_DATE and END_DATE (inclusive)
start_dt = pd.to_datetime(START_DATE).tz_localize(BUSINESS_TIMEZONE) # Use Business Timezone
end_dt = pd.to_datetime(END_DATE).tz_localize(BUSINESS_TIMEZONE) + timedelta(days=1) # Use Business Timezone, Make end date inclusive

df = df[(df['Date Added'] >= start_dt) & (df['Date Added'] < end_dt)] # Use '< end_dt' for inclusivity up to EOD


# ---- Process Leads and Find First Touch Call ----
report_data = []

for index, row in df.iterrows():
    normalized_phone = re.sub(r'\D', '', str(row.get('Phone', ''))) # Ensure Phone is string
    lead_date = row['Date Added'] # Already timezone-aware

    # Assuming 'First Name' and 'Salesperson' columns exist in the CSV
    customer_first_name = row.get('First Name', '') # Use .get for safety
    salesperson = row.get('Salesperson', '') # Use .get for safety

    print(f"Processing Lead: {normalized_phone} added on {lead_date.strftime('%Y-%m-%d %H:%M:%S %Z')}")

    # Search for CTM calls using only the phone number
    # Date filtering happens after fetching all calls for the number
    calls = get_ctm_calls_for_number(normalized_phone, None, None) # Pass None for dates

    first_call_agent = None
    first_call_datetime = None
    total_elapsed_minutes = pd.NA
    # total_accountable_minutes = pd.NA # Removed

    if calls:
        print(f"  Found {len(calls)} raw calls in CTM for {normalized_phone}.")
        valid_calls = []
        for call in calls:
            try:
                # Parse 'called_at' string - it includes offset info
                call_time_str = call.get('called_at')
                if call_time_str:
                     # Use pandas to_datetime which handles various formats including offsets
                    call_datetime = pd.to_datetime(call_time_str)
                    # Ensure it's timezone-aware (should be based on format)
                    if call_datetime.tzinfo is None:
                         # This case is unlikely if parsing succeeds with offset, but handle just in case
                         call_datetime = call_datetime.tz_localize('UTC') # Or handle as error/log

                    # Compare with lead date (both must be timezone-aware)
                    if call_datetime >= (lead_date - timedelta(minutes=30)):
                        valid_calls.append({
                            'agent': call.get('agent', {}).get('name'),
                            'datetime': call_datetime
                        })
            except Exception as e:
                print(f"  Warning: Could not parse datetime '{call_time_str}' or process call {call.get('id')}: {e}")

        if valid_calls:
            # Sort calls by datetime
            valid_calls.sort(key=lambda x: x['datetime'])
            # Get the earliest call after the lead date
            first_call = valid_calls[0]
            first_call_agent = first_call['agent']
            first_call_datetime = first_call['datetime']
            print(f"  Found first touch call at {first_call_datetime.strftime('%Y-%m-%d %H:%M:%S %Z')} by {first_call_agent}")

            # --- Calculate Elapsed Minutes ---
            time_diff_seconds = (first_call_datetime - lead_date).total_seconds()
            if time_diff_seconds >= 0:
                total_elapsed_minutes = time_diff_seconds / 60
            else:
                print(f"  Warning: First call datetime {first_call_datetime} is before lead date {lead_date}. Setting elapsed minutes to 0.")
                total_elapsed_minutes = 0 # Set to 0 minutes if negative

        else:
             print(f"  No valid calls found on or after lead date {lead_date.strftime('%Y-%m-%d %H:%M:%S %Z')}.")

    else:
        print(f"  No calls found in CTM for {normalized_phone}.")

    # Determine Business Hours Status
    business_hours_status = is_during_business_hours(lead_date)

    report_data.append({
        "Customer Full Name": row['Customer Full Name'],
        "Phone": row['Phone'],
        "Lead Type": row['Lead Type'],
        "Salesperson": salesperson,
        "Business Hours Status": business_hours_status,
        "First Call Agent": first_call_agent,
        "Date Added": row['Date Added'],
        "First Touch": first_call_datetime, # Renamed from "First Call Datetime"
        "Total Elapsed Minutes": total_elapsed_minutes,
        # "Total Accountable Minutes": total_accountable_minutes # Removed
    })


# ---- Create and Save Report ----
if report_data:
    report_df = pd.DataFrame(report_data)

    # --- Sort by Business Hours Status and then Date Added ---
    report_df = report_df.sort_values(by=["Business Hours Status", "Date Added"], ascending=[True, True])

    # Ensure the output directory exists
    output_dir = 'weekly_outputs'
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, 'first_touch.xlsx')

    # Format datetime columns for CSV output
    dt_format = "%A %I:%M %p" # e.g., Monday 08:30 AM
    report_df['Date Added'] = report_df['Date Added'].dt.strftime(dt_format)
    # Handle potential NaT values in 'First Touch' before formatting
    report_df['First Touch'] = pd.to_datetime(report_df['First Touch']).dt.strftime(dt_format)

    # Round minutes for better readability in CSV
    if "Total Elapsed Minutes" in report_df.columns:
           report_df["Total Elapsed Minutes"] = report_df["Total Elapsed Minutes"].round(2)
    # if "Total Accountable Minutes" in report_df.columns: # Removed
    #        report_df["Total Accountable Minutes"] = report_df["Total Accountable Minutes"].round(2) # Removed

    # report_df.to_csv(output_path, index=False) # Commented out CSV saving

    # --- Write to Excel with Formatting ---
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet('First Touch Report') # Custom sheet name
        writer.sheets['First Touch Report'] = worksheet # Point writer to the worksheet

        # Define formats
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': False,
            'valign': 'top',
            'fg_color': '#D7E4BC', # Light green header
            'border': 1
        })
        subheader_format = workbook.add_format({
            'bold': True,
            'fg_color': '#DDEBF7', # Light blue sub-header
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        date_format_excel = workbook.add_format({'num_format': 'dddd hh:mm AM/PM', 'align': 'left'})
        default_format = workbook.add_format({'align': 'left'})

        # Write Headers
        col_names = report_df.columns.tolist()
        for col_num, value in enumerate(col_names):
            worksheet.write(0, col_num, value, header_format)

        # --- Write Data Grouped by Business Hours Status ---
        current_row = 1 # Start writing data from row 1 (below headers)
        # Ensure data is sorted correctly for grouping display
        report_df = report_df.sort_values(by=["Business Hours Status", "Date Added"], ascending=[True, True])

        for status, group_df in report_df.groupby("Business Hours Status"):
            # Write the sub-header row spanning columns
            worksheet.merge_range(current_row, 0, current_row, len(col_names) - 1, f"{status} Leads", subheader_format)
            worksheet.set_row(current_row, 20) # Set height for sub-header row
            current_row += 1

            # Write data rows for this group
            for index, data_row in group_df.iterrows():
                for col_num, col_name in enumerate(col_names):
                    cell_value = data_row[col_name]
                    cell_format = default_format
                    # Apply specific formats
                    if isinstance(cell_value, (datetime, pd.Timestamp)):
                        # Convert timezone-aware to naive for xlsxwriter if needed, or handle directly
                        if cell_value.tzinfo is not None:
                             # Excel doesn't handle tz well, write as naive local
                             cell_value = cell_value.tz_convert(BUSINESS_TIMEZONE).tz_localize(None)
                        cell_format = date_format_excel
                    elif pd.isna(cell_value):
                         cell_value = '' # Write empty string for NaN/NaT

                    worksheet.write(current_row, col_num, cell_value, cell_format)
                current_row += 1
            current_row += 1 # Add a blank row between groups

        # --- Adjust Column Widths ---
        for col_num, col_name in enumerate(col_names):
            # Find max length in the column data
            try:
                 # Handle potential NaNs and convert others to string to measure length
                max_len = max(group_df[col_name].astype(str).apply(len).max() for _, group_df in report_df.groupby("Business Hours Status"))
                # Also consider header length
                header_len = len(col_name)
                # Consider formatted date length (approximate)
                date_len_approx = 25 if "Date" in col_name or "Touch" in col_name else 0
                # Add some padding
                column_width = max(max_len, header_len, date_len_approx) + 2
            except (KeyError, ValueError, TypeError): # Handle empty groups, non-string data, or TypeError if max() receives an empty sequence
                column_width = len(col_name) + 2 # Default to header width + padding

            # Clamp width to avoid excessively wide columns (e.g., max 50)
            column_width = min(column_width, 50)
            worksheet.set_column(col_num, col_num, column_width)

        # --- Apply Conditional Formatting to 'Total Elapsed Minutes' ---
        try:
            elapsed_minutes_col_index = col_names.index("Total Elapsed Minutes")
            # Column letter (e.g., 'A', 'B', ...) based on index
            elapsed_minutes_col_letter = chr(ord('A') + elapsed_minutes_col_index)
            # Define the range: From row 2 down to the last data row written
            # current_row represents the *next* row to write, so the last written row is current_row - 1
            # We need to account for header (row 0) and sub-header rows. The exact last row might vary slightly.
            # Apply formatting to the entire column range where data *could* be.
            # Max row in Excel is 1048576, but let's use a reasonable upper limit like current_row + len(report_df)
            # or simply use the calculated current_row as the end point.
            # Let's apply from row 2 (first data row potential start) to the last row written (current_row -1)
            # The range needs to skip the header and sub-header rows.
            # It's simpler to apply it to the whole column range and let it affect only numeric cells.
            formatting_range = f'{elapsed_minutes_col_letter}2:{elapsed_minutes_col_letter}{current_row}' # Apply from row 2 downwards

            # Define formats for conditional formatting
            format_green = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'}) # Green fill, dark green text
            format_yellow = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500'}) # Yellow fill, dark yellow text
            format_red = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'}) # Red fill, dark red text
            format_dark_red = workbook.add_format({'bg_color': '#9C0006', 'font_color': '#FFFFFF'}) # Dark red fill, white text

            # Apply the rules (Order matters: Stop if blank rule first)

            # Rule 0: Stop if blank (Apply no format)
            worksheet.conditional_format(formatting_range, {'type': 'blanks',
                                                            'stop_if_true': True,
                                                            'format': None})

            # Rule 1: < 3 (Green)
            worksheet.conditional_format(formatting_range, {'type': 'cell',
                                                            'criteria': 'less than',
                                                            'value': 3,
                                                            'format': format_green})

            # Rule 2: >= 3 and <= 5 (Yellow)
            worksheet.conditional_format(formatting_range, {'type': 'cell',
                                                            'criteria': 'between',
                                                            'minimum': 3,
                                                            'maximum': 5,
                                                            'format': format_yellow})

            # Rule 3: > 5 and <= 60 (Red)
            worksheet.conditional_format(formatting_range, {'type': 'cell',
                                                            'criteria': 'between',
                                                            'minimum': 5.00001, # Slightly above 5 to avoid overlap
                                                            'maximum': 60,
                                                            'format': format_red})

            # Rule 4: > 60 (Dark Red)
            worksheet.conditional_format(formatting_range, {'type': 'cell',
                                                            'criteria': 'greater than',
                                                            'value': 60,
                                                            'format': format_dark_red})
        except ValueError:
            print("Warning: 'Total Elapsed Minutes' column not found. Skipping conditional formatting.")
        except IndexError:
             print("Warning: Could not determine column letter for conditional formatting.")

    print(f"\nReport saved to {output_path}")
    print(f"Total leads processed: {len(df)}")
    print(f"Total records in report: {len(report_df)}")
else:
    print("\nNo data processed or no leads found in the specified date range and type.")

# Remove the old print(len(df))
# print(len(df))