import pandas as pd
from datetime import datetime, timedelta, time # Removed time
import re
import requests
import os
import json
from dotenv import load_dotenv
import pytz  # Import pytz for timezone handling # Removed
import numpy as np # Add numpy import

load_dotenv()

# --- NEW: Phone Formatting Function ---
def format_phone_number(phone_str):
    """Formats a string containing digits into xxx-xxx-xxxx."""
    digits = re.sub(r'\D', '', str(phone_str)) # Ensure input is string and get only digits
    if len(digits) == 10:
        return f"{digits[0:3]}-{digits[3:6]}-{digits[6:10]}"
    elif len(digits) == 11 and digits.startswith('1'): # Handle optional leading '1'
        return f"{digits[1:4]}-{digits[4:7]}-{digits[7:11]}"
    else:
        # Return original string if format is unexpected, or just digits if preferred
        return phone_str # Or return digits
# --- END NEW ---

START_DATE = "2025-05-05"
# START_DATE = "2025-04-26"
END_DATE = "2025-05-11" 
# END_DATE = "2025-04-21" 

# --- Business Hours Constants ---
BUSINESS_START_TIME = time(8, 0)
BUSINESS_END_TIME = time(17, 0)
BUSINESS_TIMEZONE = pytz.timezone('US/Eastern')
WEEKDAYS = set(range(5)) # Monday=0 to Friday=4
lead_types_to_include = ["Form Fill", "Thumbtack", "WTR Free Trial Request", "Email Lead"]


def calculate_business_minutes_between(start_dt_aware, end_dt_aware):
    """
    Calculates the total number of minutes that fall within business hours
    (Mon-Fri, 8:00-17:00 ET) between two timezone-aware datetimes.
    """
    if pd.isna(start_dt_aware) or pd.isna(end_dt_aware) or end_dt_aware <= start_dt_aware:
        return 0 # Or pd.NA based on how you want to handle invalid inputs

    total_accountable_seconds = 0
    # Ensure calculations are done in the business timezone
    start_local = start_dt_aware.astimezone(BUSINESS_TIMEZONE)
    end_local = end_dt_aware.astimezone(BUSINESS_TIMEZONE)

    current_date = start_local.date()
    while current_date <= end_local.date():
        if current_date.weekday() in WEEKDAYS:
            # Define business hours for the current day in the local timezone
            try:
                bh_start = BUSINESS_TIMEZONE.localize(datetime.combine(current_date, BUSINESS_START_TIME))
                bh_end = BUSINESS_TIMEZONE.localize(datetime.combine(current_date, BUSINESS_END_TIME))
            except pytz.exceptions.AmbiguousTimeError:
                 # Handle DST transition if start/end falls on ambiguous time
                 bh_start = BUSINESS_TIMEZONE.localize(datetime.combine(current_date, BUSINESS_START_TIME), is_dst=True) # or False depending on context
                 bh_end = BUSINESS_TIMEZONE.localize(datetime.combine(current_date, BUSINESS_END_TIME), is_dst=True) # or False
            except pytz.exceptions.NonExistentTimeError:
                 # Handle DST transition if start/end falls on non-existent time (e.g., skip an hour forward)
                 # Adjust time or handle as needed, e.g., move to next valid time
                 bh_start = BUSINESS_TIMEZONE.normalize(BUSINESS_TIMEZONE.localize(datetime.combine(current_date, BUSINESS_START_TIME) + timedelta(hours=1), is_dst=None))
                 bh_end = BUSINESS_TIMEZONE.normalize(BUSINESS_TIMEZONE.localize(datetime.combine(current_date, BUSINESS_END_TIME) + timedelta(hours=1), is_dst=None))


            # Calculate the overlap between the overall interval [start_local, end_local]
            # and the business hours interval for the current day [bh_start, bh_end]
            overlap_start = max(start_local, bh_start)
            overlap_end = min(end_local, bh_end)

            # Add the duration of the overlap if it's positive
            if overlap_end > overlap_start:
                total_accountable_seconds += (overlap_end - overlap_start).total_seconds()

        # Move to the next day
        current_date += timedelta(days=1)

    return total_accountable_seconds / 60


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
    total_accountable_minutes = pd.NA # Initialize new column value

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
             # If no first call, elapsed and accountable minutes remain NA
             pass # No need to calculate accountable if no first call

    else:
        print(f"  No calls found in CTM for {normalized_phone}.")
        # If no calls found, elapsed and accountable minutes remain NA
        pass

    # Determine Business Hours Status for the lead submission time
    business_hours_status = is_during_business_hours(lead_date)

    # Calculate Total Accountable Minutes
    if pd.notna(first_call_datetime) and pd.notna(lead_date) and total_elapsed_minutes >= 0:
        if business_hours_status == 'In Hours':
            total_accountable_minutes = total_elapsed_minutes
        elif business_hours_status == 'Out of Hours':
            # Calculate minutes within business hours between lead and first call
            total_accountable_minutes = calculate_business_minutes_between(lead_date, first_call_datetime)
        else: # Handle 'Unknown' status if necessary
             total_accountable_minutes = pd.NA # Or some other default

    report_data.append({
        "Customer Full Name": row['Customer Full Name'],
        "Phone": row['Phone'],
        "Lead Type": row['Lead Type'],
        "Salesperson": salesperson,
        "Business Hours Status": business_hours_status,
        "First Call Agent": first_call_agent,
        "Date Submitted": row['Date Added'], # Keep original datetime for sorting/formatting later
        "Date Added": row['Date Added'], # Keep original datetime for sorting/formatting later
        "First Touch": first_call_datetime, # Renamed from "First Call Datetime"
        "Total Elapsed Minutes": total_elapsed_minutes,
        "Total Accountable Minutes": total_accountable_minutes
    })


# ---- Create and Save Report ----
if report_data:
    report_df = pd.DataFrame(report_data)

    # --- Sort by Business Hours Status and then Date Added ---
    # Sort by Date Submitted first to keep chronological order within groups
    report_df = report_df.sort_values(by=["Business Hours Status", "Date Submitted"], ascending=[True, True])

    # Ensure the output directory exists
    output_dir = 'weekly_outputs'
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, 'first_touch.xlsx')

    # Keep original datetime columns for Excel formatting
    report_df_excel = report_df.copy()

    # --- Apply Phone Formatting for Excel Output ---
    report_df_excel['Phone'] = report_df_excel['Phone'].apply(format_phone_number)
    # --- End Apply Phone Formatting ---

    # Format datetime columns for display (will be overridden by Excel formats)
    dt_format_day_time = "%A %I:%M %p" # e.g., Monday 08:30 AM
    dt_format_date_only = "%m/%d/%y" # e.g., 04/21/25

    report_df['Date Submitted'] = pd.to_datetime(report_df['Date Submitted']).dt.strftime(dt_format_date_only)
    report_df['Date Added'] = pd.to_datetime(report_df['Date Added']).dt.strftime(dt_format_day_time)
    # Handle potential NaT values in 'First Touch' before formatting
    report_df['First Touch'] = pd.to_datetime(report_df['First Touch']).dt.strftime(dt_format_day_time)

    # Round minutes for better readability
    if "Total Elapsed Minutes" in report_df.columns:
           report_df["Total Elapsed Minutes"] = report_df["Total Elapsed Minutes"].round(2)
    if "Total Accountable Minutes" in report_df.columns:
           report_df["Total Accountable Minutes"] = report_df["Total Accountable Minutes"].round(2)

    # report_df.to_csv(output_path, index=False) # Commented out CSV saving

    # --- Write to Excel with Formatting ---
    # Use report_df_excel which has the original datetime objects for Excel formatting
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
        date_format_excel_day_time = workbook.add_format({'num_format': 'dddd hh:mm AM/PM', 'align': 'left'})
        date_format_excel_date_only = workbook.add_format({'num_format': 'mm/dd/yy', 'align': 'left'})
        default_format = workbook.add_format({'align': 'left'})
        # --- New formats for Average row ---
        average_label_format = workbook.add_format({
            'bold': True,
            'fg_color': '#F2F2F2', # Light grey background
            'align': 'right',
            'valign': 'vcenter',
            'border': 1
        })
        average_value_format = workbook.add_format({
            'bold': True,
            'num_format': '0.0', # One decimal place
            'fg_color': '#F2F2F2',
            'align': 'left',
            'valign': 'vcenter',
            'border': 1
        })
        average_na_format = workbook.add_format({ # Format for N/A case
            'bold': True,
            'fg_color': '#F2F2F2',
            'align': 'left',
            'valign': 'vcenter',
            'border': 1
        })
        # --- New formats for Percentage row ---
        percentage_label_format = workbook.add_format({
            'bold': True,
            'fg_color': '#F2F2F2', # Light grey background
            'align': 'right',
            'valign': 'vcenter',
            'border': 1
        })
        percentage_value_format = workbook.add_format({
            'bold': True,
            'num_format': '0.0%', # Percentage format
            'fg_color': '#F2F2F2',
            'align': 'left',
            'valign': 'vcenter',
            'border': 1
        })
        percentage_na_format = workbook.add_format({ # Format for N/A percentage
            'bold': True,
            'fg_color': '#F2F2F2',
            'align': 'left',
            'valign': 'vcenter',
            'border': 1
        })

        # Write Headers
        col_names = report_df_excel.columns.tolist() # Get columns from the df with original data types
        for col_num, value in enumerate(col_names):
            worksheet.write(0, col_num, value, header_format)

        # --- Write Data Grouped by Business Hours Status ---
        current_row = 1 # Start writing data from row 1 (below headers)
        total_below_5_min_count = 0 # Initialize total counter for < 5 min
        total_called_count = 0      # Initialize total counter for leads with a call
        # group_below_5_min_counts = {} # Dictionary to store counts per group - No longer needed directly

        # Ensure data is sorted correctly for grouping display
        report_df_excel = report_df_excel.sort_values(by=["Business Hours Status", "Date Submitted"], ascending=[True, True])

        for status, group_df in report_df_excel.groupby("Business Hours Status"):
            # Write the sub-header row spanning columns
            worksheet.merge_range(current_row, 0, current_row, len(col_names) - 1, f"{status} Leads", subheader_format)
            worksheet.set_row(current_row, 20) # Set height for sub-header row
            current_row += 1

            # Write data rows for this group
            for index, data_row in group_df.iterrows():
                for col_num, col_name in enumerate(col_names):
                    cell_value = data_row[col_name]
                    cell_format = default_format # Start with default format

                    # Handle NaT/NaN first
                    if pd.isna(cell_value):
                        cell_value = '' # Write empty string for NaN/NaT
                        # Keep default_format for NaT/NaN values
                    # Check for datetime type only if value is not NA
                    elif isinstance(cell_value, (datetime, pd.Timestamp)):
                        # Convert timezone-aware to naive for xlsxwriter if needed
                        if cell_value.tzinfo is not None:
                             # Excel doesn't handle tz well, write as naive local
                             cell_value = cell_value.tz_convert(BUSINESS_TIMEZONE).tz_localize(None)

                        # Choose format based on column name
                        if col_name == "Date Submitted":
                            cell_format = date_format_excel_date_only
                        elif col_name == "Date Added" or col_name == "First Touch":
                            cell_format = date_format_excel_day_time
                        # else: # If other datetime columns exist, they use default or need specific formats
                        #    pass # Keep default_format or add other conditions

                    # For non-NA, non-datetime values, default_format is already set

                    # elif pd.isna(cell_value): # This check is now redundant due to the check at the beginning
                    #      cell_value = '' # Write empty string for NaN/NaT

                    worksheet.write(current_row, col_num, cell_value, cell_format)
                current_row += 1

            # --- Calculate and Write Average Row ---
            if not group_df.empty:
                # Calculate average, coercing errors and dropping NA
                valid_minutes = pd.to_numeric(group_df['Total Accountable Minutes'], errors='coerce').dropna()
                average_minutes = valid_minutes.mean() if not valid_minutes.empty else None

                # Find the column index for 'Total Accountable Minutes'
                try:
                    avg_col_index = col_names.index('Total Accountable Minutes')
                except ValueError:
                    avg_col_index = -1 # Should not happen if column exists

                # Write the label spanning first few columns (e.g., first 4)
                merge_end_col = min(3, len(col_names) - 2) # Ensure merge doesn't go beyond penultimate column
                if merge_end_col >= 0:
                    worksheet.merge_range(current_row, 0, current_row, merge_end_col, f"Average Accountable Minutes ({status}):", average_label_format)
                    # Fill blank cells in merged range for border consistency if needed (optional)
                    for i in range(1, merge_end_col + 1):
                         worksheet.write_blank(current_row, i, '', average_label_format)

                # Write the average value or N/A
                if avg_col_index != -1:
                    if average_minutes is not None:
                        worksheet.write(current_row, avg_col_index, round(average_minutes, 1), average_value_format)
                    else:
                        worksheet.write(current_row, avg_col_index, 'N/A', average_na_format)
                # Optionally fill other cells in the average row with the grey background/border
                for i in range(len(col_names)):
                    if i <= merge_end_col:
                        continue # Skip cells covered by merge
                    if i != avg_col_index:
                        worksheet.write_blank(current_row, i, '', average_na_format) # Use na_format for consistent bg/border

                worksheet.set_row(current_row, 20) # Set height for average row
                current_row += 1
            # --- End of Average Row section ---

            # --- Calculate and Write Percentage Row for < 5 Minutes ---
            if not group_df.empty:
                 # Get numeric series, dropping NA (represents leads with a calculated time)
                 valid_minutes_series = pd.to_numeric(group_df['Total Accountable Minutes'], errors='coerce').dropna()
                 group_called_count = len(valid_minutes_series) # Count of leads with a call in this group
                 group_below_5_min_count = (valid_minutes_series < 5).sum() # Count < 5 min within those called

                 # Update total counts
                 total_called_count += group_called_count
                 total_below_5_min_count += group_below_5_min_count

                 # Calculate percentage
                 if group_called_count > 0:
                     percentage_below_5 = group_below_5_min_count / group_called_count
                     display_value = percentage_below_5
                     display_format = percentage_value_format
                 else:
                     display_value = "N/A"
                     display_format = percentage_na_format

                 # Reuse merge_end_col and avg_col_index from average calculation
                 if merge_end_col >= 0:
                     worksheet.merge_range(current_row, 0, current_row, merge_end_col, f"% Called < 5 Min ({status}):", percentage_label_format)
                     # Fill blank cells in merged range
                     for i in range(1, merge_end_col + 1):
                          worksheet.write_blank(current_row, i, '', percentage_label_format)

                 if avg_col_index != -1: # Assuming percentage displayed in same column as average
                     worksheet.write(current_row, avg_col_index, display_value, display_format)

                 # Fill other cells in the percentage row
                 for i in range(len(col_names)):
                     if i <= merge_end_col: continue
                     if i != avg_col_index:
                         worksheet.write_blank(current_row, i, '', percentage_label_format) # Use label format for consistency

                 worksheet.set_row(current_row, 20) # Set height for percentage row
                 current_row += 1
            # --- End of Percentage Row section ---

            # Add a blank row between groups
            worksheet.write_blank(current_row, 0, '', None) # Ensure truly blank row if needed
            current_row += 1

        # --- Write Total Percentage Row at the end ---
        if report_df_excel.empty: # Handle case where there's no data at all
            worksheet.write(current_row, 0, "No data to process.", default_format)
            current_row +=1
        else:
            # Find column indices again (could be stored from before)
            try:
                avg_col_index = col_names.index('Total Accountable Minutes')
            except ValueError:
                avg_col_index = -1
            merge_end_col = min(3, len(col_names) - 2)

            # Calculate total percentage
            if total_called_count > 0:
                 total_percentage_below_5 = total_below_5_min_count / total_called_count
                 total_display_value = total_percentage_below_5
                 total_display_format = percentage_value_format
            else:
                 total_display_value = "N/A"
                 total_display_format = percentage_na_format

            # Write total percentage label
            if merge_end_col >= 0:
                worksheet.merge_range(current_row, 0, current_row, merge_end_col, "Total % Called < 5 Min:", percentage_label_format)
                for i in range(1, merge_end_col + 1):
                     worksheet.write_blank(current_row, i, '', percentage_label_format)

            # Write total percentage value
            if avg_col_index != -1:
                worksheet.write(current_row, avg_col_index, total_display_value, total_display_format)

            # Fill other cells in total percentage row
            for i in range(len(col_names)):
                 if i <= merge_end_col: continue
                 if i != avg_col_index:
                     worksheet.write_blank(current_row, i, '', percentage_label_format)

            worksheet.set_row(current_row, 20) # Set height for total percentage row
            current_row += 1
        # --- End of Total Percentage Row section ---

        # --- Adjust Column Widths ---
        for col_num, col_name in enumerate(col_names):
            # Find max length in the column data (use the formatted string version for length calculation)
            try:
                # Use the already formatted report_df for string length calculation
                # Handle potential NaN values when calculating max length
                str_series = report_df.iloc[group_df.index][col_name].astype(str)
                # Replace 'nan', 'NaT' etc. representation with empty string for length calculation if desired
                # str_series = str_series.replace(['nan', 'NaT'], '', regex=False)
                max_len = max(str_series.apply(len).max() for _, group_df in report_df_excel.groupby("Business Hours Status") if not group_df.empty)

                # Also consider header length
                header_len = len(col_name)

                # Consider formatted date length (approximate) - adjusted for new date formats
                date_len_approx = 0
                if col_name == "Date Submitted":
                    date_len_approx = 10 # "mm/dd/yy"
                elif col_name == "Date Added" or col_name == "First Touch":
                    date_len_approx = 25 # "dddd hh:mm AM/PM"

                # Add some padding
                column_width = max(max_len, header_len, date_len_approx) + 2
            except (KeyError, ValueError, TypeError): # Handle empty groups, non-string data, or TypeError if max() receives an empty sequence
                column_width = len(col_name) + 2 # Default to header width + padding

            # Clamp width to avoid excessively wide columns (e.g., max 50)
            column_width = min(column_width, 50)
            worksheet.set_column(col_num, col_num, column_width)

        # --- Apply Conditional Formatting to 'Total Accountable Minutes' ---
        try:
            # Find the index of the target column
            target_col_name = "Total Accountable Minutes" # Changed from "Total Elapsed Minutes"
            target_col_index = col_names.index(target_col_name)
            # Column letter (e.g., 'A', 'B', ...) based on index
            target_col_letter = chr(ord('A') + target_col_index)

            # Define the range: Apply from row 2 downwards in the target column
            formatting_range = f'{target_col_letter}2:{target_col_letter}{current_row}'

            # Define formats for conditional formatting (reuse existing formats)
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
            print("Warning: 'Total Accountable Minutes' column not found. Skipping conditional formatting.")
        except IndexError:
             print("Warning: Could not determine column letter for conditional formatting.")

    print(f"\nReport saved to {output_path}")
    print(f"Total leads processed: {len(df)}")
    print(f"Total records in report: {len(report_df)}")
else:
    print("\nNo data processed or no leads found in the specified date range and type.")

# Remove the old print(len(df))
# print(len(df))