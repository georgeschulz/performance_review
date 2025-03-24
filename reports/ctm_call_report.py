import requests
import os
import json
from dotenv import load_dotenv
import pandas as pd
import datetime
from datetime import timedelta
import openpyxl
from openpyxl.styles import Font

load_dotenv()

def ctm_call_query(start_date, end_date, user_name):
    headers = {
        "Authorization": f"Basic {os.getenv('CALL_TRACKING_METRICS_API_KEY')}"
    }
    url = f"https://api.calltrackingmetrics.com/api/v1/accounts/{os.getenv('CTM_ACCOUNT_ID')}/reports/series.json"
    
    params = {
        "start_date": start_date,
        "end_date": end_date,
        "by": "agent",
    }
    response = requests.get(url, headers=headers, params=params)
    data = response.json()
    
    # Find the specific user and return their total calls
    for item in data.get("groups", {}).get("items", []):
        if item.get("name", {}).get("name") == user_name:
            return item.get("metrics", {}).get("total", {}).get("value", 0)
    
    # Return 0 if user not found
    return 0

def get_fiscal_year_start(current_date=None):
    """
    Calculate the start date of the current fiscal year (September 1st)
    """
    if current_date is None:
        current_date = datetime.datetime.now()
    
    # If current date is before September 1st, fiscal year started last year
    if current_date.month < 9:
        return datetime.datetime(current_date.year - 1, 9, 1)
    # Otherwise fiscal year started this year
    else:
        return datetime.datetime(current_date.year, 9, 1)

def get_month_start(current_date=None):
    """
    Calculate the start date of the current month
    """
    if current_date is None:
        current_date = datetime.datetime.now()
    
    return datetime.datetime(current_date.year, current_date.month, 1)

def ctm_call_report(week_start=None, agents=None, exclude_call_statuses=[], replacements={}, current_date=None):
    """
    Generate call report for a specific week using Call Tracking Metrics API
    
    Parameters:
    - week_start: datetime or string, start date of the week (Monday)
    - agents: list of agents to include
    - exclude_call_statuses: list of call statuses to exclude (not used with CTM API)
    - replacements: dict of agent name replacements
    - current_date: datetime, date to use as "now" (defaults to today)
    
    Returns a dict with weekly, mtd, and ytd metrics for the specified week
    """
    # Ensure output directory exists
    os.makedirs('weekly_outputs', exist_ok=True)
    
    # Set current date if not provided
    if current_date is None:
        current_date = datetime.datetime.now()
    elif isinstance(current_date, str):
        current_date = pd.to_datetime(current_date)
    
    # If week_start not provided, get the Monday of the current week
    if week_start is None:
        days_since_monday = current_date.weekday()
        week_start = current_date - timedelta(days=days_since_monday)
    elif isinstance(week_start, str):
        week_start = pd.to_datetime(week_start)
    
    # Calculate week end (Sunday)
    week_end = week_start + timedelta(days=6)
    
    # Calculate fiscal year and month start dates
    fiscal_year_start = get_fiscal_year_start(week_end)
    month_start = get_month_start(week_end)
    
    # Use all agents if none provided
    if agents is None:
        # You may need to fetch the list of agents from CTM API
        # For now, use an empty list that will need to be populated
        agents = []
    
    # Initialize results
    results = []
    
    # Get metrics for each agent
    for agent in agents:
        # Get weekly data
        weekly_calls = ctm_call_query(
            week_start.strftime('%Y-%m-%d'),
            week_end.strftime('%Y-%m-%d'),
            agent
        )
        
        # Get month-to-date data
        mtd_calls = ctm_call_query(
            month_start.strftime('%Y-%m-%d'),
            week_end.strftime('%Y-%m-%d'),
            agent
        )
        
        # Get year-to-date data
        ytd_calls = ctm_call_query(
            fiscal_year_start.strftime('%Y-%m-%d'),
            week_end.strftime('%Y-%m-%d'),
            agent
        )
        
        # Add to results
        results.append({
            'Agent': agent,
            'Week Start': week_start.strftime('%Y-%m-%d'),
            'Week End': week_end.strftime('%Y-%m-%d'),
            # Weekly metrics
            'Weekly Total Calls': weekly_calls,
            # MTD metrics
            'MTD Total Calls': mtd_calls,
            # YTD metrics
            'YTD Total Calls': ytd_calls
        })
    
    # Convert results to DataFrame
    results_df = pd.DataFrame(results)
    
    # Save as Excel
    excel_path = f'weekly_outputs/CTM_Calls_Report_{week_start.strftime("%Y-%m-%d")}.xlsx'
    results_df.to_excel(excel_path, index=False, engine='openpyxl')
    
    # Apply formatting to the Excel file
    workbook = openpyxl.load_workbook(excel_path)
    worksheet = workbook.active
    
    # Adjust column widths
    for column in worksheet.columns:
        max_length = 0
        column_letter = column[0].column_letter
        
        # Find the maximum length in the column
        for cell in column:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        
        # Set the column width (with some padding)
        adjusted_width = max_length + 4
        worksheet.column_dimensions[column_letter].width = adjusted_width
    
    # Save the formatted workbook
    workbook.save(excel_path)
    
    return {
        'weekly': results_df[['Agent', 'Week Start', 'Week End', 'Weekly Total Calls']],
        'mtd': results_df[['Agent', 'MTD Total Calls']],
        'ytd': results_df[['Agent', 'YTD Total Calls']],
        'full_report': results_df
    }

if __name__ == "__main__":
    # Example: generate report for a specific week
    week_start = datetime.datetime(2025, 3, 17)  # March 17, 2025 (a Monday)
    agents = ["Kamaal Sherrod", "Hussam Olabi"]
    ctm_call_report(week_start=week_start, agents=agents)
