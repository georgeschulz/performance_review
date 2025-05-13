import pandas as pd
import requests
import os
from dotenv import load_dotenv

load_dotenv()

def count_salesperson_comments(record_id, salespeople_to_check=["Kamaal Sherrod", "Hussam Olabi"]):
    """
    Counts the number of comments made by specified salespeople for a given record.
    
    Args:
        record_id (str): The Airtable record ID to check for comments
        salespeople_to_check (list): List of salesperson names whose comments to count.
        
    Returns:
        int: The number of comments made by the specified salespeople
    """
    # Airtable API configuration
    base_id = "appBsty6iukfNnuEK"
    table_id = "tblenVnxR8q8iTGSk"
    api_token = os.getenv("SALES_BASE_KEY")

    # Set up the request URL and headers
    url = f"https://api.airtable.com/v0/{base_id}/{table_id}/{record_id}/comments"
    headers = {
        "Authorization": f"Bearer {api_token}"
    }
    params = {
        "pageSize": 25
    }

    # Make the API request
    response = requests.get(url, headers=headers, params=params)

    # Check if the request was successful
    if response.status_code == 200:
        comments_data = response.json()
        
        # Count comments by salespeople
        salesperson_comment_count = 0
        if comments_data.get('comments'):
            for comment in comments_data['comments']:
                if comment.get('author') and comment['author'].get('name') in salespeople_to_check:
                    salesperson_comment_count += 1
        
        return salesperson_comment_count
    else:
        # Return 0 if there was an error retrieving comments
        print(f"Error fetching comments for record {record_id}: {response.status_code} - {response.text}")
        return 0

if __name__ == "__main__":
    start_date_str = "2025-05-06"
    end_date_str = "2025-05-12"
    
    csv_file_path = "weekly_review_data/Leads-Reporting Export.csv"
    salesperson_column_name = "Salesperson" 
    record_id_column_name = "ID"
    date_added_column_name = "Date Added"
    close_status_column_name = "Close Status"

    try:
        df = pd.read_csv(csv_file_path)
    except FileNotFoundError:
        print(f"Error: The file {csv_file_path} was not found.")
        exit()
    
    required_columns = [record_id_column_name, salesperson_column_name, date_added_column_name, close_status_column_name]
    for col in required_columns:
        if col not in df.columns:
            print(f"Error: Required column '{col}' not found in the CSV.")
            exit()

    try:
        df[date_added_column_name] = pd.to_datetime(df[date_added_column_name], format="%m/%d/%Y %I:%M%p")
    except ValueError as e:
        print(f"Error converting '{date_added_column_name}' to datetime: {e}")
        print(f"Please ensure the '{date_added_column_name}' column is in 'MM/DD/YYYY HH:MMam/pm' format.")
        exit()
        
    start_date = pd.to_datetime(start_date_str)
    end_date = pd.to_datetime(end_date_str)
    
    df[close_status_column_name] = df[close_status_column_name].astype(str).fillna('')
    filtered_df = df[
        (df[date_added_column_name] >= start_date) & 
        (df[date_added_column_name] <= end_date) & 
        ~df[close_status_column_name].str.contains("Disqualified", case=False, na=False)
    ].copy()

    if filtered_df.empty:
        print(f"No leads found for the period {start_date_str} to {end_date_str} after filtering.")
        exit()

    commenting_salespeople = ["Kamaal Sherrod", "Hussam Olabi"]
    sales_stats = []

    filtered_df[salesperson_column_name] = filtered_df[salesperson_column_name].astype(str).fillna('Unknown')
    grouped_by_salesperson = filtered_df.groupby(salesperson_column_name)

    print(f"Processing leads for the period: {start_date_str} to {end_date_str}")
    print(f"Counting comments made by: {', '.join(commenting_salespeople)}")

    for sp_name, group_df in grouped_by_salesperson:
        record_ids = group_df[record_id_column_name].tolist()
        num_total_leads_for_sp = len(record_ids)
        
        if num_total_leads_for_sp == 0:
            continue
            
        num_having_at_least_one_comment_for_sp = 0
        
        # print(f"\\nProcessing leads for {salesperson_column_name}: {sp_name}...") # Optional: uncomment for more verbose logging per salesperson
        for record_id in record_ids:
            num_comments = count_salesperson_comments(record_id, salespeople_to_check=commenting_salespeople)
            if num_comments > 0:
                num_having_at_least_one_comment_for_sp += 1
        
        percentage = (num_having_at_least_one_comment_for_sp / num_total_leads_for_sp) * 100 if num_total_leads_for_sp > 0 else 0
        
        sales_stats.append({
            salesperson_column_name: sp_name,
            "Leads with >0 Comments by Target SPs": num_having_at_least_one_comment_for_sp,
            "Total Assigned Leads": num_total_leads_for_sp,
            "Comment Engagement Rate (%)": percentage
        })

    if not sales_stats:
        print("No data to report after processing all groups.")
    else:
        results_df = pd.DataFrame(sales_stats)
        results_df["Comment Engagement Rate (%)"] = results_df["Comment Engagement Rate (%)"].map('{:.2f}'.format) # Display as string with 2 decimal places
        
        print("\\n--- Salesperson Comment Engagement Report ---")
        print(results_df.to_string(index=False))

        # To save to CSV:
        output_filename = f"weekly_outputs/salesperson_comment_report_{start_date_str}_to_{end_date_str}.csv"
        results_df.to_csv(output_filename, index=False)
        print(f"\\nReport saved to {output_filename}")
