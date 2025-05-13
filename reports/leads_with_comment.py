import pandas as pd
import requests
import os
from dotenv import load_dotenv

load_dotenv()

def count_salesperson_comments(record_id, salespeople=["Kamaal Sherrod", "Hussam Olabi"]):
    """
    Counts the number of comments made by salespeople for a given record.
    
    Args:
        record_id (str): The Airtable record ID to check for comments
        
    Returns:
        int: The number of comments made by salespeople
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
                if comment.get('author') and comment['author'].get('name') in salespeople:
                    salesperson_comment_count += 1
        
        return salesperson_comment_count
    else:
        # Return 0 if there was an error retrieving comments
        return 0

if __name__ == "__main__":
    start_date = "2025-05-06"
    end_date = "2025-05-12"
    df = pd.read_csv("weekly_review_data/Leads-Reporting Export.csv")
    # cast Date Added to datetime
    df["Date Added"] = pd.to_datetime(df["Date Added"], format="%m/%d/%Y %I:%M%p")
    # get all the records where date added is between start_date and end_date AND Close Status does not contain the substring "Disqualified"
    df = df[(df["Date Added"] >= start_date) & (df["Date Added"] <= end_date) & ~df["Close Status"].str.contains("Disqualified")]
    # get the record ids
    record_ids = df["ID"].tolist()
    # count the comments for each record
    num_having_at_least_one_comment = 0
    num_total_leads = len(record_ids)
    for record_id in record_ids:
        num_comments = count_salesperson_comments(record_id)
        print(f"Record ID: {record_id}, Number of comments: {num_comments}")
        if num_comments > 0:
            num_having_at_least_one_comment += 1
    percentage_having_at_least_one_comment = num_having_at_least_one_comment / num_total_leads
    print(f"Number of leads with at least one comment: {num_having_at_least_one_comment} out of {num_total_leads} ({percentage_having_at_least_one_comment:.2%})")