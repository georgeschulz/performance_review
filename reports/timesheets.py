import pandas as pd
import os

def timesheets_report():
    df = pd.read_excel("data/Timesheets.xlsx")
    
    # Create full name column
    df['Full Name'] = df['First Name'] + ' ' + df['Last Name']
    
    # Group by full name and sum the hours
    summary = df.groupby('Full Name').agg({
        'Regular': 'sum',
        'OT': 'sum'
    }).reset_index()
    
    # Replace NaN with 0 for overtime
    summary['OT'] = summary['OT'].fillna(0)
    
    # Calculate total hours
    summary['Total Hours'] = summary['Regular'] + summary['OT']
    
    # Add total row
    total_row = pd.DataFrame({
        'Full Name': ['Total'],
        'Regular': [summary['Regular'].sum()],
        'OT': [summary['OT'].sum()],
        'Total Hours': [summary['Total Hours'].sum()]
    })
    summary = pd.concat([summary, total_row], ignore_index=True)
    
    # Ensure output directory exists
    os.makedirs('outputs', exist_ok=True)
    
    # Save to CSV
    summary.to_csv('outputs/Hours.csv', index=False)