import pandas as pd

def tech_leads_report():
    # Read the CSV file
    df = pd.read_csv("data/Leads-Tech Leads.csv")
    
    # Convert Date Added to datetime
    df['Date Added'] = pd.to_datetime(df['Date Added'], format='%m/%d/%Y %I:%M%p')
    
    # Create Month-Year column
    df['Month-Year'] = df['Date Added'].dt.strftime('%Y-%m')
    
    # Create pivot table
    pivot = pd.pivot_table(
        df,
        values='Customer Full Name',  # Count based on customer names
        index='Month-Year',
        columns='Tech Name',
        aggfunc='count',
        fill_value=0
    )
    
    # Save to CSV
    pivot.to_csv("outputs/Tech Leads.csv")