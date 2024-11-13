import pandas as pd

def first_year_cancels(salespeople=[]):
    # Read the CSV
    df = pd.read_csv("data/Starts.csv")
    
    # Convert dates to datetime
    df['Start Date'] = pd.to_datetime(df['Start Date'])
    df['Cancel Date'] = pd.to_datetime(df['Cancel Date'], errors='coerce')
    
    # Filter for cancelled accounts after 1/1/22
    df = df[df['Cancel Date'].notna() & 
            (df['Cancel Date'] >= '2022-01-01')]
    
    # Calculate time difference and filter for <= 1 year
    df['Cancel Time'] = (df['Cancel Date'] - df['Start Date']).dt.total_seconds() / (365.25 * 24 * 60 * 60)
    df = df[df['Cancel Time'] <= 1]
    
    # Filter for specified salespeople if list provided
    if salespeople:
        df = df[df['Salesperson'].isin(salespeople)]
    
    # Create pivot table by month and salesperson
    pivot_df = pd.pivot_table(
        df,
        index=df['Cancel Date'].dt.strftime('%Y-%m'),
        columns='Salesperson', 
        values='LocationID',
        aggfunc='count',
        fill_value=0
    )
    
    # Sort index chronologically
    pivot_df = pivot_df.reindex(sorted(pivot_df.index))
    
    # Rename index to desired format (Jun 24)
    pivot_df.index = pd.to_datetime(pivot_df.index + '-01').strftime('%b %y')
    
    # Save to CSV
    pivot_df.to_csv('outputs/First Year Cancels.csv')
    
    return pivot_df