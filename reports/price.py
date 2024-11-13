import pandas as pd
import os

def price_report(salespeople=[]):
    # Read the Excel file
    df = pd.read_excel("data/Price Analysis.xls")
    
    # Convert Start Date to datetime if it isn't already
    df['Start Date'] = pd.to_datetime(df['Start Date'])
    
    # Create Year-Month column from Start Date
    df['Year-Month'] = df['Start Date'].dt.to_period('M')
    
    # Convert First Year ACV to numeric, removing '$' if present
    df['First Year ACV'] = df['First Year ACV'].astype(str).str.replace('$', '').astype(float)

    # Filter by salespeople if provided
    if salespeople:
        df = df[df['Salesperson'].isin(salespeople)]
    
    # Create pivot table
    pivot_df = df.pivot_table(
        values='First Year ACV',
        index='Salesperson',
        columns='Year-Month',
        aggfunc='mean'
    )
    
    # Format values as currency
    pivot_df = pivot_df.map(lambda x: f'${x:,.2f}')
    
    # Ensure directory exists
    os.makedirs('outputs', exist_ok=True)
    
    # Save to CSV
    pivot_df.to_csv('outputs/Price.csv')