import pandas as pd

def first_year_cancels(salespeople=[]):
    # Read the CSV
    df = pd.read_excel("data/Starts.xls")
    
    # Convert dates to datetime
    df['Start Date'] = pd.to_datetime(df['Start Date'])
    df['Cancel Date'] = pd.to_datetime(df['Cancel Date'], errors='coerce')
    
    # Get all months since 2022-01-01
    months = pd.date_range('2024-03-01', df['Cancel Date'].max(), freq='MS')
    
    results = []
    for month_start in months:
        month_end = month_start + pd.offsets.MonthEnd(0)
        
        # Get accounts that are less than 1 year old at start of month
        active_start = df[
            (df['Start Date'] <= month_start) &
            ((df['Cancel Date'].isna()) | (df['Cancel Date'] > month_start))
        ].copy()
        
        active_start['Tenure_Days'] = (month_start - active_start['Start Date']).dt.days
        first_year_accounts = active_start[active_start['Tenure_Days'] <= 365]
        
        # Group by salesperson at start of month
        initial_counts = first_year_accounts.groupby('Salesperson').size()
        
        # Check which accounts survived to end of month
        first_year_accounts['Still_Active'] = (
            (first_year_accounts['Cancel Date'].isna()) | 
            (first_year_accounts['Cancel Date'] > month_end)
        )
        final_counts = first_year_accounts[first_year_accounts['Still_Active']].groupby('Salesperson').size()
        
        # Calculate cancel rates and format with counts
        cancel_rates = (1 - final_counts.div(initial_counts)) * 100
        cancel_rates = cancel_rates.round(2).fillna(0)
        
        # Format string with percentage and counts
        formatted_rates = {}
        for salesperson in cancel_rates.index:
            if salesperson in initial_counts:
                initial = initial_counts[salesperson]
                final = final_counts.get(salesperson, 0)
                rate = cancel_rates[salesperson]
                formatted_rates[salesperson] = f"{rate:.2f}% ({final}/{initial})"
            else:
                formatted_rates[salesperson] = "0.00% (0/0)"
                
        # Store results
        formatted_series = pd.Series(formatted_rates)
        formatted_series.name = month_start.strftime('%b %y')
        results.append(formatted_series)
    
    # Create pivot table
    pivot_df = pd.concat(results, axis=1).T
    
    # Filter for specified salespeople if list provided
    if salespeople:
        pivot_df = pivot_df[salespeople]
    
    # Save to CSV
    pivot_df.to_csv('outputs/First Year Cancels.csv')
    
    return pivot_df