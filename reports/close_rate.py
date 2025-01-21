import pandas as pd

def close_rate(salespeople=[], exclude_channels=[]):
    # Read the CSV file
    df = pd.read_csv('data/Leads-Reporting Export.csv')
    
    # Convert Close Date to datetime
    df['Close Date'] = pd.to_datetime(df['Close Date'], errors='coerce')
    
    # Add Month column
    df['Month'] = df['Close Date'].dt.strftime('%Y-%m')
    
    # Initialize results list
    results = []
    
    # If no salespeople specified, use all unique salespeople
    if not salespeople:
        salespeople = df['Salesperson'].unique()
    
    # Calculate metrics for each salesperson and month
    for sp in salespeople:
        sp_data = df[df['Salesperson'] == sp]
        sp_data = sp_data[~sp_data['Lead Type'].isin(exclude_channels)]

        # Group by month
        for month in sorted(sp_data['Month'].dropna().unique()):
            month_data = sp_data[sp_data['Month'] == month]
            
            # Calculate each metric
            lost = month_data['Close Status'].str.contains('Lost', na=False).sum()
            won_recurring = (month_data['Close Status'] == 'Won: Recurring').sum()
            won_one_time = (month_data['Close Status'] == 'Won: One Time').sum()
            disqualified = month_data['Close Status'].str.contains('Disqualified', na=False).sum()
            open_leads = (month_data['Close Status'].isna()).sum()
            
            # Calculate close rate
            denominator = lost + won_recurring + won_one_time
            close_rate = (won_recurring / denominator * 100) if denominator > 0 else 0
            
            # Calculate disqualify rate
            total_closed = denominator + disqualified
            disqualify_rate = (disqualified / total_closed * 100) if total_closed > 0 else 0
            
            # Create row for results
            results.append({
                'Salesperson': sp,
                'Month': month,
                'Lost': lost,
                'Won: Recurring': won_recurring,
                'Won: One Time': won_one_time,
                'Disqualified': disqualified,
                'Open': open_leads,
                'Close Rate': f'{close_rate:.2f}%',
                'Disqualify Rate': f'{disqualify_rate:.2f}%'
            })
    
    # Convert results to DataFrame and save
    results_df = pd.DataFrame(results)
    results_df.to_csv('outputs/Close Rate.csv', index=False)
    
    return results_df