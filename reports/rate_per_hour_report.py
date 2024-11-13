import pandas as pd

def rate_per_hour_report(custom_joins=[]):
    # Read the production data
    production_df = pd.read_csv("outputs/Production.csv")
    hours_df = pd.read_csv("outputs/Hours.csv")
    
    # Convert Production from string ($x,xxx.xx) to float
    production_df['Total Production'] = production_df['Total Production'].str.replace('$', '').str.replace(',', '').astype(float)
    
    # Apply custom joins to handle name mismatches
    for prod_name, hours_name in custom_joins:
        hours_df.loc[hours_df['Full Name'] == hours_name, 'Full Name'] = prod_name
    
    # Merge production and hours data
    combined_df = production_df.merge(
        hours_df,
        left_on='Technician',
        right_on='Full Name',
        how='inner'
    )
    
    # Calculate rate per hour
    combined_df['Rate Per Hour'] = combined_df['Total Production'] / combined_df['Total Hours']
    
    # Select and rename columns
    result = combined_df[[
        'Technician',
        'Total Production',
        'Regular',
        'OT',
        'Total Hours',
        'Rate Per Hour'
    ]]
    
    # Sort by rate per hour descending
    result = result.sort_values('Rate Per Hour', ascending=False)
    
    # Format currency columns
    result['Total Production'] = result['Total Production'].apply(lambda x: f"${x:,.2f}")
    result['Rate Per Hour'] = result['Rate Per Hour'].apply(lambda x: f"${x:,.2f}")
    
    # Save to CSV
    result.to_csv('outputs/Rate Per Hour.csv', index=False)