import pandas as pd

def rate_per_hour_report(custom_joins=[], exclude_techs=[]):
    # Read the raw invoice data and hours data
    df = pd.read_csv("data/Monthly Invoice Report.csv")
    hours_df = pd.read_csv("outputs/Hours.csv")
    
    # Combine First Name and Last Name into Technician
    df['Technician'] = df['First Name'] + ' ' + df['Last Name']
    
    
    # Calculate production metrics
    daily_production = df.groupby(['Tech', 'Work Date'])['Total'].sum().reset_index()
    total_orders = df.groupby('Tech').size().reset_index(name='Total Orders')
    total_stops = df.groupby('Tech').agg({
        'Account': lambda x: len(pd.unique(x.astype(str) + df.loc[x.index, 'Work Date'].astype(str)))
    }).reset_index()
    total_stops.columns = ['Tech', 'Total Stops']
    
    # Calculate tech totals
    tech_totals = daily_production.groupby('Tech').agg({
        'Total': ['sum', lambda x: x.count()]
    }).reset_index()
    tech_totals.columns = ['Technician', 'Total Production', 'Days Worked']
    
    # Merge in orders and stops
    tech_totals = tech_totals.merge(total_orders, left_on='Technician', right_on='Tech').drop('Tech', axis=1)
    tech_totals = tech_totals.merge(total_stops, left_on='Technician', right_on='Tech').drop('Tech', axis=1)
    
    # Calculate averages
    tech_totals['Average Production Per Day'] = tech_totals['Total Production'] / tech_totals['Days Worked']
    tech_totals['Average Orders Per Day'] = tech_totals['Total Orders'] / tech_totals['Days Worked']
    tech_totals['Average Stops Per Day'] = tech_totals['Total Stops'] / tech_totals['Days Worked']
    
    # Apply custom joins to handle name mismatches
    for prod_name, hours_name in custom_joins:
        hours_df.loc[hours_df['Full Name'] == hours_name, 'Full Name'] = prod_name
    
    # Merge production and hours data
    combined_df = tech_totals.merge(
        hours_df,
        left_on='Technician',
        right_on='Full Name',
        how='inner'
    )

    # Filter out excluded techs
    combined_df = combined_df[~combined_df['Technician'].isin(exclude_techs)]
    
    # Calculate rate per hour
    combined_df['Rate Per Hour'] = combined_df['Total Production'] / combined_df['Total Hours']
    
    # Select and rename columns for final output
    result = combined_df[[
        'Technician',
        'Total Production',
        'Days Worked',
        'Total Orders',
        'Total Stops',
        'Average Production Per Day',
        'Average Orders Per Day',
        'Average Stops Per Day',
        'Regular',
        'OT',
        'Total Hours',
        'Rate Per Hour'
    ]]
    
    # Sort by rate per hour descending
    result = result.sort_values('Rate Per Hour', ascending=False)
    
    # Format currency columns
    result['Total Production'] = result['Total Production'].apply(lambda x: f"${x:,.2f}")
    result['Average Production Per Day'] = result['Average Production Per Day'].apply(lambda x: f"${x:,.2f}")
    result['Rate Per Hour'] = result['Rate Per Hour'].apply(lambda x: f"${x:,.2f}")
    
    # Round average metrics to 2 decimal places
    result['Average Orders Per Day'] = result['Average Orders Per Day'].round(1)
    result['Average Stops Per Day'] = result['Average Stops Per Day'].round(1)
    
    # Calculate summary statistics
    summary_df = pd.DataFrame({
        'Technician': ['TOTALS'],
        'Total Production': [combined_df['Total Production'].sum()],
        'Total Hours': [combined_df['Total Hours'].sum()],
        'Regular': [combined_df['Regular'].sum()],
        'OT': [combined_df['OT'].sum()],
        'Rate Per Hour': [combined_df['Total Production'].sum() / combined_df['Total Hours'].sum()]
    })
    
    # Format currency columns in summary
    summary_df['Total Production'] = summary_df['Total Production'].apply(lambda x: f"${x:,.2f}")
    summary_df['Rate Per Hour'] = summary_df['Rate Per Hour'].apply(lambda x: f"${x:,.2f}")
    
    # Create a separator row with empty values
    separator_df = pd.DataFrame({col: [''] for col in result.columns}, index=[0])
    
    # Combine result, separator and summary
    result = pd.concat([result, separator_df, summary_df], ignore_index=True)
    
    # Save to CSV
    result.to_csv('outputs/Productivity.csv', index=False)