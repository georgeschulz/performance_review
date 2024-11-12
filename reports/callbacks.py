import pandas as pd

def callbacks_report():
    pd.options.mode.chained_assignment = None  # Turn off SettingWithCopyWarning
    current_period = pd.read_csv('data/Monthly Invoice Report.csv')
    historical = pd.read_csv('data/Historical Invoice Report.csv')

    # Get callbacks from current period
    callbacks = current_period[current_period['Invoice Type'] == 'Call Back']
    
    # Convert Work Date to datetime for proper comparison
    callbacks['Work Date'] = pd.to_datetime(callbacks['Work Date'])
    historical['Work Date'] = pd.to_datetime(historical['Work Date'])
    
    # Initialize Original Tech column with "Unknown"
    callbacks['Original Tech'] = "Unknown"
    
    # Find original tech for each callback
    for idx, callback in callbacks.iterrows():
        # Get historical records for this account before the callback date
        account_history = historical[
            (historical['Account'] == callback['Account']) & 
            (historical['Work Date'] < callback['Work Date'])
        ]
        
        if not account_history.empty:
            # Get the most recent service date and tech
            latest_service = account_history.loc[account_history['Work Date'].idxmax()]
            callbacks.at[idx, 'Original Tech'] = latest_service['Tech']
    
    # Get total invoices per tech from current period
    tech_invoice_counts = current_period['Tech'].value_counts()
    
    # Count callbacks by original tech
    tech_callback_counts = callbacks['Original Tech'].value_counts()
    
    # Create a DataFrame with combined metrics
    tech_metrics = pd.DataFrame({
        'Total Invoices': tech_invoice_counts,
        'Callbacks': tech_callback_counts
    }).fillna(0)  # Fill NaN with 0 for techs with no callbacks
    
    # Calculate callback percentage as numeric first
    tech_metrics['Callback Rate'] = (tech_metrics['Callbacks'] / tech_metrics['Total Invoices'] * 100).round(2)
    
    # Sort by callback rate descending
    tech_metrics = tech_metrics.sort_values('Callback Rate', ascending=False)
    
    # Add percentage symbol after sorting
    tech_metrics['Callback Rate'] = tech_metrics['Callback Rate'].astype(str) + '%'
    
    # save this to outputs/Callback Report
    tech_metrics.to_csv('outputs/Callback Report.csv')