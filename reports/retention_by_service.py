import pandas as pd
from datetime import datetime

def retention_by_service_report(start_date, end_date):
    # Read and prepare data
    df = pd.read_excel("data/Starts.xls")
    df['Start Date'] = pd.to_datetime(df['Start Date'])
    df['Cancel Date'] = pd.to_datetime(df['Cancel Date'], errors='coerce')
    
    start_date = pd.to_datetime(start_date)
    end_date = pd.to_datetime(end_date)
    
    # Get unique service types
    service_types = df['Service Type'].unique()
    results = []
    
    for service_type in service_types:
        # Filter for current service type
        service_df = df[df['Service Type'] == service_type]
        
        # Customers active at the start of the period
        active_start = service_df[
            (service_df['Start Date'] <= start_date) &
            ((service_df['Cancel Date'].isna()) | (service_df['Cancel Date'] > start_date))
        ].copy()
        
        # Calculate tenure and categorize customers
        active_start['Tenure_Days_Start'] = (start_date - active_start['Start Date']).dt.days
        initial_1yr_customers = active_start[active_start['Tenure_Days_Start'] <= 365]
        initial_long_customers = active_start[active_start['Tenure_Days_Start'] > 365]
        
        initial_1yr = len(initial_1yr_customers)
        initial_long = len(initial_long_customers)
        
        # Check end-of-period status
        initial_1yr_customers['Still_Active'] = ((initial_1yr_customers['Cancel Date'].isna()) | 
                                               (initial_1yr_customers['Cancel Date'] > end_date))
        initial_long_customers['Still_Active'] = ((initial_long_customers['Cancel Date'].isna()) | 
                                                (initial_long_customers['Cancel Date'] > end_date))
        
        final_1yr = initial_1yr_customers['Still_Active'].sum()
        final_long = initial_long_customers['Still_Active'].sum()
        
        # New customers in this period
        new_customers = service_df[
            (service_df['Start Date'] >= start_date) &
            (service_df['Start Date'] <= end_date)
        ]
        new_customers_count = len(new_customers)
        
        # Calculate retention rates
        first_year_retention = (final_1yr / initial_1yr * 100) if initial_1yr > 0 else 0
        long_time_retention = (final_long / initial_long * 100) if initial_long > 0 else 0
        
        initial_total = initial_1yr + initial_long
        final_total = final_1yr + final_long
        combined_retention = (final_total / initial_total * 100) if initial_total > 0 else 0
        
        # Compile results
        result = {
            'Service Type': service_type,
            'Initial Customers <1yr': initial_1yr,
            'Initial Customers >1yr': initial_long,
            'Final Customers <1yr': final_1yr,
            'Final Customers >1yr': final_long,
            'New Customers Added': new_customers_count,
            'First Year Retention Rate': round(first_year_retention, 2),
            'Long Time Customer Retention Rate': round(long_time_retention, 2),
            'Combined Retention Rate': round(combined_retention, 2)
        }
        
        results.append(result)
    
    # Create DataFrame and sort by Service Type
    retention_df = pd.DataFrame(results)
    retention_df = retention_df.sort_values('Service Type')
    
    # Save to CSV
    retention_df.to_csv('outputs/Service_Type_Retention.csv', index=False)

    return retention_df

print(retention_by_service_report("2023-01-01", "2024-12-31"))