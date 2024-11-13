import pandas as pd
from datetime import datetime

def retention_report():
    df = pd.read_csv("data/Starts.csv")
    df['Start Date'] = pd.to_datetime(df['Start Date'])
    df['Cancel Date'] = pd.to_datetime(df['Cancel Date'], errors='coerce')

    # Define the start and end dates for the report
    start_date = pd.to_datetime('2022-01-01')
    end_date = pd.to_datetime('2024-11-01')  # Adjust as per your data's availability

    months = pd.date_range(start_date, end_date, freq='MS')  # Generate a list of month starts

    results = []

    for month_start in months:
        # Define the end of the month
        month_end = month_start + pd.offsets.MonthEnd(0)

        # Customers active at the start of the month
        active_start = df[
            (df['Start Date'] <= month_start) &
            ((df['Cancel Date'].isna()) | (df['Cancel Date'] > month_start))
        ].copy()

        # Compute tenure in days at the start of the month
        active_start['Tenure_Days_Start'] = (month_start - active_start['Start Date']).dt.days

        # Categorize customers at the start
        initial_1yr_customers = active_start[active_start['Tenure_Days_Start'] <= 365]
        initial_long_customers = active_start[active_start['Tenure_Days_Start'] > 365]

        initial_1yr = len(initial_1yr_customers)
        initial_long = len(initial_long_customers)

        # Check if initial customers are still active at the end of the month
        initial_1yr_customers['Still_Active'] = ((initial_1yr_customers['Cancel Date'].isna()) | (initial_1yr_customers['Cancel Date'] > month_end))
        initial_long_customers['Still_Active'] = ((initial_long_customers['Cancel Date'].isna()) | (initial_long_customers['Cancel Date'] > month_end))

        final_1yr = initial_1yr_customers['Still_Active'].sum()
        final_long = initial_long_customers['Still_Active'].sum()

        # New customers added during the period
        new_customers = df[
            (df['Start Date'] >= month_start) &
            (df['Start Date'] <= month_end)
        ].copy()
        new_customers_count = len(new_customers)

        # Calculate retention rates
        if initial_1yr > 0:
            first_year_retention = (final_1yr / initial_1yr) * 100
        else:
            first_year_retention = 0

        if initial_long > 0:
            long_time_retention = (final_long / initial_long) * 100
        else:
            long_time_retention = 0

        initial_total = initial_1yr + initial_long
        final_total = final_1yr + final_long

        if initial_total > 0:
            combined_retention = (final_total / initial_total) * 100
        else:
            combined_retention = 0

        # Compile results for the month
        month_result = {
            'Month': month_start.strftime('%b %y'),
            'Initial Customers <1yr': initial_1yr,
            'Initial Customers >1yr': initial_long,
            'Final Customers <1yr': final_1yr,
            'Final Customers >1yr': final_long,
            'New Customers Added': new_customers_count,
            'First Year Retention Rate': round(first_year_retention, 2),
            'Long Time Customer Retention Rate': round(long_time_retention, 2),
            'Combined Retention Rate': round(combined_retention, 2)
        }

        results.append(month_result)

    # Create a DataFrame from the results
    retention_df = pd.DataFrame(results)

    # Now, calculate the trailing 12-month retention rates
    t12_results = []

    for i in range(11, len(months)):
        month_start = months[i]
        month_end = month_start + pd.offsets.MonthEnd(0)
        period_start = months[i - 11]  # Start of the 12-month period
        period_end = month_end  # End of the 12-month period

        # Customers active at the start of the 12-month period
        active_start = df[
            (df['Start Date'] <= period_start) &
            ((df['Cancel Date'].isna()) | (df['Cancel Date'] > period_start))
        ].copy()

        # Compute tenure in days at the period start
        active_start['Tenure_Days_Start'] = (period_start - active_start['Start Date']).dt.days

        # Categorize customers at the period start
        initial_1yr_customers = active_start[active_start['Tenure_Days_Start'] <= 365]
        initial_long_customers = active_start[active_start['Tenure_Days_Start'] > 365]

        initial_1yr = len(initial_1yr_customers)
        initial_long = len(initial_long_customers)

        # Check if initial customers are still active at the end of the 12-month period
        initial_1yr_customers['Still_Active'] = ((initial_1yr_customers['Cancel Date'].isna()) | (initial_1yr_customers['Cancel Date'] > period_end))
        initial_long_customers['Still_Active'] = ((initial_long_customers['Cancel Date'].isna()) | (initial_long_customers['Cancel Date'] > period_end))

        final_1yr = initial_1yr_customers['Still_Active'].sum()
        final_long = initial_long_customers['Still_Active'].sum()

        # Calculate retention rates
        if initial_1yr > 0:
            first_year_retention = (final_1yr / initial_1yr) * 100
        else:
            first_year_retention = 0

        if initial_long > 0:
            long_time_retention = (final_long / initial_long) * 100
        else:
            long_time_retention = 0

        initial_total = initial_1yr + initial_long
        final_total = final_1yr + final_long

        if initial_total > 0:
            combined_retention = (final_total / initial_total) * 100
        else:
            combined_retention = 0

        # Compile results for the trailing 12-month period
        t12_result = {
            'Month': month_start.strftime('%b %y'),
            'Trailing 12-Month Initial Customers <1yr': initial_1yr,
            'Trailing 12-Month Initial Customers >1yr': initial_long,
            'Trailing 12-Month Final Customers <1yr': final_1yr,
            'Trailing 12-Month Final Customers >1yr': final_long,
            'First Year Retention Rate (T12M)': round(first_year_retention, 2),
            'Long Time Customer Retention Rate (T12M)': round(long_time_retention, 2),
            'Combined Retention Rate (T12M)': round(combined_retention, 2)
        }

        t12_results.append(t12_result)

    # Create a DataFrame from the trailing 12-month results
    t12_retention_df = pd.DataFrame(t12_results)

    # Write both DataFrames to CSV, appending the trailing 12-month data below the monthly data
    with open('outputs/Retention.csv', 'w') as f:
        retention_df.to_csv(f, index=False)
        f.write('\n')  # Add an empty line to separate the tables
        t12_retention_df.to_csv(f, index=False)

