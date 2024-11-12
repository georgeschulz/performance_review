import pandas as pd

def sales_data_report():
    df = pd.read_csv("data/Sales-Performance Review Export.csv")
    
    # Filter rows where 'Closed?' is 'Closed' and 'Period' is not 'Old'
    closed_df = df[(df['Closed?'] == 'Closed') & (df['Period'] != 'Old')]
    
    # Remove the dollar sign and convert 'Sales Value' to float
    closed_df['Sales Value'] = closed_df['Sales Value'].replace('[\$,]', '', regex=True).astype(float)
    
    # Ensure 'First Date in Period' is converted to datetime
    closed_df['First Date in Period'] = pd.to_datetime(closed_df['First Date in Period'], errors='coerce')
    
    # Sort by 'First Date in Period' in ascending order
    closed_df = closed_df.sort_values('First Date in Period', ascending=True)
    
    # Group by 'Salespeople' and 'Period'
    aggregated_df = closed_df.groupby(['Salespeople', 'Period']).agg(
        Recurring_Count=('Type', lambda x: ((x == 'Recurring') & (closed_df.loc[x.index, 'Service'] != 'CHARGEBACK')).sum()),
        One_Time_Count=('Type', lambda x: ((x == 'One Time') & (closed_df.loc[x.index, 'Service'] != 'CHARGEBACK')).sum()),
        Total_Sales_Value=('Sales Value', 'sum')
    ).reset_index()
    
    # Rename columns to more friendly names
    aggregated_df.columns = ['Salesperson', 'Sales Period', 'Recurring Sales Count', 'One-Time Sales Count', 'Total Sales Value']
    
    # Format 'Total Sales Value' as currency
    aggregated_df['Total Sales Value'] = aggregated_df['Total Sales Value'].apply(lambda x: f"${x:,.2f}")
    
    # Sort the aggregated data by 'Sales Period' and 'Salesperson'
    aggregated_df = aggregated_df.sort_values(['Sales Period', 'Salesperson'], ascending=True)
    
    # Save to outputs/Sales Data.csv
    aggregated_df.to_csv("outputs/Sales Data.csv", index=False)