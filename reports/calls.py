import pandas as pd

def calls_report(user_mappings=[]):
    # CLEAN MALFORMED CSV FILE
    # Read file lines and handle first line formatting
    with open("data/Calls.csv", "r") as file:
        lines = file.readlines()

    # Remove first line if blank    
    if lines[0].strip() == "":
        lines = lines[1:]
        
    # Add comma to first line if needed
    if not lines[0].rstrip().endswith(","):
        lines[0] = lines[0].rstrip() + ",\n"
        
    # Write back to file
    with open("data/Calls.csv", "w") as file:
        file.writelines(lines)

    # Read the CSV file with proper column names
    df = pd.read_csv("data/Calls.csv")
    
    # Create mapping dictionary for extensions to names
    extension_to_name = {ext: name for ext, name in user_mappings}
    
    # Add new column for person name based on From/To
    def get_person(row):
        if row['From'] in extension_to_name:
            return extension_to_name[row['From']]
        elif row['To'] in extension_to_name:
            return extension_to_name[row['To']]
        return 'Other'
        
    df['Person'] = df.apply(get_person, axis=1)

    # Convert Call Date to datetime
    df['Call Date'] = pd.to_datetime(df['Call Date'])
    
    # Create date range from min to max date
    date_range = pd.date_range(start=df['Call Date'].min(), end=df['Call Date'].max(), freq='D')
    
    # Count calls per person per day
    daily_counts = df[df['Person'] != 'Other'].pivot_table(
        index='Call Date',
        columns='Person',
        values='Call Time',
        aggfunc='count',
        fill_value=0
    )
    
    # Reindex with full date range to include missing dates
    daily_counts = daily_counts.reindex(date_range, fill_value=0)
    
    # Calculate averages excluding weekends and 0s
    averages = pd.Series(index=daily_counts.columns)
    for column in daily_counts.columns:
        non_zero_values = daily_counts[column][daily_counts[column] > 0]
        averages[column] = round(non_zero_values.mean(), 1)
    
    # Create empty row
    empty_row = pd.Series(index=daily_counts.columns, data=[''] * len(daily_counts.columns))
    
    # Add empty row and averages row
    daily_counts.loc[''] = empty_row
    daily_counts.loc['Average'] = averages
    
    # Save to CSV
    daily_counts.to_csv('outputs/Daily_Calls.csv')