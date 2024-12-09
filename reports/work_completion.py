import pandas as pd

def work_completion_report(excluded_techs=[]):
    df = pd.read_csv("data/Work Completion - Stop Data at Start of Day.csv")
    # remove any columns where location Code is NaN
    df = df.dropna(subset=['Location Code'])

    # Convert 'Work Date' to datetime
    df['Work Date'] = pd.to_datetime(df['Work Date'])

    # Create a new column for month-year
    df['Month-Year'] = df['Work Date'].dt.to_period('M')

    # Filter out the excluded techs
    df = df[~df['Tech 1'].isin(excluded_techs)]

    # Count and print "Rescheduled - No One's Fault" rows before filtering
    no_fault_count = len(df[df['Status'] == "Rescheduled - No One's Fault"])
    print(f"Number of 'Rescheduled - No One's Fault' records: {no_fault_count}")
    
    # Filter out "Rescheduled - No One's Fault" rows
    df = df[df['Status'] != "Rescheduled - No One's Fault"]

    # Create a pivot table
    pivot = pd.pivot_table(
        df,
        values='Status',
        index='Month-Year',
        columns='Tech 1',
        aggfunc=lambda x: (x == 'Done').mean(),
        fill_value=0
    )

    # Sort the index chronologically
    pivot = pivot.sort_index()

    # Format cells as percentages
    # Using .map() instead of deprecated .applymap()
    pivot = pivot.map(lambda x: f"{x:.2%}")

    # Save to CSV
    pivot.to_csv("outputs/Work Completion.csv")
