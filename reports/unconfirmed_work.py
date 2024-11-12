import pandas as pd

def unconfirmed_work():
    # Read the CSV
    df = pd.read_csv("data/Work Completion - Stop Data at Start of Day.csv")
    
    # Convert Work Date to datetime
    df['Work Date'] = pd.to_datetime(df['Work Date'])
    
    # Create Year-Month column
    df['Year-Month'] = df['Work Date'].dt.strftime('%Y-%m')
    
    # Create the summary report
    report = pd.DataFrame({
        'Yellow_Count': df[df['Color'] == 'Yellow'].groupby('Year-Month').size(),
        'Other_Colors': df[df['Color'] != 'Yellow'].groupby('Year-Month').size()
    }).fillna(0)
    
    # Calculate total and percentage
    report['Total'] = report['Yellow_Count'] + report['Other_Colors']
    report['Percent_Unconfirmed'] = (report['Yellow_Count'] / report['Total'] * 100).round(2).astype(str) + '%'
    
    # Sort by Year-Month
    report = report.sort_index()
    
    # Save to CSV
    report.to_csv('outputs/Unconfirmed.csv')